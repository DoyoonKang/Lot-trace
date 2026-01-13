import altair as alt
import streamlit as st
import pandas as pd
import datetime as dt
import re
from pathlib import Path
from openpyxl import load_workbook
import requests
from io import StringIO
from typing import Optional, List, Dict


# =========================
# Page Config (딱 1번만!)
# =========================
st.set_page_config(
    page_title="액상 잉크 Lot 추적 관리",
    page_icon="🧪",
    layout="wide",
)


# =========================
# Google Sheets (Public) Reader
# =========================
@st.cache_data(ttl=60, show_spinner=False)
def read_gsheet_csv(sheet_id: str, sheet_name: str) -> pd.DataFrame:
    base = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
    r = requests.get(base, params={"tqx": "out:csv", "sheet": sheet_name}, timeout=20)
    r.raise_for_status()
    r.encoding = "utf-8"
    return pd.read_csv(StringIO(r.text))


# =========================
# Config
# =========================
DEFAULT_XLSX = "액상잉크_Lot추적관리_FINAL.xlsx"

SHEET_BINDER = "바인더_제조_입고"
SHEET_SINGLE = "단일색_수입검사"
SHEET_SPEC_BINDER = "Spec_Binder"
SHEET_SPEC_SINGLE = "Spec_Single_H&S"
SHEET_BINDER_VISC = "Binder_Visc"
SHEET_BASE_LAB = "기준LAB"

# 업체 반환(반품) 기록용 시트
SHEET_BINDER_RETURN = "바인더_업체반환"

COLOR_CODE = {
    "Black": "B",
    "White": "W",
    "Blue": "U",
    "Green": "G",
    "Yellow": "Y",
    "Red": "R",
    "Pink": "P",
}

# 바인더 입출고(구글시트)
BINDER_SHEET_ID = "1H2fFxnf5AvpSlu-uoZ4NpTv8LYLNwTNAzvlntRQ7FS8"
BINDER_SHEET_HEMA = "HEMA 바인더 입출고 관리대장"
BINDER_SHEET_SIL = "Silicon바인더 입출고 관리대장"


# =========================
# Helpers
# =========================
def _read_excel_from_path(xlsx_path: str) -> Dict[str, pd.DataFrame]:
    def read(name: str) -> pd.DataFrame:
        return pd.read_excel(xlsx_path, sheet_name=name)

    return {
        "binder": read(SHEET_BINDER),
        "single": read(SHEET_SINGLE),
        "spec_binder": read(SHEET_SPEC_BINDER),
        "spec_single": read(SHEET_SPEC_SINGLE),
        "binder_visc": read(SHEET_BINDER_VISC),
        "base_lab": read(SHEET_BASE_LAB),
    }


@st.cache_data(show_spinner=False)
def load_data(xlsx_path: str) -> Dict[str, pd.DataFrame]:
    return _read_excel_from_path(xlsx_path)


def normalize_date(x):
    if pd.isna(x):
        return None
    if isinstance(x, (dt.date, dt.datetime)):
        return x.date() if isinstance(x, dt.datetime) else x
    try:
        return pd.to_datetime(x).date()
    except Exception:
        return None


def delta_e76(lab1, lab2):
    return float(((lab1[0] - lab2[0]) ** 2 + (lab1[1] - lab2[1]) ** 2 + (lab1[2] - lab2[2]) ** 2) ** 0.5)


def extract_delta_e_from_note(note: str):
    if note is None or pd.isna(note):
        return None
    s = str(note)
    m = re.search(r"\[ΔE76=([0-9]+(?:\.[0-9]+)?)\]", s)
    if m:
        try:
            return float(m.group(1))
        except Exception:
            return None
    return None


def get_binder_limits(spec_binder: pd.DataFrame, binder_name: str):
    df = spec_binder[spec_binder["바인더명"] == binder_name].copy()
    visc = df[df["시험항목"].astype(str).str.contains("점도", na=False)]
    uv = df[df["시험항목"].astype(str).str.contains("UV", na=False)]

    visc_lo = float(visc["하한"].dropna().iloc[0]) if len(visc["하한"].dropna()) else None
    visc_hi = float(visc["상한"].dropna().iloc[0]) if len(visc["상한"].dropna()) else None
    uv_hi = float(uv["상한"].dropna().iloc[0]) if len(uv["상한"].dropna()) else None
    rule = (
        df["Lot부여규칙"].dropna().iloc[0]
        if "Lot부여규칙" in df.columns and len(df["Lot부여규칙"].dropna())
        else None
    )
    return visc_lo, visc_hi, uv_hi, rule


def parse_binder_rule_prefix(rule: Optional[str], binder_name_fallback: str):
    if rule:
        m = re.match(r"^([A-Za-z0-9]+)\+YYYYMMDD(-##)?$", str(rule).strip())
        if m:
            prefix = m.group(1)
            has_seq = bool(m.group(2))
            return prefix, has_seq
    prefix = re.sub(r"\W+", "", str(binder_name_fallback))[:6].upper()
    return prefix, True


def infer_binder_type_from_lot(spec_binder: pd.DataFrame, binder_lot: str):
    if not binder_lot:
        return None
    rules = (
        spec_binder[["바인더명", "Lot부여규칙"]]
        .dropna()
        .drop_duplicates()
        .to_dict("records")
    )
    for r in rules:
        rule = str(r["Lot부여규칙"])
        m = re.match(r"^([A-Za-z0-9]+)\+", rule)
        if m:
            prefix = m.group(1)
            if str(binder_lot).startswith(prefix):
                return r["바인더명"]
    return None


def next_seq_for_pattern(existing_lots: pd.Series, prefix: str, date_str: str, sep: str = "-"):
    lots = existing_lots.dropna().astype(str).tolist()
    seqs = []
    for lot in lots:
        if not lot.startswith(prefix + date_str):
            continue
        rest = lot[len(prefix + date_str):]
        if sep and rest.startswith(sep):
            rest = rest[len(sep):]
        m = re.match(r"^(\d+)", rest)
        if m:
            try:
                seqs.append(int(m.group(1)))
            except Exception:
                pass
    return (max(seqs) + 1) if seqs else 1


def generate_binder_lot(prefix: str, mfg_date: dt.date, seq: Optional[int]):
    date_str = mfg_date.strftime("%Y%m%d")
    if seq is None:
        return f"{prefix}{date_str}"
    return f"{prefix}{date_str}-{seq:02d}"


def generate_single_lot(single_df: pd.DataFrame, product_code: str, color_group: str, in_date: dt.date):
    code = (product_code or "").strip()
    color_code = COLOR_CODE.get(color_group)
    if not color_code:
        return None

    if code.startswith("NPL"):
        prefix = "NPL"
    elif code.startswith("PL"):
        prefix = "PL"
    elif code.startswith("SL") or code.startswith("NSL"):
        prefix = "SL"
    else:
        prefix = "PL"

    date_str = in_date.strftime("%y%m%d")
    patt_prefix = f"{prefix}{color_code}{date_str}"
    lots = single_df.get("단일색잉크 Lot", pd.Series(dtype=str)).dropna().astype(str).tolist()

    seqs = []
    for lot in lots:
        if lot.startswith(patt_prefix):
            rest = lot[len(patt_prefix):]
            m = re.match(r"^(\d{2,})", rest)
            if m:
                seqs.append(int(m.group(1)))
    seq = (max(seqs) + 1) if seqs else 1
    return f"{patt_prefix}{seq:02d}"


def judge_range(value, lo, hi):
    if value is None or pd.isna(value):
        return None
    try:
        v = float(value)
    except Exception:
        return None
    if lo is not None and v < float(lo):
        return "부적합"
    if hi is not None and v > float(hi):
        return "부적합"
    return "적합"


def ensure_sheet_with_headers(xlsx_path: str, sheet_name: str, headers: List[str]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        wb.save(xlsx_path)


def append_row_to_sheet(xlsx_path: str, sheet_name: str, row: dict):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    values = [row.get(h, None) for h in headers]
    ws.append(values)
    wb.save(xlsx_path)


def append_rows_to_sheet(xlsx_path: str, sheet_name: str, rows: List[dict]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    for row in rows:
        values = [row.get(h, None) for h in headers]
        ws.append(values)
    wb.save(xlsx_path)


def df_quick_filter(df: pd.DataFrame, text: str, cols: List[str]):
    if not text:
        return df
    t = str(text).strip()
    if not t:
        return df
    mask = False
    for c in cols:
        if c not in df.columns:
            continue
        mask = mask | df[c].astype(str).str.contains(t, case=False, na=False)
    return df[mask]


def sort_df_by_any_date_col(df: pd.DataFrame):
    if df is None or len(df) == 0:
        return df
    candidates = ["일자", "날짜", "입출고일", "입고일", "출고일", "Date", "date"]
    hit = None
    for c in candidates:
        if c in df.columns:
            hit = c
            break
    if hit is None:
        return df
    tmp = df.copy()
    tmp["_sort_date"] = pd.to_datetime(tmp[hit], errors="coerce")
    tmp = tmp.sort_values("_sort_date", ascending=False).drop(columns=["_sort_date"])
    return tmp


# =========================
# UI Header
# =========================
st.title("액상 잉크 Lot 추적 관리 대시보드")
st.caption("상사용(요약→표→그래프) 흐름으로 한눈에 보이도록 구성했습니다.")


# =========================
# Data file selection (Excel)
# =========================
with st.sidebar:
    st.header("데이터 파일")
    xlsx_path_input = st.text_input(
        "엑셀 파일 경로(권장)",
        value=DEFAULT_XLSX,
        help="가능하면 파일 경로 방식으로 사용하시는 것이 가장 안정적입니다."
    )
    uploaded = st.file_uploader("또는 엑셀 업로드(업로드 모드는 세션/다운로드 방식)", type=["xlsx"])

# 업로드 모드: 최초 1회만 tmp에 저장하고, 이후 rerun 때 덮어쓰지 않음
xlsx_path = xlsx_path_input
if uploaded is not None:
    if "_work_xlsx_path" not in st.session_state:
        tmp_path = Path(st.session_state.get("_tmp_xlsx_path", "WORK_LotTrace.xlsx"))
        tmp_path.write_bytes(uploaded.getvalue())
        st.session_state["_work_xlsx_path"] = str(tmp_path)
        st.session_state["_uploaded_bytes"] = uploaded.getvalue()
    xlsx_path = st.session_state["_work_xlsx_path"]

    st.sidebar.info("업로드 모드로 실행 중입니다. 입력 누적은 현재 세션 파일에 저장됩니다.")
    try:
        current_bytes = Path(xlsx_path).read_bytes()
        st.sidebar.download_button(
            "현재 파일 다운로드(누적본)",
            data=current_bytes,
            file_name="액상잉크_Lot추적관리_UPDATED.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception:
        pass

    if st.sidebar.button("업로드 원본으로 되돌리기(초기화)"):
        Path(xlsx_path).write_bytes(st.session_state.get("_uploaded_bytes", uploaded.getvalue()))
        st.cache_data.clear()
        st.rerun()

if not Path(xlsx_path).exists():
    st.error(f"엑셀 파일을 찾을 수 없습니다: {xlsx_path}")
    st.stop()

# 반품 시트 보장
ensure_sheet_with_headers(
    xlsx_path,
    SHEET_BINDER_RETURN,
    headers=["반품일자", "바인더명", "관련 Lot(선택)", "반품수량(kg)", "비고"]
)

data = load_data(xlsx_path)
binder_df = data["binder"].copy()
single_df = data["single"].copy()
spec_binder = data["spec_binder"].copy()
spec_single = data["spec_single"].copy()
base_lab = data["base_lab"].copy()

if "제조/입고일" in binder_df.columns:
    binder_df["제조/입고일"] = binder_df["제조/입고일"].apply(normalize_date)
if "입고일" in single_df.columns:
    single_df["입고일"] = single_df["입고일"].apply(normalize_date)


# =========================
# Tabs
# =========================
tab_dash, tab_ink_in, tab_binder, tab_search = st.tabs(
    ["📊 대시보드", "✍️ 잉크 입고", "📦 바인더 입출고", "🔎 빠른검색"]
)


# =========================
# 1) Dashboard (표/그래프는 여기만)
# =========================
with tab_dash:
    with st.expander("📌 이 화면(대시보드) 읽는 방법", expanded=True):
        st.markdown(
            """
- **상단 요약**: 최근 30일 기준 *입고/부적합/평균 점도/최신일*을 한 번에 봅니다.  
- **단일색 표(엑셀형)**: 기간/색상군/제품코드/바인더Lot/검색으로 좁혀서 확인합니다.  
- **색상군 평균 점도(점+값)**: 색상군 평균만 빠르게 비교합니다.  
- **Lot 추이**: 선택 Lot의 점도 변화를 시간 순으로 봅니다.
            """
        )

    st.divider()

    today = dt.date.today()
    days = 30

    s_df = single_df.copy()
    s_df["_입고일_dt"] = pd.to_datetime(s_df.get("입고일", pd.Series(dtype=str)), errors="coerce")
    s_recent = s_df[s_df["_입고일_dt"].dt.date >= (today - dt.timedelta(days=days))].copy()
    s_recent_total = len(s_recent)
    s_recent_ng = int((s_recent.get("점도판정", pd.Series(dtype=str)) == "부적합").sum()) if "점도판정" in s_recent.columns else 0
    s_recent_ng_rate = (s_recent_ng / s_recent_total * 100.0) if s_recent_total else 0.0
    s_recent_mean = float(pd.to_numeric(s_recent.get("점도측정값(cP)", pd.Series(dtype=float)), errors="coerce").dropna().mean()) if s_recent_total else 0.0
    s_latest = s_df["_입고일_dt"].max()
    s_latest_txt = s_latest.date().isoformat() if pd.notna(s_latest) else "-"

    b_df = binder_df.copy()
    b_df["_일자_dt"] = pd.to_datetime(b_df.get("제조/입고일", pd.Series(dtype=str)), errors="coerce")
    b_recent = b_df[b_df["_일자_dt"].dt.date >= (today - dt.timedelta(days=days))].copy()
    b_recent_total = len(b_recent)
    b_recent_ng = int((b_recent.get("판정", pd.Series(dtype=str)) == "부적합").sum()) if "판정" in b_recent.columns else 0
    b_latest = b_df["_일자_dt"].max()
    b_latest_txt = b_latest.date().isoformat() if pd.notna(b_latest) else "-"

    c1, c2, c3, c4, c5 = st.columns([1.3, 1.3, 1.3, 1.3, 1.8])
    c1.metric(f"최근 {days}일 단일색 입고", f"{s_recent_total:,}")
    c2.metric(f"최근 {days}일 단일색 부적합", f"{s_recent_ng:,}", f"{s_recent_ng_rate:.1f}%")
    c3.metric(f"최근 {days}일 단일색 평균 점도", f"{s_recent_mean:,.0f} cP")
    c4.metric(f"최근 {days}일 바인더 입고", f"{b_recent_total:,}", f"부적합 {b_recent_ng:,}")
    c5.metric("데이터 최신일", f"단일색 {s_latest_txt} / 바인더 {b_latest_txt}")

    st.divider()

    # 1) 단일색 표(엑셀형)
    st.subheader("1) 단일색 데이터 (엑셀 형태)")

    df_view = single_df.copy()
    df_view["색차값(ΔE76)"] = df_view.get("비고", pd.Series(dtype=str)).apply(extract_delta_e_from_note)

    col_map = {
        "입고일": "제조일자(=입고일)",
        "색상군": "색상군",
        "제품코드": "제품코드",
        "사용된 바인더 Lot": "사용된바인더Lot",
        "점도측정값(cP)": "점도(cP)",
        "색차값(ΔE76)": "색차값(ΔE76)",
    }
    keep_cols = [c for c in col_map.keys() if c in df_view.columns]
    df_show = df_view[keep_cols].rename(columns=col_map)

    df_show["_date"] = pd.to_datetime(df_show.get("제조일자(=입고일)", pd.Series(dtype=str)), errors="coerce")
    dmin = df_show["_date"].min()
    dmax = df_show["_date"].max()
    dmin = dmin.date() if pd.notna(dmin) else today - dt.timedelta(days=90)
    dmax = dmax.date() if pd.notna(dmax) else today

    f1, f2, f3, f4, f5 = st.columns([1.2, 1.2, 1.4, 1.6, 2.2])
    with f1:
        start = st.date_input("기간 시작", value=max(dmin, dmax - dt.timedelta(days=90)), key="tbl_start")
    with f2:
        end = st.date_input("기간 종료", value=dmax, key="tbl_end")
    with f3:
        cg_list = sorted(df_show["색상군"].dropna().astype(str).unique().tolist()) if "색상군" in df_show.columns else []
        cg_pick = st.multiselect("색상군", cg_list, key="tbl_cg")
    with f4:
        pc_list = sorted(df_show["제품코드"].dropna().astype(str).unique().tolist()) if "제품코드" in df_show.columns else []
        pc_pick = st.multiselect("제품코드", pc_list, key="tbl_pc")
    with f5:
        q = st.text_input("검색(바인더Lot/제품코드 등)", value="", key="tbl_q")

    if start > end:
        start, end = end, start

    df_filtered = df_show.copy()
    df_filtered = df_filtered[(df_filtered["_date"].dt.date >= start) & (df_filtered["_date"].dt.date <= end)]

    if cg_pick and "색상군" in df_filtered.columns:
        df_filtered = df_filtered[df_filtered["색상군"].astype(str).isin([str(x) for x in cg_pick])]
    if pc_pick and "제품코드" in df_filtered.columns:
        df_filtered = df_filtered[df_filtered["제품코드"].astype(str).isin([str(x) for x in pc_pick])]

    if q.strip():
        qq = q.strip()
        mask = False
        for c in ["사용된바인더Lot", "제품코드", "색상군"]:
            if c in df_filtered.columns:
                mask = mask | df_filtered[c].astype(str).str.contains(qq, case=False, na=False)
        df_filtered = df_filtered[mask]

    df_filtered = df_filtered.sort_values("_date", ascending=False).drop(columns=["_date"])
    if "색차값(ΔE76)" in df_filtered.columns:
        df_filtered["색차값(ΔE76)"] = pd.to_numeric(df_filtered["색차값(ΔE76)"], errors="coerce").round(2)

    st.caption(f"표시 건수: {len(df_filtered):,}건")
    st.dataframe(df_filtered, use_container_width=True, height=280)

    st.divider()

    # 색상군별 평균 점도 (점+값)
    st.subheader("색상군별 평균 점도 (점 + 평균값 표시)")

    if "색상군" in single_df.columns and "점도측정값(cP)" in single_df.columns:
        mean_df = (
            single_df[["색상군", "점도측정값(cP)"]]
            .dropna()
            .assign(**{"점도측정값(cP)": pd.to_numeric(single_df["점도측정값(cP)"], errors="coerce")})
            .dropna()
            .groupby("색상군", as_index=False)["점도측정값(cP)"]
            .mean()
            .rename(columns={"점도측정값(cP)": "평균점도"})
            .sort_values("평균점도", ascending=False)
        )

        pts = alt.Chart(mean_df).mark_point(size=220).encode(
            y=alt.Y("색상군:N", title="색상군", sort=mean_df["색상군"].tolist()),
            x=alt.X("평균점도:Q", title="평균 점도(cP)"),
            tooltip=[alt.Tooltip("색상군:N"), alt.Tooltip("평균점도:Q", format=".1f")],
        )
        txt = alt.Chart(mean_df).mark_text(dx=10).encode(
            y=alt.Y("색상군:N", sort=mean_df["색상군"].tolist()),
            x="평균점도:Q",
            text=alt.Text("평균점도:Q", format=".0f"),
        )
        st.altair_chart((pts + txt).interactive(), use_container_width=True)
    else:
        st.info("단일색 데이터에 '색상군' 또는 '점도측정값(cP)' 컬럼이 없습니다.")

    st.divider()

    # 2) Lot 추이(점 크게 + 라벨 토글)
    st.subheader("2) 단일색 점도 변화 추이 (Lot별)")
    st.caption("점은 크게 표시되며, 필요 시 점도값 라벨도 표시할 수 있습니다.")

    df = single_df.copy()
    need_cols = ["입고일", "단일색잉크 Lot", "점도측정값(cP)"]
    miss = [c for c in need_cols if c not in df.columns]
    if miss:
        st.warning(f"단일색 데이터에 필요한 컬럼이 없습니다: {miss}")
    else:
        df = df.dropna(subset=need_cols).copy()
        df["입고일"] = pd.to_datetime(df["입고일"], errors="coerce")
        df = df.dropna(subset=["입고일"]).sort_values("입고일")

        f1, f2, f3, f4, f5 = st.columns([1.2, 1.2, 1.6, 2.0, 1.0])
        with f1:
            dmin = df["입고일"].min().date()
            dmax = df["입고일"].max().date()
            start = st.date_input("시작일", value=max(dmin, dmax - dt.timedelta(days=90)), key="trend_start")
        with f2:
            end = st.date_input("종료일", value=dmax, key="trend_end")
        with f3:
            cg = st.multiselect("색상군", sorted(df["색상군"].dropna().unique().tolist()) if "색상군" in df.columns else [], key="trend_cg")
        with f4:
            pc = st.multiselect("제품코드", sorted(df["제품코드"].dropna().unique().tolist()) if "제품코드" in df.columns else [], key="trend_pc")
        with f5:
            show_labels = st.checkbox("점도값 표시", value=True, key="trend_labels")

        if start > end:
            start, end = end, start

        df = df[(df["입고일"].dt.date >= start) & (df["입고일"].dt.date <= end)]
        if cg and "색상군" in df.columns:
            df = df[df["색상군"].isin(cg)]
        if pc and "제품코드" in df.columns:
            df = df[df["제품코드"].isin(pc)]

        lot_list = sorted(df["단일색잉크 Lot"].astype(str).unique().tolist())
        default_pick = lot_list[-5:] if len(lot_list) > 5 else lot_list
        pick = st.multiselect("표시할 단일색 Lot(복수 선택)", lot_list, default=default_pick, key="trend_lots")
        if pick:
            df = df[df["단일색잉크 Lot"].astype(str).isin(pick)]

        if len(df) == 0:
            st.info("선택한 조건에 해당하는 데이터가 없습니다.")
        else:
            df = df.sort_values("입고일")
            tooltip_cols = ["입고일:T", "단일색잉크 Lot:N", "점도측정값(cP):Q"]

            base = alt.Chart(df).encode(
                x=alt.X("입고일:T", title="입고일"),
                y=alt.Y("점도측정값(cP):Q", title="점도(cP)"),
                color=alt.Color("단일색잉크 Lot:N", title="Lot"),
                tooltip=tooltip_cols,
            )
            line = base.mark_line()
            points = base.mark_point(size=260)

            chart = line + points
            if show_labels and len(df) <= 250:
                labels = alt.Chart(df).mark_text(dx=10, dy=-10).encode(
                    x="입고일:T",
                    y="점도측정값(cP):Q",
                    color=alt.Color("단일색잉크 Lot:N", legend=None),
                    text=alt.Text("점도측정값(cP):Q", format=".0f"),
                )
                chart = chart + labels
            elif show_labels and len(df) > 250:
                st.info("데이터가 많아 라벨 표시는 자동으로 생략했습니다(250건 이하에서만 표시).")

            st.altair_chart(chart.interactive(), use_container_width=True)


# =========================
# 2) 잉크 입고 (단일색 입력)
# =========================
with tab_ink_in:
    st.subheader("단일색 잉크 입력")

    ink_types = ["HEMA", "Silicone"]
    color_groups = sorted(spec_single["색상군"].dropna().unique().tolist())
    product_codes = sorted(spec_single["제품코드"].dropna().unique().tolist())

    binder_lots = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).tolist()
    binder_lots = sorted(set(binder_lots), reverse=True)

    with st.form("single_form", clear_on_submit=True):
        col1, col2, col3, col4 = st.columns([1.2, 1.3, 1.5, 2.0])
        with col1:
            in_date = st.date_input("입고일", value=dt.date.today(), key="single_in_date")
            ink_type = st.selectbox("잉크타입", ink_types)
            color_group = st.selectbox("색상군", color_groups)
        with col2:
            product_code = st.selectbox("제품코드", product_codes)
            binder_lot = st.selectbox("사용된 바인더 Lot", binder_lots)
        with col3:
            visc_meas = st.number_input("점도측정값(cP)", min_value=0.0, step=1.0, format="%.1f")
            supplier = st.selectbox("바인더제조처", ["내부", "외주"], index=0)
        with col4:
            st.caption("선택: 착색력(L*a*b*) 입력 시 ΔE(76)을 비고에 자동 기록합니다.")
            L = st.number_input("착색력_L*", value=0.0, step=0.1, format="%.2f")
            a = st.number_input("착색력_a*", value=0.0, step=0.1, format="%.2f")
            b = st.number_input("착색력_b*", value=0.0, step=0.1, format="%.2f")
            lab_enabled = st.checkbox("L*a*b* 입력함", value=False)

        note = st.text_input("비고", value="", key="single_note")
        submit_s = st.form_submit_button("저장(단일색)", type="primary")

    if submit_s:
        binder_type = infer_binder_type_from_lot(spec_binder, binder_lot)

        spec_hit = spec_single[
            (spec_single["색상군"] == color_group) &
            (spec_single["제품코드"] == product_code)
        ].copy()

        if binder_type and "BinderType" in spec_hit.columns:
            spec_hit = spec_hit[spec_hit["BinderType"] == binder_type]

        if len(spec_hit) == 0:
            lo, hi = None, None
            visc_judge = None
            st.warning("점도 기준을 Spec_Single_H&S에서 찾지 못했습니다. (색상군/제품코드/바인더타입 조합 확인)")
        else:
            lo = float(spec_hit["하한"].iloc[0])
            hi = float(spec_hit["상한"].iloc[0])
            visc_judge = judge_range(visc_meas, lo, hi)

        new_lot = generate_single_lot(single_df, product_code, color_group, in_date)
        if new_lot is None:
            st.error("단일색 Lot 자동 생성에 실패했습니다. (색상군 매핑 확인 필요)")
        else:
            note2 = note
            if lab_enabled:
                base_hit = base_lab[base_lab.get("제품코드", pd.Series(dtype=str)) == product_code]
                if len(base_hit) == 1:
                    base = (float(base_hit.iloc[0]["기준_L*"]), float(base_hit.iloc[0]["기준_a*"]), float(base_hit.iloc[0]["기준_b*"]))
                    de = delta_e76((L, a, b), base)
                    note2 = (note2 + " " if note2 else "") + f"[ΔE76={de:.2f}]"
                else:
                    note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"

            row = {
                "입고일": in_date,
                "잉크타입\n(HEMA/Silicone)": ink_type,
                "색상군": color_group,
                "제품코드": product_code,
                "단일색잉크 Lot": new_lot,
                "사용된 바인더 Lot": binder_lot,
                "바인더제조처\n(내부/외주)": supplier,
                "BinderType(자동)": binder_type,
                "점도측정값(cP)": float(visc_meas),
                "점도하한": lo,
                "점도상한": hi,
                "점도판정": visc_judge,
                "착색력_L*": float(L) if lab_enabled else None,
                "착색력_a*": float(a) if lab_enabled else None,
                "착색력_b*": float(b) if lab_enabled else None,
                "비고": note2,
            }

            try:
                append_row_to_sheet(xlsx_path, SHEET_SINGLE, row)
                st.success(f"저장 완료! 단일색 Lot = {new_lot} / 점도판정 = {visc_judge}")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")


# =========================
# 3) 바인더 입출고
# =========================
with tab_binder:
    st.subheader("바인더 입출고")

    # 업체 반환 입력(kg)
    st.markdown("### ✅ 업체 반환(반품) 입력 (kg 단위)")
    binder_names = sorted(spec_binder["바인더명"].dropna().unique().tolist())
    binder_lot_choices = sorted(binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).unique().tolist(), reverse=True)

    with st.form("binder_return_form", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([1.2, 1.6, 2.0, 1.4])
        with c1:
            ret_date = st.date_input("반품일자", value=dt.date.today(), key="ret_date")
        with c2:
            ret_name = st.selectbox("바인더명", binder_names, key="ret_name")
        with c3:
            ret_lot = st.selectbox("관련 Lot(선택)", ["(선택안함)"] + binder_lot_choices, key="ret_lot")
        with c4:
            ret_kg = st.number_input("반품수량(kg)", min_value=0.0, step=0.1, format="%.1f", key="ret_kg")
        ret_note = st.text_input("비고", value="", key="ret_note")
        ret_submit = st.form_submit_button("반품 저장", type="primary")

    if ret_submit:
        if ret_kg <= 0:
            st.error("반품수량(kg)은 0보다 커야 합니다.")
        else:
            row = {
                "반품일자": ret_date,
                "바인더명": ret_name,
                "관련 Lot(선택)": "" if ret_lot == "(선택안함)" else ret_lot,
                "반품수량(kg)": float(ret_kg),
                "비고": ret_note,
            }
            try:
                append_row_to_sheet(xlsx_path, SHEET_BINDER_RETURN, row)
                st.success("반품 저장 완료!")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    st.divider()

    # 엑셀 바인더 기록(확인용)
    st.markdown("### 📌 엑셀 바인더 기록(최신순, 최근 50건)")
    if "제조/입고일" in binder_df.columns:
        st.dataframe(binder_df.sort_values("제조/입고일", ascending=False).head(50), use_container_width=True)
    else:
        st.dataframe(binder_df.head(50), use_container_width=True)

    st.divider()

    # 구글시트 조회(최신순)
    st.markdown("### 바인더 입출고(구글시트) 조회 - 최신순")
    st.caption("※ 여기 표는 구글시트이며, 위 입력(엑셀 저장)과는 별개입니다.")

    try:
        df_hema = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_HEMA)
        df_sil = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_SIL)
    except Exception as e:
        st.error("구글시트에서 데이터를 불러오지 못했습니다. 공유/웹게시/시트명/ID를 확인해주세요.")
        st.exception(e)
        st.stop()

    df_hema = sort_df_by_any_date_col(df_hema)
    df_sil = sort_df_by_any_date_col(df_sil)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### HEMA (최신순)")
        st.dataframe(df_hema, use_container_width=True)
    with c2:
        st.markdown("#### Silicon (최신순)")
        st.dataframe(df_sil, use_container_width=True)

    if st.button("지금 최신값으로 다시 불러오기", key="binder_refresh"):
        st.cache_data.clear()
        st.rerun()


# =========================
# 4) Search
# =========================
with tab_search:
    c1, c2, c3 = st.columns([2, 2, 3])
    with c1:
        mode = st.selectbox("검색 종류", ["바인더 Lot", "단일색 잉크 Lot", "제품코드", "색상군", "기간(입고일)"])
    with c2:
        q = st.text_input("검색어", placeholder="예: PCB20250112-01 / PLB25041501 / PL-835-1 ...")
    with c3:
        st.write("")
        st.caption("💡 바인더 Lot 검색: 바인더 정보 + 연결된 단일색 잉크 목록을 함께 보여줍니다.")

    if mode == "기간(입고일)":
        d1, d2 = st.columns(2)
        with d1:
            start = st.date_input("시작일", value=dt.date.today() - dt.timedelta(days=30), key="search_start")
        with d2:
            end = st.date_input("종료일", value=dt.date.today(), key="search_end")
        df = single_df.copy()
        if "입고일" in df.columns:
            df = df[df["입고일"].between(start, end)]
        st.subheader("단일색_수입검사")
        st.dataframe(df, use_container_width=True)

    elif mode == "바인더 Lot":
        b = binder_df.copy()
        b_hit = df_quick_filter(b, q, ["Lot(자동)", "바인더명", "비고"])
        st.subheader("바인더_제조_입고")
        if "제조/입고일" in b_hit.columns:
            st.dataframe(b_hit.sort_values(by="제조/입고일", ascending=False), use_container_width=True)
        else:
            st.dataframe(b_hit, use_container_width=True)

        if q and "사용된 바인더 Lot" in single_df.columns:
            s_hit = single_df[single_df["사용된 바인더 Lot"].astype(str).str.contains(str(q).strip(), case=False, na=False)]
            st.subheader("연결된 단일색_수입검사 (사용된 바인더 Lot)")
            if "입고일" in s_hit.columns:
                st.dataframe(s_hit.sort_values(by="입고일", ascending=False), use_container_width=True)
            else:
                st.dataframe(s_hit, use_container_width=True)

    elif mode == "단일색 잉크 Lot":
        s = single_df.copy()
        s_hit = df_quick_filter(s, q, ["단일색잉크 Lot", "제품코드", "사용된 바인더 Lot", "색상군", "비고"])
        st.subheader("단일색_수입검사")
        if "입고일" in s_hit.columns:
            st.dataframe(s_hit.sort_values(by="입고일", ascending=False), use_container_width=True)
        else:
            st.dataframe(s_hit, use_container_width=True)

    elif mode == "제품코드":
        s = single_df.copy()
        s_hit = df_quick_filter(s, q, ["제품코드"])
        st.subheader("단일색_수입검사")
        if "입고일" in s_hit.columns:
            st.dataframe(s_hit.sort_values(by="입고일", ascending=False), use_container_width=True)
        else:
            st.dataframe(s_hit, use_container_width=True)

    elif mode == "색상군":
        s = single_df.copy()
        s_hit = df_quick_filter(s, q, ["색상군"])
        st.subheader("단일색_수입검사")
        if "입고일" in s_hit.columns:
            st.dataframe(s_hit.sort_values(by="입고일", ascending=False), use_container_width=True)
        else:
            st.dataframe(s_hit, use_container_width=True)
