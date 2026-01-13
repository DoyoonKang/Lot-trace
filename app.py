import altair as alt
import streamlit as st
import pandas as pd
import datetime as dt
import re
from pathlib import Path
from openpyxl import load_workbook
import requests
from io import StringIO


# =========================
# Page Config (딱 1번만!)
# =========================
st.set_page_config(
    page_title="액상 잉크 Lot 추적 관리",
    page_icon="🧪",
    layout="wide",
)

# =========================
# Constants
# =========================
DEFAULT_XLSX = "액상잉크_Lot추적관리_FINAL.xlsx"

SHEET_BINDER = "바인더_제조_입고"
SHEET_SINGLE = "단일색_수입검사"
SHEET_SPEC_BINDER = "Spec_Binder"
SHEET_SPEC_SINGLE = "Spec_Single_H&S"
SHEET_BINDER_VISC = "Binder_Visc"
SHEET_BASE_LAB = "기준LAB"

# ✅ 반품(업체 반환) 로그 시트
SHEET_BINDER_RETURN = "바인더_반품로그"
BINDER_RETURN_HEADERS = [
    "반품일",
    "바인더구분(HEMA/Silicon)",
    "바인더명",
    "Lot(자동)",
    "반품수량(kg)",
    "비고",
]

# ✅ 1통 기준(kg)
DRUM_KG = 20.0

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
# Helpers
# =========================
def _read_excel_from_path(xlsx_path: str) -> dict[str, pd.DataFrame]:
    def read(name: str) -> pd.DataFrame:
        return pd.read_excel(xlsx_path, sheet_name=name)

    data = {
        "binder": read(SHEET_BINDER),
        "single": read(SHEET_SINGLE),
        "spec_binder": read(SHEET_SPEC_BINDER),
        "spec_single": read(SHEET_SPEC_SINGLE),
        "binder_visc": read(SHEET_BINDER_VISC),
        "base_lab": read(SHEET_BASE_LAB),
    }

    # 반품로그는 없을 수 있음
    try:
        data["binder_return"] = read(SHEET_BINDER_RETURN)
    except Exception:
        data["binder_return"] = pd.DataFrame(columns=BINDER_RETURN_HEADERS)

    return data


@st.cache_data(show_spinner=False)
def load_data(xlsx_path: str) -> dict[str, pd.DataFrame]:
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


def to_date_safe(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    if isinstance(x, dt.datetime):
        return x.date()
    if isinstance(x, dt.date):
        return x
    try:
        return pd.to_datetime(x).date()
    except Exception:
        return None


def delta_e76(lab1, lab2):
    return float(((lab1[0]-lab2[0])**2 + (lab1[1]-lab2[1])**2 + (lab1[2]-lab2[2])**2) ** 0.5)


def get_binder_limits(spec_binder: pd.DataFrame, binder_name: str):
    df = spec_binder[spec_binder["바인더명"] == binder_name].copy()
    visc = df[df["시험항목"].astype(str).str.contains("점도", na=False)]
    uv = df[df["시험항목"].astype(str).str.contains("UV", na=False)]

    visc_lo = float(visc["하한"].dropna().iloc[0]) if len(visc["하한"].dropna()) else None
    visc_hi = float(visc["상한"].dropna().iloc[0]) if len(visc["상한"].dropna()) else None
    uv_hi = float(uv["상한"].dropna().iloc[0]) if len(uv["상한"].dropna()) else None
    rule = df["Lot부여규칙"].dropna().iloc[0] if "Lot부여규칙" in df.columns and len(df["Lot부여규칙"].dropna()) else None
    return visc_lo, visc_hi, uv_hi, rule


def next_seq_for_pattern(existing_lots: pd.Series, prefix: str, date_str: str, digits: int = 2, sep: str = "-"):
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


def generate_binder_lot(spec_binder: pd.DataFrame, binder_name: str, mfg_date: dt.date, existing_binder_lots: pd.Series):
    _, _, _, rule = get_binder_limits(spec_binder, binder_name)
    if not rule:
        code = re.sub(r"\W+", "", binder_name)[:6].upper()
        return f"{code}{mfg_date.strftime('%Y%m%d')}-01"

    m = re.match(r"^([A-Za-z0-9]+)\+YYYYMMDD(-##)?$", str(rule).strip())
    if not m:
        code = re.sub(r"\W+", "", binder_name)[:6].upper()
        return f"{code}{mfg_date.strftime('%Y%m%d')}-01"

    prefix = m.group(1)
    has_seq = bool(m.group(2))
    date_str = mfg_date.strftime("%Y%m%d")
    if has_seq:
        seq = next_seq_for_pattern(existing_binder_lots, prefix, date_str, digits=2, sep="-")
        return f"{prefix}{date_str}-{seq:02d}"
    return f"{prefix}{date_str}"


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


def ensure_sheet(xlsx_path: str, sheet_name: str, headers: list[str]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        wb.save(xlsx_path)
        return

    ws = wb[sheet_name]
    first_row = [c.value for c in ws[1]]
    if all(v is None for v in first_row):
        ws.delete_rows(1)
        ws.append(headers)
        wb.save(xlsx_path)
        return

    wb.save(xlsx_path)


def ensure_return_sheet_schema(xlsx_path: str):
    ensure_sheet(xlsx_path, SHEET_BINDER_RETURN, BINDER_RETURN_HEADERS)

    wb = load_workbook(xlsx_path)
    ws = wb[SHEET_BINDER_RETURN]
    headers = [c.value for c in ws[1]]

    if "반품수량(kg)" not in headers and "수량(통)" in headers:
        idx = headers.index("수량(통)") + 1
        ws.cell(row=1, column=idx).value = "반품수량(kg)"
        wb.save(xlsx_path)
        return

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


def append_rows_to_sheet(xlsx_path: str, sheet_name: str, rows: list[dict]):
    if not rows:
        return
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]

    for row in rows:
        values = [row.get(h, None) for h in headers]
        ws.append(values)

    wb.save(xlsx_path)


def df_quick_filter(df: pd.DataFrame, text: str, cols: list[str]):
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


def _pick_col(df: pd.DataFrame, candidates: list[str]):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _pick_col_contains(df: pd.DataFrame, keywords: list[str]):
    for col in df.columns:
        low = str(col).lower()
        for kw in keywords:
            if kw.lower() in low:
                return col
    return None


def try_sort_by_date_desc(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    date_col = _pick_col(df, ["일자", "날짜", "Date", "date", "입출고일", "입고일", "출고일", "반출일", "사용일"])
    if date_col is None:
        date_col = _pick_col_contains(df, ["date", "일자", "날짜"])

    if date_col is None:
        return df

    dd = df.copy()
    dd[date_col] = pd.to_datetime(dd[date_col], errors="coerce")
    dd = dd.sort_values(by=date_col, ascending=False, na_position="last")
    return dd


def compute_delta_e_column(single_df: pd.DataFrame, base_lab: pd.DataFrame) -> pd.Series:
    """
    ΔE76 컬럼을 최대한 계산해서 반환합니다.
    - 착색력_L*/a*/b* + 기준LAB(제품코드 매칭)이 있으면 계산
    - 없으면 비고에서 [ΔE76=...] 패턴을 찾아 숫자만 추출
    """
    # base lookup
    base_map = {}
    if base_lab is not None and not base_lab.empty and "제품코드" in base_lab.columns:
        for _, r in base_lab.dropna(subset=["제품코드"]).iterrows():
            try:
                base_map[str(r["제품코드"]).strip()] = (
                    float(r["기준_L*"]), float(r["기준_a*"]), float(r["기준_b*"])
                )
            except Exception:
                continue

    def _row_de(r):
        # 1) compute by LAB columns
        try:
            pc = str(r.get("제품코드", "")).strip()
            if pc and pc in base_map:
                L = r.get("착색력_L*", None)
                a = r.get("착색력_a*", None)
                b = r.get("착색력_b*", None)
                if pd.notna(L) and pd.notna(a) and pd.notna(b):
                    return delta_e76((float(L), float(a), float(b)), base_map[pc])
        except Exception:
            pass

        # 2) parse from note
        try:
            note = str(r.get("비고", ""))
            m = re.search(r"ΔE76\s*=\s*([0-9]+(?:\.[0-9]+)?)", note)
            if m:
                return float(m.group(1))
        except Exception:
            pass

        return None

    return single_df.apply(_row_de, axis=1)


# =========================
# UI Header
# =========================
st.title("액상 잉크 Lot 추적 관리 대시보드")
st.caption("✅ 빠른 검색 + ✅ 잉크 입고(엑셀 누적) + ✅ 대시보드(표+평균점도 점그래프/추이) + ✅ 바인더 입출고(구글시트 자동 반영) + ✅ 반품(업체 반환) kg 입력")


# =========================
# Data file selection (Excel)
# =========================
with st.sidebar:
    st.header("데이터 파일")
    xlsx_path = st.text_input(
        "엑셀 파일 경로",
        value=DEFAULT_XLSX,
        help="로컬 실행 시, app.py와 같은 폴더에 엑셀을 두면 기본값 그대로 사용 가능합니다."
    )
    uploaded = st.file_uploader("또는 엑셀 업로드(읽기 전용 권장)", type=["xlsx"])

if uploaded is not None:
    tmp_bytes = uploaded.read()
    tmp_path = Path(st.session_state.get("_tmp_xlsx_path", ".streamlit_tmp.xlsx"))
    tmp_path.write_bytes(tmp_bytes)
    xlsx_path = str(tmp_path)
    st.sidebar.info("업로드 파일로 실행 중입니다. (이 모드에서는 저장해도 서버에 영구 누적이 보장되지 않을 수 있습니다.)")

if not Path(xlsx_path).exists():
    st.error(f"엑셀 파일을 찾을 수 없습니다: {xlsx_path}")
    st.stop()

ensure_return_sheet_schema(xlsx_path)

data = load_data(xlsx_path)
binder_df = data["binder"].copy()
single_df = data["single"].copy()
spec_binder = data["spec_binder"].copy()
spec_single = data["spec_single"].copy()
base_lab = data["base_lab"].copy()
binder_return_df = data.get("binder_return", pd.DataFrame(columns=BINDER_RETURN_HEADERS)).copy()

if "제조/입고일" in binder_df.columns:
    binder_df["제조/입고일"] = binder_df["제조/입고일"].apply(normalize_date)
if "입고일" in single_df.columns:
    single_df["입고일"] = single_df["입고일"].apply(normalize_date)

if "반품수량(kg)" not in binder_return_df.columns and "수량(통)" in binder_return_df.columns:
    binder_return_df = binder_return_df.rename(columns={"수량(통)": "반품수량(kg)"})


# =========================
# Tabs
# =========================
tab_dash, tab_ink_in, tab_binder, tab_search = st.tabs(
    ["📊 대시보드", "✍️ 잉크 입고", "📦 바인더 입출고", "🔎 빠른검색"]
)


# =========================
# Dashboard (그래프/표는 여기만)
# =========================
with tab_dash:
    b_total = len(binder_df)
    s_total = len(single_df)
    b_ng = int((binder_df.get("판정", pd.Series(dtype=str)) == "부적합").sum()) if "판정" in binder_df.columns else 0
    s_ng = int((single_df.get("점도판정", pd.Series(dtype=str)) == "부적합").sum()) if "점도판정" in single_df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("바인더 기록", f"{b_total:,}")
    c2.metric("바인더 부적합", f"{b_ng:,}")
    c3.metric("단일색 기록", f"{s_total:,}")
    c4.metric("단일색(점도) 부적합", f"{s_ng:,}")

    st.divider()

    # =========================================================
    # 1) 단일색 표(엑셀처럼) + 색상군별 평균 점도(점 + 라벨)
    # =========================================================
    st.subheader("1) 단일색 현황 (엑셀 형태) + 색상군별 평균 점도")

    needed = ["입고일", "색상군", "제품코드", "사용된 바인더 Lot", "점도측정값(cP)"]
    miss = [c for c in needed if c not in single_df.columns]
    if miss:
        st.warning(f"단일색 데이터에 필요한 컬럼이 없습니다: {miss}")
    else:
        # ΔE76 계산(가능한 범위에서)
        df_table = single_df.copy()
        df_table["ΔE76"] = compute_delta_e_column(df_table, base_lab)

        # 표시용 컬럼명(입고일 -> 제조일자)
        display = pd.DataFrame({
            "제조일자": df_table["입고일"],
            "색상군": df_table["색상군"],
            "제품코드": df_table["제품코드"],
            "사용된바인더": df_table["사용된 바인더 Lot"],
            "점도(cP)": df_table["점도측정값(cP)"],
            "색차값(ΔE76)": df_table["ΔE76"],
        })

        # 필터(가볍게: 기간/색상군/제품코드)
        f1, f2, f3 = st.columns([1.3, 1.7, 2.0])
        with f1:
            dmin = pd.to_datetime(display["제조일자"], errors="coerce").min()
            dmax = pd.to_datetime(display["제조일자"], errors="coerce").max()
            if pd.isna(dmin) or pd.isna(dmax):
                start = dt.date.today() - dt.timedelta(days=90)
                end = dt.date.today()
            else:
                start = st.date_input("시작일", value=max(dmin.date(), dmax.date() - dt.timedelta(days=90)), key="dash_tbl_start")
                end = st.date_input("종료일", value=dmax.date(), key="dash_tbl_end")
        with f2:
            cg_list = sorted(display["색상군"].dropna().astype(str).unique().tolist())
            cg_pick = st.multiselect("색상군", cg_list, default=[], key="dash_tbl_cg")
        with f3:
            pc_list = sorted(display["제품코드"].dropna().astype(str).unique().tolist())
            pc_pick = st.multiselect("제품코드", pc_list, default=[], key="dash_tbl_pc")

        # 적용
        dcol = pd.to_datetime(display["제조일자"], errors="coerce")
        mask = (dcol.dt.date >= start) & (dcol.dt.date <= end)
        if cg_pick:
            mask = mask & (display["색상군"].astype(str).isin(cg_pick))
        if pc_pick:
            mask = mask & (display["제품코드"].astype(str).isin(pc_pick))

        view = display.loc[mask].copy()
        view["제조일자"] = pd.to_datetime(view["제조일자"], errors="coerce")
        view = view.sort_values(by="제조일자", ascending=False, na_position="last")
        view["제조일자"] = view["제조일자"].dt.date

        st.dataframe(view, use_container_width=True, height=340)

        # ---- 평균 점도(점 + 라벨)
        st.markdown("##### 색상군별 평균 점도 (점 + 평균값 표시)")
        mean_src = view.dropna(subset=["색상군", "점도(cP)"]).copy()
        if mean_src.empty:
            st.info("현재 필터 조건에서 평균 점도를 계산할 데이터가 없습니다.")
        else:
            mean_df = (
                mean_src.groupby("색상군", as_index=False)["점도(cP)"]
                .mean()
                .rename(columns={"점도(cP)": "평균점도"})
            )

            # 색상군 정렬(알파벳)
            mean_df["색상군"] = mean_df["색상군"].astype(str)
            mean_df = mean_df.sort_values("색상군")

            base = alt.Chart(mean_df).encode(
                x=alt.X("색상군:N", sort=mean_df["색상군"].tolist(), title="색상군"),
                y=alt.Y("평균점도:Q", title="평균 점도(cP)"),
                tooltip=[
                    alt.Tooltip("색상군:N"),
                    alt.Tooltip("평균점도:Q", format=",.1f"),
                ],
            )

            points = base.mark_point(size=120)
            labels = base.mark_text(align="left", dx=8, dy=-6).encode(
                text=alt.Text("평균점도:Q", format=",.0f")
            )

            st.altair_chart((points + labels).interactive(), use_container_width=True)

    st.divider()

    # =========================
    # 2) 단일색 점도 변화 추이 (Lot별)
    # =========================
    st.subheader("2) 단일색 점도 변화 추이 (Lot별)")
    st.caption("선택한 Lot별로 '입고일' 기준으로 선으로 연결해 추이를 확인합니다.")

    df = single_df.copy()
    need_cols = ["입고일", "단일색잉크 Lot", "점도측정값(cP)"]
    miss2 = [c for c in need_cols if c not in df.columns]
    if miss2:
        st.warning(f"단일색 데이터에 필요한 컬럼이 없습니다: {miss2}")
    else:
        df = df.dropna(subset=need_cols).copy()
        df["입고일"] = pd.to_datetime(df["입고일"])
        df = df.sort_values("입고일")

        f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
        with f1:
            dmin = df["입고일"].min().date()
            dmax = df["입고일"].max().date()
            start = st.date_input("시작일(추이)", value=max(dmin, dmax - dt.timedelta(days=90)), key="trend_start")
        with f2:
            end = st.date_input("종료일(추이)", value=dmax, key="trend_end")
        with f3:
            cg = st.multiselect("색상군(추이)", sorted(df["색상군"].dropna().unique().tolist()) if "색상군" in df.columns else [], key="trend_cg")
        with f4:
            pc = st.multiselect("제품코드(추이)", sorted(df["제품코드"].dropna().unique().tolist()) if "제품코드" in df.columns else [], key="trend_pc")

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

        if df.empty:
            st.info("선택한 조건에 해당하는 데이터가 없습니다.")
        else:
            df = df.sort_values(["단일색잉크 Lot", "입고일"])

            tooltip_cols = ["입고일:T", "단일색잉크 Lot:N", "점도측정값(cP):Q"]
            if "제품코드" in df.columns:
                tooltip_cols.insert(2, "제품코드:N")
            if "색상군" in df.columns:
                tooltip_cols.insert(3, "색상군:N")
            if "사용된 바인더 Lot" in df.columns:
                tooltip_cols.append("사용된 바인더 Lot:N")

            chart = (
                alt.Chart(df)
                .mark_line(point=True)
                .encode(
                    x=alt.X("입고일:T", title="입고일"),
                    y=alt.Y("점도측정값(cP):Q", title="점도(cP)"),
                    color=alt.Color("단일색잉크 Lot:N", title="Lot"),
                    tooltip=tooltip_cols,
                )
                .interactive()
            )
            st.altair_chart(chart, use_container_width=True)


# =========================
# Ink In (단일색 입력)
# =========================
with tab_ink_in:
    st.info("이 탭은 **단일색 잉크 입고 데이터**를 엑셀에 누적 저장합니다.")
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
            ink_type = st.selectbox("잉크타입", ink_types, key="single_ink_type")
            color_group = st.selectbox("색상군", color_groups, key="single_color_group")
        with col2:
            product_code = st.selectbox("제품코드", product_codes, key="single_product_code")
            binder_lot = st.selectbox("사용된 바인더 Lot", binder_lots, key="single_binder_lot")
        with col3:
            visc_meas = st.number_input("점도측정값(cP)", min_value=0.0, step=1.0, format="%.1f", key="single_visc")
            supplier = st.selectbox("바인더제조처", ["내부", "외주"], index=0, key="single_supplier")
        with col4:
            st.caption("선택: 착색력(L*a*b*) 입력 시 기준LAB이 있으면 ΔE(76)을 자동 계산해 비고에 기록합니다.")
            L = st.number_input("착색력_L*", value=0.0, step=0.1, format="%.2f", key="single_L")
            a = st.number_input("착색력_a*", value=0.0, step=0.1, format="%.2f", key="single_a")
            b = st.number_input("착색력_b*", value=0.0, step=0.1, format="%.2f", key="single_b")
            lab_enabled = st.checkbox("L*a*b* 입력함", value=False, key="single_lab_en")

        note = st.text_input("비고", value="", key="single_note")
        submit_s = st.form_submit_button("저장(단일색)")

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
            st.warning("점도 기준을 Spec_Single_H&S에서 찾지 못했습니다. (색상군/제품코드/바인더타입 조합 확인 필요)")
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
# Binder IO (반품 입력 + 바인더 입력)
# =========================
with tab_binder:
    st.subheader("바인더 반품(업체 반환) 입력 (kg 단위)")
    st.caption(f"기준: 1통 = {DRUM_KG:.0f}kg. 예) 20kg 중 10kg 사용 → 10kg 반품 입력")

    binder_lots = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).tolist()
    binder_lots = sorted(set(binder_lots), reverse=True)

    with st.form("binder_return_form", clear_on_submit=True):
        r1, r2, r3 = st.columns([1.2, 1.2, 2.6])
        with r1:
            ret_date = st.date_input("반품일", value=dt.date.today(), key="ret_date")
        with r2:
            ret_type = st.selectbox("구분", ["HEMA", "Silicon"], key="ret_type")
        with r3:
            ret_binder_name = st.text_input("바인더명(선택)", value="", key="ret_name")

        r4, r5, r6 = st.columns([2.0, 1.3, 2.0])
        with r4:
            ret_lot = st.selectbox("반품 Lot(선택)", [""] + binder_lots, key="ret_lot")
        with r5:
            ret_kg = st.number_input("반품 수량(kg)", min_value=0.0, step=0.1, format="%.1f", value=0.0, key="ret_kg")
        with r6:
            drums = (ret_kg / DRUM_KG) if DRUM_KG else 0.0
            st.markdown(f"**통 환산:** {drums:.2f} 통")

        ret_note = st.text_input("비고", value="", key="ret_note")
        submit_ret = st.form_submit_button("반품 저장")

    if submit_ret:
        if ret_kg <= 0:
            st.error("반품 수량(kg)은 0보다 커야 합니다.")
        else:
            ensure_return_sheet_schema(xlsx_path)
            row = {
                "반품일": ret_date,
                "바인더구분(HEMA/Silicon)": ret_type,
                "바인더명": ret_binder_name.strip(),
                "Lot(자동)": ret_lot.strip(),
                "반품수량(kg)": float(ret_kg),
                "비고": ret_note.strip(),
            }
            try:
                append_row_to_sheet(xlsx_path, SHEET_BINDER_RETURN, row)
                st.success("반품 내역 저장 완료!")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"반품 저장 실패: {e}")

    st.divider()

    st.subheader("바인더 입력 (제조/입고) — 이 탭 최상단에서 입력")
    binder_names = sorted(spec_binder["바인더명"].dropna().unique().tolist())

    input_mode = st.radio(
        "입력 방식",
        ["개별 입력", "일괄 입력(날짜/수량/점도 직접 입력)"],
        horizontal=True,
        key="binder_input_mode"
    )

    if input_mode == "개별 입력":
        with st.form("binder_form_single", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                mfg_date = st.date_input("제조/입고일", value=dt.date.today(), key="b_single_date")
                binder_name = st.selectbox("바인더명", binder_names, key="b_single_name")
            with col2:
                visc = st.number_input("점도(cP)", min_value=0.0, step=1.0, format="%.1f", key="b_single_visc")
                uv = st.number_input("UV흡광도(선택)", min_value=0.0, step=0.01, format="%.3f", key="b_single_uv")
                uv_enabled = st.checkbox("UV 값 입력함", value=False, key="b_single_uv_en")
            with col3:
                note = st.text_input("비고", value="", key="b_single_note")
                submit_b = st.form_submit_button("저장(바인더)")

        if submit_b:
            visc_lo, visc_hi, uv_hi, _ = get_binder_limits(spec_binder, binder_name)
            lot = generate_binder_lot(spec_binder, binder_name, mfg_date, binder_df.get("Lot(자동)", pd.Series(dtype=str)))

            judge_v = judge_range(visc, visc_lo, visc_hi)
            judge_u = judge_range(uv if uv_enabled else None, None, uv_hi)
            judge = "부적합" if (judge_v == "부적합" or judge_u == "부적합") else "적합"

            row = {
                "제조/입고일": mfg_date,
                "바인더명": binder_name,
                "Lot(자동)": lot,
                "점도(cP)": float(visc),
                "UV흡광도(선택)": float(uv) if uv_enabled else None,
                "판정": judge,
                "비고": note,
            }
            try:
                append_row_to_sheet(xlsx_path, SHEET_BINDER, row)
                st.success(f"저장 완료! 바인더 Lot = {lot}")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    else:
        st.caption("✅ 날짜별로 필요할 때만 행을 추가하고, 저장 버튼 한 번으로 일괄 저장합니다.")
        binder_name = st.selectbox("바인더명(공통)", binder_names, key="b_batch_name")

        st.markdown("#### 새 행 기본값")
        d1, d2, d3, d4, d5 = st.columns([1.2, 1.1, 1.2, 1.2, 2.3])
        with d1:
            default_date = st.date_input("기본 날짜", value=dt.date.today(), key="b_def_date")
        with d2:
            default_qty = st.number_input("기본 수량(통)", min_value=1, max_value=200, value=8, step=1, key="b_def_qty")
        with d3:
            default_visc = st.number_input("기본 점도(cP)", min_value=0.0, step=1.0, format="%.1f", key="b_def_visc")
        with d4:
            default_uv = st.number_input("기본 UV(선택)", min_value=0.0, step=0.01, format="%.3f", key="b_def_uv")
            default_uv_use = st.checkbox("UV 사용", value=False, key="b_def_uv_use")
        with d5:
            default_note = st.text_input("기본 비고", value="", key="b_def_note")

        st.markdown("#### 날짜별 입고 행(직접 입력)")
        if "b_batch_table" not in st.session_state or st.session_state["b_batch_table"] is None:
            st.session_state["b_batch_table"] = pd.DataFrame([{
                "제조/입고일": dt.date.today(),
                "수량(통)": 8,
                "점도(cP)": 0.0,
                "UV흡광도(선택)": None,
                "비고": ""
            }])

        cbtn1, cbtn2 = st.columns([1.2, 2.8])
        with cbtn1:
            if st.button("행 추가(기본값)", key="b_add_row"):
                df0 = st.session_state["b_batch_table"].copy()
                df0.loc[len(df0)] = {
                    "제조/입고일": default_date,
                    "수량(통)": int(default_qty),
                    "점도(cP)": float(default_visc),
                    "UV흡광도(선택)": float(default_uv) if default_uv_use else None,
                    "비고": default_note
                }
                st.session_state["b_batch_table"] = df0
                st.rerun()
        with cbtn2:
            if st.button("테이블 초기화(1행)", key="b_reset_table"):
                st.session_state["b_batch_table"] = pd.DataFrame([{
                    "제조/입고일": dt.date.today(),
                    "수량(통)": int(default_qty),
                    "점도(cP)": float(default_visc),
                    "UV흡광도(선택)": float(default_uv) if default_uv_use else None,
                    "비고": default_note
                }])
                st.rerun()

        date_bundle_df = st.data_editor(
            st.session_state["b_batch_table"],
            use_container_width=True,
            num_rows="dynamic",
            key="b_batch_editor",
        )
        st.session_state["b_batch_table"] = date_bundle_df

        st.divider()
        submit_batch = st.button("일괄 저장(바인더)", type="primary", key="b_batch_submit")

        if submit_batch:
            visc_lo, visc_hi, uv_hi, rule = get_binder_limits(spec_binder, binder_name)
            m = re.match(r"^([A-Za-z0-9]+)\+YYYYMMDD(-##)?$", str(rule).strip()) if rule else None
            if not m:
                st.error("Spec_Binder의 Lot부여규칙을 해석할 수 없습니다. (예: PCB+YYYYMMDD-## 형태인지 확인 필요)")
                st.stop()

            prefix = m.group(1)
            has_seq = bool(m.group(2))

            existing = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str)
            next_seq_map: dict[str, int] = {}
            rows_to_write: list[dict] = []
            preview: list[dict] = []

            src = st.session_state["b_batch_table"].copy()
            src["제조/입고일"] = src["제조/입고일"].apply(to_date_safe)
            src = src.dropna(subset=["제조/입고일"]).sort_values(by="제조/입고일")

            if src.empty:
                st.warning("저장할 데이터가 없습니다. (제조/입고일이 비어있지 않은지 확인해주세요)")
                st.stop()

            for _, rr in src.iterrows():
                mfg_date = rr["제조/입고일"]
                qty = int(rr.get("수량(통)", 1) or 1)
                qty = max(qty, 1)

                v = rr.get("점도(cP)", None)
                u = rr.get("UV흡광도(선택)", None)
                note = rr.get("비고", "")

                date_str = mfg_date.strftime("%Y%m%d")

                if (not has_seq) and qty > 1:
                    st.error(f"Lot부여규칙에 순번(-##)이 없어 '{mfg_date}' 날짜에서 수량(통)={qty}는 불가합니다. (수량을 1로 입력해주세요)")
                    st.stop()

                if has_seq:
                    if date_str not in next_seq_map:
                        next_seq_map[date_str] = next_seq_for_pattern(existing, prefix, date_str, digits=2, sep="-")
                    start_seq = next_seq_map[date_str]
                else:
                    start_seq = 1

                for i in range(qty):
                    lot = f"{prefix}{date_str}-{(start_seq + i):02d}" if has_seq else f"{prefix}{date_str}"

                    judge_v = judge_range(v, visc_lo, visc_hi)
                    judge_u = judge_range(u, None, uv_hi) if (u is not None and not pd.isna(u)) else None
                    judge = "부적합" if (judge_v == "부적합" or judge_u == "부적합") else "적합"

                    row = {
                        "제조/입고일": mfg_date,
                        "바인더명": binder_name,
                        "Lot(자동)": lot,
                        "점도(cP)": float(v) if (v is not None and not pd.isna(v)) else None,
                        "UV흡광도(선택)": float(u) if (u is not None and not pd.isna(u)) else None,
                        "판정": judge,
                        "비고": str(note) if note is not None else "",
                    }
                    rows_to_write.append(row)
                    preview.append({
                        "제조/입고일": mfg_date,
                        "Lot(자동)": lot,
                        "점도(cP)": row["점도(cP)"],
                        "UV흡광도(선택)": row["UV흡광도(선택)"],
                        "판정": judge,
                        "비고": row["비고"],
                    })

                if has_seq:
                    next_seq_map[date_str] = start_seq + qty

            st.write("저장 미리보기")
            st.dataframe(pd.DataFrame(preview), use_container_width=True)

            try:
                append_rows_to_sheet(xlsx_path, SHEET_BINDER, rows_to_write)
                st.success(f"일괄 저장 완료! (총 {len(rows_to_write)}통)")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    st.divider()

    st.subheader("바인더 입출고 (Google Sheets 자동 반영)")
    st.caption("구글 시트를 수정하면, 새로고침 시 최신 값이 반영됩니다. (캐시 60초)")

    try:
        df_hema = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_HEMA)
        df_sil = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_SIL)
    except Exception as e:
        st.error("구글시트에서 데이터를 못 불러왔습니다. (공유/웹게시/시트명/ID 확인 필요)")
        st.exception(e)
        st.stop()

    df_hema_sorted = try_sort_by_date_desc(df_hema)
    df_sil_sorted = try_sort_by_date_desc(df_sil)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### HEMA (최신순)")
        st.dataframe(df_hema_sorted, use_container_width=True)
    with c2:
        st.markdown("### Silicon (최신순)")
        st.dataframe(df_sil_sorted, use_container_width=True)

    if st.button("지금 최신값으로 다시 불러오기", key="binder_refresh"):
        st.cache_data.clear()
        st.rerun()


# =========================
# Search
# =========================
with tab_search:
    c1, c2, c3 = st.columns([2, 2, 3])
    with c1:
        mode = st.selectbox("검색 종류", ["바인더 Lot", "단일색 잉크 Lot", "제품코드", "색상군", "기간(입고일)"], key="search_mode")
    with c2:
        q = st.text_input("검색어", placeholder="예: PCB20250112-01 / PLB25041501 / PL-835-1 ...", key="search_q")
    with c3:
        st.write("")
        st.caption("💡 바인더 Lot로 검색하면: 바인더 정보 + 해당 바인더를 사용한 단일색 잉크 목록을 같이 보여드립니다.")

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

        if len(s_hit) == 1 and "사용된 바인더 Lot" in s_hit.columns:
            binder_lot = str(s_hit.iloc[0].get("사용된 바인더 Lot", "")).strip()
            if binder_lot:
                b_hit = binder_df[binder_df["Lot(자동)"].astype(str) == binder_lot]
                if len(b_hit):
                    st.subheader("연결된 바인더_제조_입고")
                    st.dataframe(b_hit, use_container_width=True)

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
