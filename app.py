import altair as alt
import streamlit as st
import pandas as pd
import datetime as dt
import re
from pathlib import Path
from openpyxl import load_workbook
import requests
from io import StringIO
from typing import List, Dict, Optional


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


def coerce_date_series(s: pd.Series) -> pd.Series:
    """
    날짜 파싱을 최대한 강하게:
    - 일반 문자열/날짜객체 -> pd.to_datetime
    - 엑셀 날짜 숫자(예: 45234) -> origin=1899-12-30 로 변환
    """
    if s is None:
        return pd.Series([pd.NaT] * 0)

    x = s.copy()

    # 1) 일반 파싱
    dt1 = pd.to_datetime(x, errors="coerce")

    # 2) 엑셀 숫자 날짜 보정(일부만 NaT인 경우도 보완)
    num = pd.to_numeric(x, errors="coerce")
    dt2 = pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")

    return dt1.fillna(dt2)


def coerce_float_series(s: pd.Series) -> pd.Series:
    """
    '45,000' 같이 쉼표 포함/문자 포함 숫자도 안전하게 float로 변환
    """
    if s is None:
        return pd.Series([pd.NA] * 0)
    x = s.copy()
    x = x.astype(str).str.replace(",", "", regex=False).str.strip()
    x = x.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaN": pd.NA})
    return pd.to_numeric(x, errors="coerce")


def safe_minmax_dates(values, fallback_days: int = 90):
    s = pd.to_datetime(values, errors="coerce").dropna()
    today = dt.date.today()
    if len(s) == 0:
        return today - dt.timedelta(days=fallback_days), today
    return s.min().date(), s.max().date()


def delta_e76(lab1, lab2):
    return float(((lab1[0] - lab2[0]) ** 2 + (lab1[1] - lab2[1]) ** 2 + (lab1[2] - lab2[2]) ** 2) ** 0.5)


def extract_delta_e_from_note(note: str) -> Optional[float]:
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
    rule = df["Lot부여규칙"].dropna().iloc[0] if "Lot부여규칙" in df.columns and len(df["Lot부여규칙"].dropna()) else None
    return visc_lo, visc_hi, uv_hi, rule


def parse_binder_rule(rule: Optional[str]):
    if not rule:
        return None, False
    m = re.match(r"^([A-Za-z0-9]+)\+YYYYMMDD(-##)?$", str(rule).strip())
    if not m:
        return None, False
    return m.group(1), bool(m.group(2))


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
st.caption("✅ 대시보드(단일색 요약/추이)  |  ✅ 잉크 입고 입력(엑셀 누적)  |  ✅ 바인더 입출고/업체반환  |  ✅ 빠른검색")


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
    st.sidebar.info("업로드 파일로 실행 중입니다. (이 모드에서는 저장해도 서버에 영구 누적이 보장되지 않습니다.)")

if not Path(xlsx_path).exists():
    st.error(f"엑셀 파일을 찾을 수 없습니다: {xlsx_path}")
    st.stop()

# 업체반환 시트 없으면 자동 생성
ensure_sheet_with_headers(
    xlsx_path,
    SHEET_BINDER_RETURN,
    headers=["반환일자", "바인더명", "관련 Lot(선택)", "반환량(kg)", "비고"]
)

# 데이터 로드
data = load_data(xlsx_path)
binder_df = data["binder"].copy()
single_df = data["single"].copy()
spec_binder = data["spec_binder"].copy()
spec_single = data["spec_single"].copy()
base_lab = data["base_lab"].copy()

# normalize (있으면)
if "제조/입고일" in binder_df.columns:
    binder_df["제조/입고일"] = binder_df["제조/입고일"].apply(normalize_date)
if "입고일" in single_df.columns:
    single_df["입고일"] = single_df["입고일"].apply(normalize_date)

# 업체반환 로드
try:
    binder_return_df = pd.read_excel(xlsx_path, sheet_name=SHEET_BINDER_RETURN).copy()
    if "반환일자" in binder_return_df.columns:
        binder_return_df["반환일자"] = binder_return_df["반환일자"].apply(normalize_date)
except Exception:
    binder_return_df = pd.DataFrame(columns=["반환일자", "바인더명", "관련 Lot(선택)", "반환량(kg)", "비고"])


# =========================
# Tabs
# =========================
tab_dash, tab_ink_in, tab_binder, tab_search = st.tabs(
    ["📊 대시보드", "🧾 잉크 입고", "📦 바인더 입출고", "🔎 빠른검색"]
)


# =========================
# Dashboard (그래프는 이 탭에만)
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

    # -------------------------
    # 1) 단일색 데이터 표
    # -------------------------
    st.subheader("1) 단일색 데이터 목록 (색상군/제품코드/바인더/점도/색차)")
    st.caption("보고용으로 필요한 컬럼만 정리해 한 번에 보여드립니다.")

    s = single_df.copy()

    if "비고" in s.columns:
        s["색차값(ΔE76)"] = s["비고"].apply(extract_delta_e_from_note)
    else:
        s["색차값(ΔE76)"] = None

    # 날짜/점도 강제 변환(표에서도 일관되게)
    if "입고일" in s.columns:
        s["_in_dt"] = coerce_date_series(s["입고일"])
    else:
        s["_in_dt"] = pd.NaT

    if "점도측정값(cP)" in s.columns:
        s["_visc"] = coerce_float_series(s["점도측정값(cP)"])
    else:
        s["_visc"] = pd.NA

    f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
    with f1:
        dmin, dmax = safe_minmax_dates(s["_in_dt"], fallback_days=90)
        start = st.date_input("시작일", value=dmin, key="dash_list_start")
    with f2:
        end = st.date_input("종료일", value=dmax, key="dash_list_end")
    with f3:
        cg = st.multiselect("색상군", sorted(s["색상군"].dropna().unique().tolist()), key="dash_list_cg") if "색상군" in s.columns else []
    with f4:
        pc = st.multiselect("제품코드", sorted(s["제품코드"].dropna().unique().tolist()), key="dash_list_pc") if "제품코드" in s.columns else []

    if start > end:
        start, end = end, start

    s_view = s.copy()
    s_view = s_view.dropna(subset=["_in_dt"])
    s_view = s_view[(s_view["_in_dt"].dt.date >= start) & (s_view["_in_dt"].dt.date <= end)]

    if cg and "색상군" in s_view.columns:
        s_view = s_view[s_view["색상군"].isin(cg)]
    if pc and "제품코드" in s_view.columns:
        s_view = s_view[s_view["제품코드"].isin(pc)]

    show_cols = []
    s_view["제조일자"] = s_view["_in_dt"].dt.date
    show_cols.append("제조일자")
    if "색상군" in s_view.columns:
        show_cols.append("색상군")
    if "제품코드" in s_view.columns:
        show_cols.append("제품코드")
    if "사용된 바인더 Lot" in s_view.columns:
        s_view["사용된바인더"] = s_view["사용된 바인더 Lot"].astype(str)
        show_cols.append("사용된바인더")
    if "점도측정값(cP)" in s_view.columns:
        s_view["점도(cP)"] = s_view["_visc"]
        show_cols.append("점도(cP)")
    show_cols.append("색차값(ΔE76)")

    if len(s_view) == 0:
        st.info("선택한 조건에 해당하는 단일색 데이터가 없습니다.")
    else:
        st.dataframe(
            s_view.sort_values("_in_dt", ascending=False)[show_cols],
            use_container_width=True,
            hide_index=True,
        )

    st.divider()

    # -------------------------
    # 1-2) 색상군별 평균 점도 (점 + 라벨)
    # -------------------------
    st.subheader("색상군별 평균 점도")
    st.caption("각 색상군의 평균 점도를 점으로 표시하고 옆에 값을 표기합니다.")

    if "색상군" in single_df.columns and "점도측정값(cP)" in single_df.columns:
        tmp = single_df.copy()
        tmp["_visc"] = coerce_float_series(tmp["점도측정값(cP)"])
        avg_df = (
            tmp.dropna(subset=["색상군", "_visc"])
            .groupby("색상군", as_index=False)["_visc"]
            .mean()
            .rename(columns={"_visc": "평균점도(cP)"})
        )
        if len(avg_df) == 0:
            st.info("평균을 계산할 점도 데이터가 없습니다. (점도값 형식/쉼표/공백 확인)")
        else:
            avg_df["라벨"] = avg_df["평균점도(cP)"].round(1).astype(str)

            base = alt.Chart(avg_df).encode(
                x=alt.X("색상군:N", sort=sorted(avg_df["색상군"].tolist()), title="색상군"),
                y=alt.Y("평균점도(cP):Q", title="평균 점도(cP)"),
                tooltip=["색상군:N", "평균점도(cP):Q"],
            )
            points = base.mark_point(size=180)
            labels = base.mark_text(dx=8, dy=-8, align="left").encode(text="라벨:N")
            st.altair_chart((points + labels).interactive(), use_container_width=True)
    else:
        st.info("단일색 데이터에 '색상군' 또는 '점도측정값(cP)' 컬럼이 없습니다.")

    st.divider()

    # -------------------------
    # 2) 단일색 점도 변화 추이 (Lot별)
    # -------------------------
    st.subheader("2) 단일색 점도 변화 추이 (Lot별)")
    st.caption("선택한 Lot별로 입고일 기준 점도 변화를 확인합니다. (점 크기/라벨 강화)")

    need_cols = ["입고일", "단일색잉크 Lot", "점도측정값(cP)"]
    miss = [c for c in need_cols if c not in single_df.columns]
    if miss:
        st.warning(f"단일색 데이터에 필요한 컬럼이 없습니다: {miss}")
    else:
        df = single_df.copy()
        df["_in_dt"] = coerce_date_series(df["입고일"])
        df["_visc"] = coerce_float_series(df["점도측정값(cP)"])

        # Lot 정리(빈문자/None 제거)
        df["_lot"] = df["단일색잉크 Lot"].astype(str).str.strip()
        df.loc[df["_lot"].isin(["", "nan", "None", "NaN"]), "_lot"] = pd.NA

        total_n = len(df)
        valid_date_n = int(df["_in_dt"].notna().sum())
        valid_visc_n = int(df["_visc"].notna().sum())
        valid_lot_n = int(df["_lot"].notna().sum())

        df = df.dropna(subset=["_in_dt", "_visc", "_lot"]).copy()
        df = df.sort_values("_in_dt")

        if len(df) == 0:
            st.info("입고일/점도 값이 비어있어 추이 그래프를 표시할 수 없습니다.")
            with st.expander("🔍 데이터 진단(왜 그래프가 안 뜨는지 확인)", expanded=True):
                st.write(f"- 전체 행 수: {total_n}")
                st.write(f"- 날짜 파싱 성공: {valid_date_n}")
                st.write(f"- 점도 숫자 변환 성공: {valid_visc_n}")
                st.write(f"- Lot 값 존재: {valid_lot_n}")
                st.write("아래는 원본 일부(상위 20건)와 파싱 결과입니다.")
                diag = single_df[need_cols].copy().head(20)
                diag["_parsed_date"] = coerce_date_series(diag["입고일"])
                diag["_parsed_visc"] = coerce_float_series(diag["점도측정값(cP)"])
                st.dataframe(diag, use_container_width=True)
        else:
            f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
            with f1:
                dmin, dmax = safe_minmax_dates(df["_in_dt"], fallback_days=90)
                start = st.date_input("시작일(추이)", value=dmin, key="trend_start")
            with f2:
                end = st.date_input("종료일(추이)", value=dmax, key="trend_end")
            with f3:
                cg = st.multiselect("색상군(추이)", sorted(df["색상군"].dropna().unique().tolist()), key="trend_cg") if "색상군" in df.columns else []
            with f4:
                pc = st.multiselect("제품코드(추이)", sorted(df["제품코드"].dropna().unique().tolist()), key="trend_pc") if "제품코드" in df.columns else []

            if start > end:
                start, end = end, start

            df = df[(df["_in_dt"].dt.date >= start) & (df["_in_dt"].dt.date <= end)]
            if cg and "색상군" in df.columns:
                df = df[df["색상군"].isin(cg)]
            if pc and "제품코드" in df.columns:
                df = df[df["제품코드"].isin(pc)]

            lot_list = sorted(df["_lot"].astype(str).unique().tolist())
            default_pick = lot_list[-5:] if len(lot_list) > 5 else lot_list
            pick = st.multiselect("표시할 단일색 Lot(복수 선택)", lot_list, default=default_pick, key="trend_lots")

            if pick:
                df = df[df["_lot"].astype(str).isin(pick)]

            if len(df) == 0:
                st.info("선택한 조건에 해당하는 데이터가 없습니다.")
            else:
                df = df.sort_values("_in_dt")
                df["라벨"] = df["_visc"].round(1).astype(str)

                tooltip_cols = ["_in_dt:T", "_lot:N", "_visc:Q"]
                if "제품코드" in df.columns:
                    tooltip_cols.insert(2, "제품코드:N")
                if "색상군" in df.columns:
                    tooltip_cols.insert(3, "색상군:N")
                if "사용된 바인더 Lot" in df.columns:
                    tooltip_cols.append("사용된 바인더 Lot:N")

                base = alt.Chart(df).encode(
                    x=alt.X("_in_dt:T", sort="ascending", title="입고일"),
                    y=alt.Y("_visc:Q", title="점도(cP)"),
                    color=alt.Color("_lot:N", title="Lot"),
                    tooltip=tooltip_cols,
                )

                line = base.mark_line()
                points = base.mark_point(size=260)  # 점 더 크게
                labels = base.mark_text(dx=10, dy=-12, align="left").encode(text="라벨:N")

                st.altair_chart((line + points + labels).interactive(), use_container_width=True)


# =========================
# 잉크 입고 (단일색 입력만)
# =========================
with tab_ink_in:
    st.subheader("단일색 잉크 입고 입력")
    st.caption("이 탭은 **엑셀 파일에 행을 추가(Append)** 하여 데이터가 누적됩니다.")

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
            st.caption("선택: 착색력(L*a*b*) 입력 시, 기준LAB이 있으면 ΔE(76)을 자동 계산해 비고에 기록합니다.")
            L = st.number_input("착색력_L*", value=0.0, step=0.1, format="%.2f")
            a = st.number_input("착색력_a*", value=0.0, step=0.1, format="%.2f")
            b = st.number_input("착색력_b*", value=0.0, step=0.1, format="%.2f")
            lab_enabled = st.checkbox("L*a*b* 입력함", value=False)

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
                    base = (
                        float(base_hit.iloc[0]["기준_L*"]),
                        float(base_hit.iloc[0]["기준_a*"]),
                        float(base_hit.iloc[0]["기준_b*"])
                    )
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
# 바인더 입출고 + 업체반환 + (구글시트 보기)
# =========================
with tab_binder:
    st.subheader("바인더 입출고 / 업체 반환")
    st.caption("바인더 품질 데이터(제조/입고)와 업체 반환(kg)을 이 탭에서 함께 관리합니다.")

    binder_names = sorted(spec_binder["바인더명"].dropna().unique().tolist())
    binder_lots_all = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).tolist()
    binder_lots_all = sorted(set(binder_lots_all), reverse=True)

    # (0) 업체 반환 입력 (최상단)
    st.markdown("### 0) 바인더 업체 반환 입력 (kg 단위)")
    with st.form("binder_return_form", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([1.2, 1.6, 1.6, 2.6])
        with c1:
            r_date = st.date_input("반환일자", value=dt.date.today(), key="ret_date")
        with c2:
            r_name = st.selectbox("바인더명", binder_names, key="ret_name")
        with c3:
            r_lot = st.selectbox("관련 Lot(선택)", ["(선택안함)"] + binder_lots_all, key="ret_lot")
        with c4:
            r_kg = st.number_input("반환량(kg)", min_value=0.0, step=0.1, format="%.1f", key="ret_kg")

        r_note = st.text_input("비고(선택)", value="", key="ret_note")
        ret_submit = st.form_submit_button("저장(업체반환)", type="primary")

    if ret_submit:
        if r_kg <= 0:
            st.warning("반환량(kg)은 0보다 커야 합니다.")
        else:
            row = {
                "반환일자": r_date,
                "바인더명": r_name,
                "관련 Lot(선택)": "" if r_lot == "(선택안함)" else r_lot,
                "반환량(kg)": float(r_kg),
                "비고": r_note,
            }
            try:
                append_row_to_sheet(xlsx_path, SHEET_BINDER_RETURN, row)
                st.success("업체 반환 입력이 저장되었습니다.")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    with st.expander("업체 반환 내역 보기", expanded=True):
        if len(binder_return_df):
            tmp = binder_return_df.copy()
            tmp["_d"] = pd.to_datetime(tmp.get("반환일자"), errors="coerce")
            tmp = tmp.sort_values("_d", ascending=False).drop(columns=["_d"])
            st.dataframe(tmp, use_container_width=True, hide_index=True)
        else:
            st.info("업체 반환 데이터가 아직 없습니다.")

    st.divider()

    # (2) Google Sheets 보기
    st.markdown("### 1) 바인더 입출고 (Google Sheets 자동 반영)")
    st.caption("구글 시트를 수정하면, 이 화면은 새로고침 시 자동으로 최신 값이 반영됩니다. (캐시 60초)")

    try:
        df_hema = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_HEMA)
        df_sil = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_SIL)
        df_hema = sort_df_by_any_date_col(df_hema)
        df_sil = sort_df_by_any_date_col(df_sil)
    except Exception as e:
        st.error("구글시트에서 데이터를 못 불러왔습니다. 시트 공유/웹게시/시트명/ID를 확인해주세요.")
        st.exception(e)
        st.stop()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### HEMA (최신순)")
        st.dataframe(df_hema, use_container_width=True, hide_index=True)
    with c2:
        st.markdown("#### Silicon (최신순)")
        st.dataframe(df_sil, use_container_width=True, hide_index=True)

    if st.button("지금 최신값으로 다시 불러오기", key="binder_refresh"):
        st.cache_data.clear()
        st.rerun()


# =========================
# Search
# =========================
with tab_search:
    c1, c2, c3 = st.columns([2, 2, 3])
    with c1:
        mode = st.selectbox("검색 종류", ["바인더 Lot", "단일색 잉크 Lot", "제품코드", "색상군", "기간(입고일)"])
    with c2:
        q = st.text_input("검색어", placeholder="예: PCB20250112-01 / PLB25041501 / PL-835-1 ...")
    with c3:
        st.write("")
        st.caption("💡 바인더 Lot로 검색하면: 바인더 정보 + 해당 바인더를 사용한 단일색 잉크 목록을 같이 보여줍니다.")

    if mode == "기간(입고일)":
        d1, d2 = st.columns(2)
        with d1:
            start = st.date_input("시작일", value=dt.date.today() - dt.timedelta(days=30), key="search_start")
        with d2:
            end = st.date_input("종료일", value=dt.date.today(), key="search_end")
        df = single_df.copy()
        if "입고일" in df.columns:
            df["_in_dt"] = coerce_date_series(df["입고일"])
            df = df.dropna(subset=["_in_dt"])
            df = df[df["_in_dt"].dt.date.between(start, end)]
        st.subheader("단일색_수입검사")
        st.dataframe(df.sort_values("_in_dt", ascending=False) if "_in_dt" in df.columns else df, use_container_width=True)

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
