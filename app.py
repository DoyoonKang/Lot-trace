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
SHEET_BINDER_RETURN = "바인더_반품"  # 없으면 자동 생성

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
def _norm_header(x) -> str:
    if x is None:
        return ""
    s = str(x)
    s = s.replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_df_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {_c: _norm_header(_c) for _c in df.columns}
    return df.rename(columns=mapping)


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
    return float(((lab1[0]-lab2[0])**2 + (lab1[1]-lab2[1])**2 + (lab1[2]-lab2[2])**2) ** 0.5)


def parse_deltae_from_note(note: str):
    if note is None or (isinstance(note, float) and pd.isna(note)):
        return None
    s = str(note)
    m = re.search(r"ΔE76\s*=\s*([0-9]+(?:\.[0-9]+)?)", s)
    if not m:
        m = re.search(r"DE76\s*=\s*([0-9]+(?:\.[0-9]+)?)", s, flags=re.IGNORECASE)
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


def ensure_sheet_exists(xlsx_path: str, sheet_name: str, headers: list[str]):
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

    values = []
    for h in headers:
        if h in row:
            values.append(row.get(h, None))
        else:
            values.append(row.get(_norm_header(h), None))
    ws.append(values)
    wb.save(xlsx_path)


def append_rows_to_sheet(xlsx_path: str, sheet_name: str, rows: list[dict]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]

    for row in rows:
        values = []
        for h in headers:
            if h in row:
                values.append(row.get(h, None))
            else:
                values.append(row.get(_norm_header(h), None))
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


@st.cache_data(show_spinner=False)
def load_data(xlsx_path: str) -> dict[str, pd.DataFrame]:
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


def safe_date_bounds(dts: pd.Series, fallback_days: int = 90):
    dts = pd.to_datetime(dts, errors="coerce")
    dts = dts.dropna()
    today = dt.date.today()
    if len(dts) == 0:
        return today - dt.timedelta(days=fallback_days), today
    return dts.min().date(), dts.max().date()


def coerce_numeric(s: pd.Series):
    return pd.to_numeric(s, errors="coerce")


# =========================
# UI Header
# =========================
st.title("액상 잉크 Lot 추적 관리 대시보드")
st.caption("✅ 빠른 검색 · ✅ 잉크 입고 등록(엑셀 누적) · ✅ 대시보드(단일색 평균/추이) · ✅ 바인더 입출고(구글시트 연동)")


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
    st.sidebar.warning("업로드 파일로 실행 중입니다. 이 모드에서는 저장해도 서버에 영구 누적이 보장되지 않습니다.")

if not Path(xlsx_path).exists():
    st.error(f"엑셀 파일을 찾을 수 없습니다: {xlsx_path}")
    st.stop()

data = load_data(xlsx_path)

binder_df = normalize_df_columns(data["binder"]).copy()
single_df = normalize_df_columns(data["single"]).copy()
spec_binder = normalize_df_columns(data["spec_binder"]).copy()
spec_single = normalize_df_columns(data["spec_single"]).copy()
base_lab = normalize_df_columns(data["base_lab"]).copy()

# 날짜 정규화
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
# Dashboard (✅ 그래프/표는 여기(첫 탭)에만)
# =========================
with tab_dash:
    # KPI
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

    # 1) 테이블
    st.subheader("1) 단일색 데이터 (엑셀 형태)")
    st.caption("제조일자(입고일)·색상군·제품코드·사용된 바인더·점도·색차값(ΔE) 기준으로 한눈에 보이도록 정리했습니다.")

    req_cols = ["입고일", "색상군", "제품코드", "사용된 바인더 Lot", "점도측정값(cP)", "비고"]
    miss = [c for c in req_cols if c not in single_df.columns]
    if miss:
        st.warning(f"단일색 시트에서 필요한 컬럼을 찾지 못했습니다: {miss}")
    else:
        view = single_df[req_cols].copy()
        view["점도측정값(cP)"] = coerce_numeric(view["점도측정값(cP)"])
        view["색차값(ΔE76)"] = view["비고"].apply(parse_deltae_from_note)

        view = view.rename(columns={
            "입고일": "제조일자/입고일",
            "사용된 바인더 Lot": "사용된바인더",
            "점도측정값(cP)": "점도(cP)",
        })[["제조일자/입고일", "색상군", "제품코드", "사용된바인더", "점도(cP)", "색차값(ΔE76)"]]

        dmin, dmax = safe_date_bounds(view["제조일자/입고일"])
        f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
        with f1:
            start = st.date_input("시작일(테이블)", value=max(dmin, dmax - dt.timedelta(days=90)), key="tbl_start")
        with f2:
            end = st.date_input("종료일(테이블)", value=dmax, key="tbl_end")
        with f3:
            cg_list = sorted(view["색상군"].dropna().astype(str).unique().tolist())
            cg = st.multiselect("색상군", cg_list, key="tbl_cg")
        with f4:
            pc_list = sorted(view["제품코드"].dropna().astype(str).unique().tolist())
            pc = st.multiselect("제품코드", pc_list, key="tbl_pc")

        if start > end:
            start, end = end, start

        v2 = view.copy()
        v2["제조일자/입고일"] = pd.to_datetime(v2["제조일자/입고일"], errors="coerce")
        v2 = v2.dropna(subset=["제조일자/입고일"])
        v2 = v2[(v2["제조일자/입고일"].dt.date >= start) & (v2["제조일자/입고일"].dt.date <= end)]
        if cg:
            v2 = v2[v2["색상군"].astype(str).isin([str(x) for x in cg])]
        if pc:
            v2 = v2[v2["제품코드"].astype(str).isin([str(x) for x in pc])]

        v2 = v2.sort_values("제조일자/입고일", ascending=False)
        st.dataframe(v2, use_container_width=True, height=320)

    st.divider()

    # 1-2) 평균 점도 (점 + 라벨)
    st.subheader("색상군별 평균 점도")
    st.caption("막대 대신 점으로 표시하고, 옆에 평균 점도 값을 함께 표기했습니다.")

    if "색상군" in single_df.columns and "점도측정값(cP)" in single_df.columns:
        mean_df = single_df[["색상군", "점도측정값(cP)"]].copy()
        mean_df["점도측정값(cP)"] = coerce_numeric(mean_df["점도측정값(cP)"])
        mean_df = mean_df.dropna(subset=["색상군", "점도측정값(cP)"])
        mean_df = mean_df.groupby("색상군", as_index=False)["점도측정값(cP)"].mean()
        mean_df = mean_df.rename(columns={"점도측정값(cP)": "평균점도(cP)"})

        base = alt.Chart(mean_df).encode(
            y=alt.Y("색상군:N", sort="-x", title="색상군"),
            x=alt.X("평균점도(cP):Q", title="평균 점도(cP)"),
            tooltip=["색상군:N", alt.Tooltip("평균점도(cP):Q", format=",.0f")]
        )
        pts = base.mark_point(size=220)
        txt = base.mark_text(align="left", dx=8, baseline="middle").encode(
            text=alt.Text("평균점도(cP):Q", format=",.0f")
        )
        st.altair_chart((pts + txt), use_container_width=True)
    else:
        st.info("단일색 데이터에 '색상군' 또는 '점도측정값(cP)' 컬럼이 없습니다.")

    st.divider()

    # 2) 추이 (Lot별) - 점 크게 + 라벨
    st.subheader("2) 단일색 점도 변화 추이 (Lot별)")
    st.caption("선택한 Lot별로 입고일 기준 점도 변화를 확인합니다. (점 크기/라벨 강화)")

    need_cols = ["입고일", "단일색잉크 Lot", "점도측정값(cP)"]
    miss = [c for c in need_cols if c not in single_df.columns]
    if miss:
        st.warning(f"단일색 데이터에 필요한 컬럼이 없습니다: {miss}")
    else:
        extra = [c for c in ["색상군", "제품코드", "사용된 바인더 Lot"] if c in single_df.columns]
        df = single_df[need_cols + extra].copy()

        df["입고일"] = pd.to_datetime(df["입고일"], errors="coerce")
        df["점도측정값(cP)"] = coerce_numeric(df["점도측정값(cP)"])
        df["단일색잉크 Lot"] = df["단일색잉크 Lot"].astype(str).str.strip()

        df = df.dropna(subset=["입고일", "점도측정값(cP)"])
        df = df[df["단일색잉크 Lot"].ne("") & df["단일색잉크 Lot"].ne("nan")]

        if len(df) == 0:
            st.info("입고일/점도/Lot 값이 비어 있어 추이 그래프를 표시할 수 없습니다.")
        else:
            dmin, dmax = safe_date_bounds(df["입고일"])
            f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
            with f1:
                start = st.date_input("시작일(추이)", value=max(dmin, dmax - dt.timedelta(days=90)), key="trend_start")
            with f2:
                end = st.date_input("종료일(추이)", value=dmax, key="trend_end")
            with f3:
                if "색상군" in df.columns:
                    cg_list = sorted(df["색상군"].dropna().astype(str).unique().tolist())
                    cg = st.multiselect("색상군(추이)", cg_list, key="trend_cg")
                else:
                    cg = []
            with f4:
                if "제품코드" in df.columns:
                    pc_list = sorted(df["제품코드"].dropna().astype(str).unique().tolist())
                    pc = st.multiselect("제품코드(추이)", pc_list, key="trend_pc")
                else:
                    pc = []

            if start > end:
                start, end = end, start

            df2 = df[(df["입고일"].dt.date >= start) & (df["입고일"].dt.date <= end)].copy()
            if cg and "색상군" in df2.columns:
                df2 = df2[df2["색상군"].astype(str).isin([str(x) for x in cg])]
            if pc and "제품코드" in df2.columns:
                df2 = df2[df2["제품코드"].astype(str).isin([str(x) for x in pc])]

            lot_list = sorted(df2["단일색잉크 Lot"].dropna().astype(str).unique().tolist())
            default_pick = lot_list[-5:] if len(lot_list) > 5 else lot_list
            pick = st.multiselect("표시할 단일색 Lot(복수 선택)", lot_list, default=default_pick, key="trend_lots")
            if pick:
                df2 = df2[df2["단일색잉크 Lot"].astype(str).isin([str(x) for x in pick])]

            if len(df2) == 0:
                st.info("선택한 조건에 해당하는 데이터가 없습니다.")
            else:
                df2 = df2.sort_values("입고일")

                tooltip_cols = ["입고일:T", "단일색잉크 Lot:N", alt.Tooltip("점도측정값(cP):Q", format=",.0f")]
                if "제품코드" in df2.columns:
                    tooltip_cols.insert(2, "제품코드:N")
                if "색상군" in df2.columns:
                    tooltip_cols.insert(3, "색상군:N")
                if "사용된 바인더 Lot" in df2.columns:
                    tooltip_cols.append("사용된 바인더 Lot:N")

                base = alt.Chart(df2).encode(
                    x=alt.X("입고일:T", title="입고일"),
                    y=alt.Y("점도측정값(cP):Q", title="점도(cP)"),
                    color=alt.Color("단일색잉크 Lot:N", title="Lot"),
                    tooltip=tooltip_cols,
                )

                line = base.mark_line(strokeWidth=2)
                points = base.mark_point(size=160)
                labels = base.mark_text(align="left", dx=8, dy=-6).encode(
                    text=alt.Text("점도측정값(cP):Q", format=",.0f")
                )

                st.altair_chart((line + points + labels).interactive(), use_container_width=True)

    st.divider()

    st.subheader("최근 20건 (단일색)")
    if "입고일" in single_df.columns:
        show = single_df.sort_values(by="입고일", ascending=False).head(20)
    else:
        show = single_df.head(20)
    st.dataframe(show, use_container_width=True)


# =========================
# Ink inbound (단일색만)
# =========================
with tab_ink_in:
    st.subheader("잉크 입고 등록 (단일색)")
    st.caption("입고 정보를 입력하면 엑셀에 누적 저장되고, 대시보드에 자동 반영됩니다.")

    ink_types = ["HEMA", "Silicone"]
    color_groups = sorted(spec_single["색상군"].dropna().astype(str).unique().tolist()) if "색상군" in spec_single.columns else []
    product_codes = sorted(spec_single["제품코드"].dropna().astype(str).unique().tolist()) if "제품코드" in spec_single.columns else []

    binder_lots = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).tolist()
    binder_lots = sorted(set([x.strip() for x in binder_lots if str(x).strip()]), reverse=True)

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
            st.caption("선택: 착색력(L*a*b*) 입력 시, 기준LAB이 있으면 ΔE(76)을 자동 계산하여 비고에 기록합니다.")
            L = st.number_input("착색력_L*", value=0.0, step=0.1, format="%.2f", key="single_L")
            a = st.number_input("착색력_a*", value=0.0, step=0.1, format="%.2f", key="single_a")
            b = st.number_input("착색력_b*", value=0.0, step=0.1, format="%.2f", key="single_b")
            lab_enabled = st.checkbox("L*a*b* 입력함", value=False, key="single_lab_en")

        note = st.text_input("비고", value="", key="single_note")
        submit_s = st.form_submit_button("저장(단일색)")

    if submit_s:
        binder_type = infer_binder_type_from_lot(spec_binder, binder_lot)

        spec_hit = spec_single.copy()
        if "색상군" in spec_hit.columns:
            spec_hit = spec_hit[spec_hit["색상군"].astype(str) == str(color_group)]
        if "제품코드" in spec_hit.columns:
            spec_hit = spec_hit[spec_hit["제품코드"].astype(str) == str(product_code)]
        if binder_type and "BinderType" in spec_hit.columns:
            spec_hit = spec_hit[spec_hit["BinderType"].astype(str) == str(binder_type)]

        if len(spec_hit) == 0 or "하한" not in spec_hit.columns or "상한" not in spec_hit.columns:
            lo, hi = None, None
            visc_judge = None
            st.warning("점도 기준을 Spec_Single_H&S에서 찾지 못했습니다. (색상군/제품코드/바인더타입 조합 확인)")
        else:
            lo = float(spec_hit["하한"].iloc[0]) if pd.notna(spec_hit["하한"].iloc[0]) else None
            hi = float(spec_hit["상한"].iloc[0]) if pd.notna(spec_hit["상한"].iloc[0]) else None
            visc_judge = judge_range(visc_meas, lo, hi)

        new_lot = generate_single_lot(single_df, product_code, color_group, in_date)
        if new_lot is None:
            st.error("단일색 Lot 자동 생성에 실패했습니다. (색상군 매핑 확인 필요)")
        else:
            note2 = note
            if lab_enabled:
                base_hit = base_lab[base_lab["제품코드"].astype(str) == str(product_code)] if "제품코드" in base_lab.columns else base_lab.iloc[0:0]
                if len(base_hit) == 1 and {"기준_L*", "기준_a*", "기준_b*"}.issubset(set(base_hit.columns)):
                    base = (float(base_hit.iloc[0]["기준_L*"]), float(base_hit.iloc[0]["기준_a*"]), float(base_hit.iloc[0]["기준_b*"]))
                    de = delta_e76((L, a, b), base)
                    note2 = (note2 + " " if note2 else "") + f"[ΔE76={de:.2f}]"
                else:
                    note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"

            row = {
                "입고일": in_date,
                "잉크타입 (HEMA/Silicone)": ink_type,
                "색상군": color_group,
                "제품코드": product_code,
                "단일색잉크 Lot": new_lot,
                "사용된 바인더 Lot": binder_lot,
                "바인더제조처 (내부/외주)": supplier,
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
# Binder IO
# =========================
with tab_binder:
    st.subheader("바인더 입출고")

    t_return, t_visc, t_view = st.tabs(["🔁 반품(업체반환) 입력", "🧪 바인더 점도 입력", "📄 입출고 현황(구글시트)"])

    # (1) 반품 입력 (kg)
    with t_return:
        st.caption("바인더는 1통(20kg) 기준으로 사용 후 남은 kg 단위로 업체에 반환하는 경우를 기록합니다.")
        ensure_sheet_exists(
            xlsx_path,
            SHEET_BINDER_RETURN,
            headers=["반품일자", "바인더명", "관련Lot(선택)", "반품kg", "비고"]
        )

        binder_names = sorted(spec_binder["바인더명"].dropna().astype(str).unique().tolist()) if "바인더명" in spec_binder.columns else []

        with st.form("binder_return_form", clear_on_submit=True):
            c1, c2, c3, c4 = st.columns([1.2, 1.6, 1.6, 1.2])
            with c1:
                r_date = st.date_input("반품일자", value=dt.date.today(), key="ret_date")
            with c2:
                r_name = st.selectbox("바인더명", binder_names, key="ret_name")
            with c3:
                r_lot = st.text_input("관련 Lot(선택)", value="", key="ret_lot")
            with c4:
                r_kg = st.number_input("반품 kg", min_value=0.0, step=0.1, format="%.1f", key="ret_kg")
            r_note = st.text_input("비고", value="", key="ret_note")
            submit_ret = st.form_submit_button("저장(반품)")

        if submit_ret:
            if r_kg <= 0:
                st.error("반품 kg는 0보다 커야 합니다.")
            else:
                row = {
                    "반품일자": r_date,
                    "바인더명": r_name,
                    "관련Lot(선택)": r_lot,
                    "반품kg": float(r_kg),
                    "비고": r_note,
                }
                try:
                    append_row_to_sheet(xlsx_path, SHEET_BINDER_RETURN, row)
                    st.success("반품 기록 저장 완료!")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"저장 실패: {e}")

        try:
            ret_df = normalize_df_columns(pd.read_excel(xlsx_path, sheet_name=SHEET_BINDER_RETURN))
            if "반품일자" in ret_df.columns:
                ret_df["반품일자"] = pd.to_datetime(ret_df["반품일자"], errors="coerce")
                ret_df = ret_df.sort_values("반품일자", ascending=False)
            st.subheader("최근 반품 기록")
            st.dataframe(ret_df.head(30), use_container_width=True)
        except Exception:
            pass

    # (2) 바인더 점도 입력 (여러 날짜/수량 일괄)
    with t_visc:
        st.caption("여러 날짜에 걸쳐 들어온 바인더 Lot들을 한 번에 입력할 수 있습니다.")

        binder_names = sorted(spec_binder["바인더명"].dropna().astype(str).unique().tolist()) if "바인더명" in spec_binder.columns else []
        binder_name = st.selectbox("바인더명", binder_names, key="b_batch_name")

        st.markdown("#### 일괄 입력 표")
        st.caption("각 행: 제조/입고일 + 수량(통) + 점도/UV + 비고 (예: 3일치가 한 번에 들어온 경우 3줄로 입력)")

        base_rows = pd.DataFrame([
            {"제조/입고일": dt.date.today(), "수량(통)": 1, "점도(cP)": 0.0, "UV흡광도(선택)": None, "비고": ""}
        ])
        edit_df = st.data_editor(
            base_rows,
            use_container_width=True,
            num_rows="dynamic",
            key="b_batch_editor",
            column_config={
                "제조/입고일": st.column_config.DateColumn("제조/입고일"),
                "수량(통)": st.column_config.NumberColumn("수량(통)", min_value=1, max_value=100, step=1),
                "점도(cP)": st.column_config.NumberColumn("점도(cP)", min_value=0.0, step=1.0, format="%.1f"),
                "UV흡광도(선택)": st.column_config.NumberColumn("UV흡광도(선택)", min_value=0.0, step=0.01, format="%.3f"),
            }
        )

        uv_enabled = st.checkbox("UV 값도 저장(입력된 값이 있을 때만)", value=False, key="b_uv_en")
        submit_batch = st.button("일괄 저장(바인더)", type="primary", key="b_batch_submit")

        if submit_batch:
            if edit_df is None or len(edit_df) == 0:
                st.error("입력 표에 데이터가 없습니다.")
                st.stop()

            visc_lo, visc_hi, uv_hi, rule = get_binder_limits(spec_binder, binder_name)
            m = re.match(r"^([A-Za-z0-9]+)\+YYYYMMDD(-##)?$", str(rule).strip()) if rule else None
            if not m:
                st.error("Spec_Binder의 Lot부여규칙을 해석할 수 없습니다. (예: PCB+YYYYMMDD-## 형태인지 확인 필요)")
                st.stop()

            prefix = m.group(1)
            has_seq = bool(m.group(2))
            if not has_seq:
                st.warning("Lot부여규칙에 순번(-##)이 없습니다. 같은 날짜에 여러 통을 넣으면 Lot가 중복될 수 있습니다.")

            rows = []
            preview = []
            existing = binder_df.get("Lot(자동)", pd.Series(dtype=str))

            for _, r in edit_df.iterrows():
                mfg_date = normalize_date(r.get("제조/입고일"))
                if mfg_date is None:
                    continue

                qty = int(r.get("수량(통)") or 0)
                if qty <= 0:
                    continue

                v = float(r.get("점도(cP)") or 0.0)
                u_raw = r.get("UV흡광도(선택)")
                u = float(u_raw) if (uv_enabled and pd.notna(u_raw)) else None
                note = str(r.get("비고") or "")

                date_str = mfg_date.strftime("%Y%m%d")
                start_seq = next_seq_for_pattern(existing, prefix, date_str, digits=2, sep="-")

                for i in range(qty):
                    lot = f"{prefix}{date_str}-{(start_seq + i):02d}" if has_seq else f"{prefix}{date_str}"

                    judge_v = judge_range(v, visc_lo, visc_hi)
                    judge_u = judge_range(u, None, uv_hi) if uv_enabled else None
                    judge = "부적합" if (judge_v == "부적합" or judge_u == "부적합") else "적합"

                    row = {
                        "제조/입고일": mfg_date,
                        "바인더명": binder_name,
                        "Lot(자동)": lot,
                        "점도(cP)": v,
                        "UV흡광도(선택)": u,
                        "판정": judge,
                        "비고": note,
                    }
                    rows.append(row)
                    preview.append({
                        "제조/입고일": mfg_date,
                        "Lot(자동)": lot,
                        "점도(cP)": v,
                        "UV흡광도(선택)": u,
                        "판정": judge,
                    })

                existing = pd.concat([existing, pd.Series([x["Lot(자동)"] for x in rows[-qty:]])], ignore_index=True)

            if len(rows) == 0:
                st.error("저장할 행이 없습니다. 날짜/수량/점도 입력을 확인해주세요.")
                st.stop()

            st.write("저장 미리보기(상위 50)")
            st.dataframe(pd.DataFrame(preview).head(50), use_container_width=True)

            try:
                append_rows_to_sheet(xlsx_path, SHEET_BINDER, rows)
                st.success(f"일괄 저장 완료! 총 {len(rows)}건")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

        st.divider()
        st.subheader("최근 바인더 점도 기록(50)")
        if "제조/입고일" in binder_df.columns:
            tmp = binder_df.copy()
            tmp["제조/입고일"] = pd.to_datetime(tmp["제조/입고일"], errors="coerce")
            tmp = tmp.sort_values("제조/입고일", ascending=False)
            st.dataframe(tmp.head(50), use_container_width=True)
        else:
            st.dataframe(binder_df.head(50), use_container_width=True)

    # (3) 구글시트 보기 (최신순)
    with t_view:
        st.caption("구글 시트를 수정하면 새로고침 시 자동으로 최신 값이 반영됩니다. (캐시 60초)")

        try:
            df_hema = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_HEMA)
            df_sil = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_SIL)
        except Exception as e:
            st.error("구글시트에서 데이터를 못 불러왔습니다. 시트 공유/웹게시/시트명/ID를 확인해주세요.")
            st.exception(e)
            st.stop()

        def sort_latest_first(df: pd.DataFrame):
            d = df.copy()
            candidates = [c for c in d.columns if any(k in str(c) for k in ["일자", "날짜", "date", "Date"])]
            for c in candidates:
                tmp = pd.to_datetime(d[c], errors="coerce")
                if tmp.notna().sum() >= max(3, int(len(d)*0.3)):
                    d["_sort_date"] = tmp
                    d = d.sort_values("_sort_date", ascending=False).drop(columns=["_sort_date"])
                    return d
            return d

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### HEMA")
            st.dataframe(sort_latest_first(df_hema), use_container_width=True, height=420)
        with c2:
            st.markdown("### Silicon")
            st.dataframe(sort_latest_first(df_sil), use_container_width=True, height=420)

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
        st.caption("💡 바인더 Lot로 검색하면: 바인더 정보 + 해당 바인더를 사용한 단일색 잉크 목록을 같이 보여줍니다.")

    if mode == "기간(입고일)":
        d1, d2 = st.columns(2)
        with d1:
            start = st.date_input("시작일", value=dt.date.today() - dt.timedelta(days=30), key="search_start")
        with d2:
            end = st.date_input("종료일", value=dt.date.today(), key="search_end")

        df = single_df.copy()
        if "입고일" in df.columns:
            df["입고일"] = pd.to_datetime(df["입고일"], errors="coerce")
            df = df.dropna(subset=["입고일"])
            df = df[(df["입고일"].dt.date >= start) & (df["입고일"].dt.date <= end)]
        st.subheader("단일색_수입검사")
        st.dataframe(df.sort_values(by="입고일", ascending=False), use_container_width=True)

    elif mode == "바인더 Lot":
        b = binder_df.copy()
        b_hit = df_quick_filter(b, q, ["Lot(자동)", "바인더명", "비고"])
        st.subheader("바인더_제조_입고")
        if "제조/입고일" in b_hit.columns:
            b_hit["제조/입고일"] = pd.to_datetime(b_hit["제조/입고일"], errors="coerce")
            st.dataframe(b_hit.sort_values(by="제조/입고일", ascending=False), use_container_width=True)
        else:
            st.dataframe(b_hit, use_container_width=True)

        if q and "사용된 바인더 Lot" in single_df.columns:
            s_hit = single_df[single_df["사용된 바인더 Lot"].astype(str).str.contains(str(q).strip(), case=False, na=False)]
            st.subheader("연결된 단일색_수입검사 (사용된 바인더 Lot)")
            if "입고일" in s_hit.columns:
                s_hit["입고일"] = pd.to_datetime(s_hit["입고일"], errors="coerce")
                st.dataframe(s_hit.sort_values(by="입고일", ascending=False), use_container_width=True)
            else:
                st.dataframe(s_hit, use_container_width=True)

    elif mode == "단일색 잉크 Lot":
        s = single_df.copy()
        s_hit = df_quick_filter(s, q, ["단일색잉크 Lot", "제품코드", "사용된 바인더 Lot", "색상군", "비고"])
        st.subheader("단일색_수입검사")
        if "입고일" in s_hit.columns:
            s_hit["입고일"] = pd.to_datetime(s_hit["입고일"], errors="coerce")
            st.dataframe(s_hit.sort_values(by="입고일", ascending=False), use_container_width=True)
        else:
            st.dataframe(s_hit, use_container_width=True)

        if len(s_hit) == 1 and "사용된 바인더 Lot" in s_hit.columns and "Lot(자동)" in binder_df.columns:
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
            s_hit["입고일"] = pd.to_datetime(s_hit["입고일"], errors="coerce")
            st.dataframe(s_hit.sort_values(by="입고일", ascending=False), use_container_width=True)
        else:
            st.dataframe(s_hit, use_container_width=True)

    elif mode == "색상군":
        s = single_df.copy()
        s_hit = df_quick_filter(s, q, ["색상군"])
        st.subheader("단일색_수입검사")
        if "입고일" in s_hit.columns:
            s_hit["입고일"] = pd.to_datetime(s_hit["입고일"], errors="coerce")
            st.dataframe(s_hit.sort_values(by="입고일", ascending=False), use_container_width=True)
        else:
            st.dataframe(s_hit, use_container_width=True)
