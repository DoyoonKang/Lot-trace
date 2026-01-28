import streamlit as st
import pandas as pd
import altair as alt
import datetime as dt
import re
from pathlib import Path
from io import StringIO
import requests
from openpyxl import load_workbook

# ==========================================================
# Page
# ==========================================================
st.set_page_config(page_title="액상 잉크 Lot 추적 관리", page_icon="🧪", layout="wide")

st.markdown(
    """
    <style>
      .block-container { padding-top: 1.1rem; padding-bottom: 1.8rem; }
      .section-title { font-size: 1.15rem; font-weight: 800; margin: 0.2rem 0 0.2rem 0; }
      .section-sub { color: rgba(49,51,63,0.65); font-size: 0.92rem; margin-bottom: 0.6rem; }
      .kpi-note { color: rgba(49,51,63,0.70); font-size: 0.85rem; margin-top: -0.2rem; }
      div[data-testid="stExpander"] > details > summary { font-weight: 700; }
    </style>
    """,
    unsafe_allow_html=True
)

# ==========================================================
# Config
# ==========================================================
DEFAULT_XLSX = "액상잉크_Lot추적관리_FINAL.xlsx"
DEFAULT_STOCK_XLSX = "액상 재고조사표_자동계산 (12).xlsx"

SHEET_BINDER = "바인더_제조_입고"
SHEET_SINGLE = "단일색_수입검사"
SHEET_SPEC_BINDER = "Spec_Binder"
SHEET_SPEC_SINGLE = "Spec_Single_H&S"
SHEET_BASE_LAB = "기준LAB"
SHEET_BINDER_RETURN = "바인더_업체반환"  # 없으면 자동 생성

# 바인더 입출고(구글시트)
BINDER_SHEET_ID = "1H2fFxnf5AvpSlu-uoZ4NpTv8LYLNwTNAzvlntRQ7FS8"
BINDER_SHEET_HEMA = "HEMA 바인더 입출고 관리대장"
BINDER_SHEET_SIL = "Silicon바인더 입출고 관리대장"


# ==========================================================
# Helpers
# ==========================================================
def norm_key(x) -> str:
    if x is None:
        return ""
    s = str(x).replace("\n", " ").replace("\r", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def find_col(df: pd.DataFrame, want: str):
    w = norm_key(want)
    for c in df.columns:
        if norm_key(c) == w:
            return c
    return None

def safe_to_float(x):
    if x is None:
        return None
    if isinstance(x, float) and pd.isna(x):
        return None
    if isinstance(x, str) and x.strip() == "":
        return None
    try:
        if isinstance(x, str):
            x = x.replace(",", "")
        return float(x)
    except Exception:
        return None

def normalize_date(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    if isinstance(x, (dt.date, dt.datetime)):
        return x.date() if isinstance(x, dt.datetime) else x
    try:
        return pd.to_datetime(x).date()
    except Exception:
        return None

def ensure_sheet_exists(xlsx_path: str, sheet_name: str, headers: list[str]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        wb.save(xlsx_path)

def add_excel_row_number(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_excel_row"] = df.index + 2  # 헤더 1행 가정
    return df

def safe_date_bounds(series: pd.Series):
    s = pd.to_datetime(series, errors="coerce").dropna()
    if len(s) == 0:
        today = dt.date.today()
        return today, today
    return s.min().date(), s.max().date()

def detect_date_col(df: pd.DataFrame):
    for c in df.columns:
        ck = norm_key(c)
        if any(k in ck.lower() for k in ["일자", "날짜", "date", "입고일", "출고일"]):
            return c
    return None

# ==========================================================
# Color/Stock helpers  (요청 반영: 화면에 BLACK/RED 등 대문자 표시)
# ==========================================================
COLOR_KEYS = ["BLACK","BLUE","GREEN","YELLOW","RED","PINK","WHITE","OTHER"]

def normalize_color_group(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "OTHER"
    s = str(x).strip()
    if not s or s.lower() in ("nan", "none"):
        return "OTHER"

    u = s.upper()
    # 한국어/영문 혼용 대응
    if "BLACK" in u or "검정" in s or "흑" in s:
        return "BLACK"
    if "WHITE" in u or "흰" in s or "백" in s:
        return "WHITE"
    if "RED" in u or "빨" in s or "적" in s:
        return "RED"
    if "YELLOW" in u or "노" in s or "황" in s or "옐" in s:
        return "YELLOW"
    if "GREEN" in u or "초" in s or "녹" in s:
        return "GREEN"
    if "BLUE" in u or "파" in s or "청" in s:
        return "BLUE"
    if "PINK" in u or "핑" in s:
        return "PINK"

    if u in COLOR_KEYS:
        return u
    return "OTHER"

def normalize_product_code(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip()
    if not s or s.lower() in ("nan", "none"):
        return ""
    s = s.replace("–", "-").replace("—", "-").replace("−", "-")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace("(액상잉크)", "").replace("액상잉크", "").strip()
    return s

def build_product_to_color_map(spec_single: pd.DataFrame, single_df: pd.DataFrame) -> dict[str, str]:
    mapping: dict[str, str] = {}

    sp_pc = find_col(spec_single, "제품코드")
    sp_cg = find_col(spec_single, "색상군")
    if sp_pc and sp_cg and len(spec_single):
        tmp = spec_single[[sp_pc, sp_cg]].dropna()
        tmp[sp_pc] = tmp[sp_pc].apply(normalize_product_code)
        tmp[sp_cg] = tmp[sp_cg].apply(normalize_color_group)
        for pc, g in tmp.groupby(sp_pc)[sp_cg]:
            mapping[str(pc)] = g.value_counts().idxmax()

    s_pc = find_col(single_df, "제품코드")
    s_cg = find_col(single_df, "색상군")
    if s_pc and s_cg and len(single_df):
        tmp = single_df[[s_pc, s_cg]].dropna()
        tmp[s_pc] = tmp[s_pc].apply(normalize_product_code)
        tmp[s_cg] = tmp[s_cg].apply(normalize_color_group)
        for pc, g in tmp.groupby(s_pc)[s_cg]:
            pc = str(pc)
            if pc not in mapping:
                mapping[pc] = g.value_counts().idxmax()

    return mapping

def _parse_stock_sheet_date(sheet_name: str, today: dt.date):
    s = str(sheet_name).strip()
    m = re.match(r"^(\d{1,2})\.(\d{1,2})$", s)  # 예: 1.15
    if not m:
        return None
    month = int(m.group(1))
    day = int(m.group(2))
    year = today.year
    if month > (today.month + 1):
        year -= 1
    try:
        return dt.date(year, month, day)
    except ValueError:
        return None

@st.cache_data(show_spinner=False)
def load_stock_history(stock_xlsx_path: str, product_to_color: dict[str, str]) -> pd.DataFrame:
    p = Path(stock_xlsx_path)
    if not stock_xlsx_path or not p.exists():
        return pd.DataFrame()

    today = dt.date.today()
    xls = pd.ExcelFile(stock_xlsx_path, engine="openpyxl")

    frames = []
    for sh in xls.sheet_names:
        d = _parse_stock_sheet_date(sh, today)
        if d is None:
            continue

        df = pd.read_excel(xls, sheet_name=sh)
        df = df.rename(columns=lambda x: str(x).strip())

        c_div = find_col(df, "구분")
        c_item = find_col(df, "품목명")
        c_curr = find_col(df, "금일 재고(kg)") or find_col(df, "금일재고(kg)") or find_col(df, "재고(kg)")
        c_used = find_col(df, "하루 사용량(kg)") or find_col(df, "사용량(kg)") or find_col(df, "사용량")

        if not (c_item and c_curr and c_used):
            continue

        df["_division"] = df[c_div].astype(str).str.strip() if c_div else ""
        df["_product"] = df[c_item].apply(normalize_product_code)
        df["_curr"] = pd.to_numeric(df[c_curr].astype(str).str.replace(",", "", regex=False), errors="coerce")
        df["_used_raw"] = pd.to_numeric(df[c_used].astype(str).str.replace(",", "", regex=False), errors="coerce")

        df = df.dropna(subset=["_product", "_curr"])
        df["used_kg"] = df["_used_raw"].clip(lower=0).fillna(0)
        df["inbound_kg"] = (-df["_used_raw"]).clip(lower=0).fillna(0)
        df["inbound_event"] = (df["inbound_kg"] > 0).astype(int)
        df["curr_stock_kg"] = df["_curr"].fillna(0)

        df["color_group"] = df["_product"].map(product_to_color).fillna("OTHER").apply(normalize_color_group)
        df["date"] = pd.to_datetime(d)

        frames.append(df[["date","_division","_product","color_group","curr_stock_kg","used_kg","inbound_kg","inbound_event"]])

    if not frames:
        return pd.DataFrame()

    hist = pd.concat(frames, ignore_index=True)
    hist = hist.rename(columns={"_division":"division", "_product":"product_code"})
    hist = hist.sort_values(["date","division","product_code"]).reset_index(drop=True)
    return hist

def _color_scale_color_group():
    # 도메인은 반드시 데이터와 동일해야 함(대문자)
    domain = ["BLACK","BLUE","GREEN","YELLOW","RED","PINK","WHITE","OTHER"]
    rng = ["#111111","#1f77b4","#2ca02c","#f1c40f","#d62728","#e377c2","#dddddd","#7f7f7f"]
    return alt.Scale(domain=domain, range=rng)

# ==========================================================
# Google Sheets Reader (public)
# ==========================================================
@st.cache_data(ttl=60, show_spinner=False)
def read_gsheet_csv(sheet_id: str, sheet_name: str) -> pd.DataFrame:
    base = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
    r = requests.get(base, params={"tqx": "out:csv", "sheet": sheet_name}, timeout=20)
    r.raise_for_status()
    r.encoding = "utf-8"
    return pd.read_csv(StringIO(r.text))

# ==========================================================
# Load main excel sheets (Lot 관리)
# ==========================================================
@st.cache_data(show_spinner=False)
def load_dataframes(xlsx_path: str) -> dict[str, pd.DataFrame]:
    def read(name: str) -> pd.DataFrame:
        return pd.read_excel(xlsx_path, sheet_name=name)

    out = {
        "binder": read(SHEET_BINDER),
        "single": read(SHEET_SINGLE),
        "spec_binder": read(SHEET_SPEC_BINDER),
        "spec_single": read(SHEET_SPEC_SINGLE),
        "base_lab": read(SHEET_BASE_LAB),
    }
    try:
        out["binder_return"] = pd.read_excel(xlsx_path, sheet_name=SHEET_BINDER_RETURN)
    except Exception:
        out["binder_return"] = pd.DataFrame(columns=["일자", "바인더타입", "바인더명", "바인더 Lot", "반환량(kg)", "비고"])
    return out

# ==========================================================
# Binder IO file upload
# ==========================================================
def _guess_hema_sil_sheets(sheet_names: list[str]):
    hema = None
    sil = None
    for s in sheet_names:
        u = str(s).upper()
        if hema is None and ("HEMA" in u or "헤마" in s):
            hema = s
        if sil is None and (("SIL" in u) or ("SILIC" in u) or ("실리" in s) or ("실리콘" in s)):
            sil = s
    return hema, sil

@st.cache_data(show_spinner=False)
def load_binder_io_excel(xlsx_bytes: bytes, filename: str) -> dict[str, pd.DataFrame]:
    tmp = Path(f".binder_io_{re.sub(r'[^A-Za-z0-9_.-]','_', filename)}")
    tmp.write_bytes(xlsx_bytes)

    xls = pd.ExcelFile(tmp, engine="openpyxl")
    hema_sh, sil_sh = _guess_hema_sil_sheets(xls.sheet_names)

    out = {}
    if hema_sh:
        out["HEMA"] = pd.read_excel(xls, sheet_name=hema_sh)
    if sil_sh:
        out["Silicone"] = pd.read_excel(xls, sheet_name=sil_sh)

    if not out:
        out["ALL"] = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

    # 날짜 컬럼 있으면 최신순 정렬
    for k, df in list(out.items()):
        if df is None or df.empty:
            continue
        dc = detect_date_col(df)
        if dc:
            df2 = df.copy()
            df2["_sort_date"] = pd.to_datetime(df2[dc], errors="coerce")
            df2 = df2.sort_values(by="_sort_date", ascending=False).drop(columns=["_sort_date"])
            out[k] = df2

    return out

# ==========================================================
# Title
# ==========================================================
st.title("액상 잉크 Lot 추적 관리 대시보드")
st.caption("✅ 대시보드 | ✅ 요약 | ✅ 액상잉크 재고관리(재고/입고/사용량) | ✅ 바인더 입출고(파일 업로드/구글시트) | ✅ 빠른검색")

# ==========================================================
# Sidebar - files
# ==========================================================
with st.sidebar:
    st.header("데이터 파일 (Lot 관리)")
    xlsx_path = st.text_input("엑셀 파일 경로", value=DEFAULT_XLSX)
    uploaded = st.file_uploader("또는 엑셀 업로드(.xlsx)", type=["xlsx"], key="lot_upload")

    st.divider()
    st.header("재고 파일")
    stock_xlsx_path = st.text_input("재고 엑셀 파일 경로", value=DEFAULT_STOCK_XLSX, key="stock_path")
    uploaded_stock = st.file_uploader("또는 재고 엑셀 업로드(.xlsx)", type=["xlsx"], key="stock_upload")

# 업로드 파일을 임시 파일로 사용(전체 교체용)
if uploaded is not None:
    sig = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.get("_uploaded_sig") != sig:
        tmp_path = Path(".streamlit_tmp.xlsx")
        tmp_path.write_bytes(uploaded.getvalue())
        st.session_state["_uploaded_sig"] = sig
        st.session_state["_tmp_xlsx_path"] = str(tmp_path)
    xlsx_path = st.session_state.get("_tmp_xlsx_path", xlsx_path)
    st.sidebar.info("업로드 파일(Lot 관리)로 실행 중입니다. (서버 재시작 시 누적 저장은 보장되지 않습니다.)")

if uploaded_stock is not None:
    sig = f"{uploaded_stock.name}:{uploaded_stock.size}"
    if st.session_state.get("_uploaded_sig_stock") != sig:
        tmp_path = Path(".streamlit_tmp_stock.xlsx")
        tmp_path.write_bytes(uploaded_stock.getvalue())
        st.session_state["_uploaded_sig_stock"] = sig
        st.session_state["_tmp_stock_path"] = str(tmp_path)
    stock_xlsx_path = st.session_state.get("_tmp_stock_path", stock_xlsx_path)
    st.sidebar.info("업로드 파일(재고)로 실행 중입니다.")

# ==========================================================
# Load Lot excel (중요: 파일 없으면 멈추지 않고 '빈 데이터'로 화면 표시)
# ==========================================================
if not Path(xlsx_path).exists():
    st.error(f"엑셀 파일을 찾을 수 없습니다: {xlsx_path}")
    st.info("좌측 사이드바에서 Lot관리 엑셀을 업로드하거나, 경로를 올바르게 수정해주세요. (현재는 '빈 데이터'로 화면만 표시합니다.)")

    binder_df = pd.DataFrame(columns=["제조/입고일", "Lot(자동)", "판정"])
    single_df = pd.DataFrame(columns=["입고일","점도측정값(cP)","점도판정","단일색잉크 Lot","사용된 바인더 Lot","색상군","제품코드"])
    spec_binder = pd.DataFrame()
    spec_single = pd.DataFrame(columns=["제품코드", "색상군"])
    base_lab = pd.DataFrame()
    binder_return_df = pd.DataFrame(columns=["일자", "바인더타입", "바인더명", "바인더 Lot", "반환량(kg)", "비고"])
else:
    ensure_sheet_exists(
        xlsx_path,
        SHEET_BINDER_RETURN,
        headers=["일자", "바인더타입", "바인더명", "바인더 Lot", "반환량(kg)", "비고"]
    )
    data = load_dataframes(xlsx_path)
    binder_df = data["binder"].copy()
    single_df = data["single"].copy()
    spec_binder = data["spec_binder"].copy()
    spec_single = data["spec_single"].copy()
    base_lab = data["base_lab"].copy()
    binder_return_df = data["binder_return"].copy()

# normalize dates
c_b_date = find_col(binder_df, "제조/입고일")
c_s_date = find_col(single_df, "입고일")
if c_b_date and c_b_date in binder_df.columns:
    binder_df[c_b_date] = binder_df[c_b_date].apply(normalize_date)
if c_s_date and c_s_date in single_df.columns:
    single_df[c_s_date] = single_df[c_s_date].apply(normalize_date)

# common cols
c_s_visc = find_col(single_df, "점도측정값(cP)")
c_s_judge = find_col(single_df, "점도판정")
c_s_lot = find_col(single_df, "단일색잉크 Lot")
c_s_blot = find_col(single_df, "사용된 바인더 Lot")
c_s_cg = find_col(single_df, "색상군")
c_s_pc = find_col(single_df, "제품코드")

# ==========================================================
# Tabs
# ==========================================================
tab_dash, tab_summary, tab_stock, tab_binder, tab_search = st.tabs(
    ["📊 대시보드", "📌 요약", "📦 액상잉크 재고관리", "📦 바인더 입출고", "🔎 빠른검색"]
)

# ==========================================================
# Summary tab
# ==========================================================
def render_summary():
    st.markdown('<div class="section-title">📌 요약</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">상사가 “한 번에 이해”할 수 있게 KPI + 그래프 4개 + 상세(펼침) 구조</div>', unsafe_allow_html=True)

    # 재고(최근 30일)
    stock_ok = bool(stock_xlsx_path and Path(stock_xlsx_path).exists())
    product_to_color = build_product_to_color_map(spec_single, single_df)

    inv_color = pd.DataFrame()
    use_color = pd.DataFrame()
    cov_alert = pd.DataFrame()
    stock_kpis = {}

    if stock_ok:
        hist = load_stock_history(stock_xlsx_path, product_to_color)
        if not hist.empty:
            max_d = hist["date"].max().date()
            start = max(hist["date"].min().date(), max_d - dt.timedelta(days=29))
            end = max_d
            day_span = max(1, (end - start).days + 1)

            hist_f = hist[(hist["date"].dt.date >= start) & (hist["date"].dt.date <= end)].copy()
            latest_df = hist[hist["date"].dt.date == end].copy()

            stock_kpis["재고 최신일"] = end.isoformat()
            stock_kpis["현재 총 재고(kg)"] = float(latest_df["curr_stock_kg"].sum())
            stock_kpis["최근 30일 사용량(kg)"] = float(hist_f["used_kg"].sum())
            stock_kpis["최근 30일 입고(건)"] = int(hist_f["inbound_event"].sum())
            stock_kpis["평균 사용량(kg/일)"] = float(stock_kpis["최근 30일 사용량(kg)"] / day_span)

            inv_color = (
                latest_df.groupby("color_group", as_index=False)["curr_stock_kg"]
                .sum().rename(columns={"curr_stock_kg": "kg"})
                .sort_values("kg", ascending=False)
            )
            use_color = (
                hist_f.groupby("color_group", as_index=False)["used_kg"]
                .sum().rename(columns={"used_kg": "kg"})
                .sort_values("kg", ascending=False)
            )

            use_by_product = hist_f.groupby("product_code", as_index=False)["used_kg"].sum()
            use_by_product["avg_daily_use"] = use_by_product["used_kg"] / day_span
            stock_by_product = latest_df.groupby("product_code", as_index=False)["curr_stock_kg"].sum().rename(
                columns={"curr_stock_kg": "stock_kg"}
            )
            cov = stock_by_product.merge(use_by_product[["product_code", "avg_daily_use"]], on="product_code", how="left")
            cov["avg_daily_use"] = cov["avg_daily_use"].fillna(0.0)
            cov["cover_days"] = cov.apply(
                lambda r: (r["stock_kg"] / r["avg_daily_use"]) if r["avg_daily_use"] > 0 else None, axis=1
            )
            cov_alert = cov[cov["cover_days"].notna()].sort_values("cover_days").head(10)
        else:
            stock_ok = False

    # 점도(최근 30일)
    visc_ok = bool(c_s_date and c_s_visc and c_s_pc and (c_s_date in single_df.columns) and (c_s_visc in single_df.columns) and (c_s_pc in single_df.columns))
    visc_kpis = {}
    daily_visc = pd.DataFrame()
    top_ng = pd.DataFrame()

    if visc_ok:
        df = single_df.copy()
        df[c_s_date] = pd.to_datetime(df[c_s_date], errors="coerce")
        df["_점도"] = pd.to_numeric(df[c_s_visc].astype(str).str.replace(",", "", regex=False), errors="coerce")
        df = df.dropna(subset=[c_s_date, "_점도", c_s_pc])

        if len(df):
            max_d = df[c_s_date].max().date()
            start = max(df[c_s_date].min().date(), max_d - dt.timedelta(days=29))
            df30 = df[(df[c_s_date].dt.date >= start) & (df[c_s_date].dt.date <= max_d)].copy()

            total = len(df30)
            ng = int((df30[c_s_judge] == "부적합").sum()) if c_s_judge and (c_s_judge in df30.columns) else 0
            ng_rate = (ng / total * 100) if total else 0.0

            visc_kpis = {
                "점도 최신일": max_d.isoformat(),
                "최근 30일 측정(건)": total,
                "부적합(건)": ng,
                "부적합률(%)": ng_rate,
            }

            daily_visc = (
                df30.groupby(df30[c_s_date].dt.date)
                .agg(mean_visc=("_점도", "mean"), cnt=("_점도", "size"))
                .reset_index()
                .rename(columns={df30.groupby(df30[c_s_date].dt.date).agg(mean_visc=("_점도","mean")).reset_index().columns[0]: "date"})
            )
            daily_visc["date"] = pd.to_datetime(daily_visc["date"])

            if c_s_judge and (c_s_judge in df30.columns):
                top_ng = (
                    df30[df30[c_s_judge] == "부적합"]
                    .groupby(c_s_pc).size().reset_index(name="ng_cnt")
                    .sort_values("ng_cnt", ascending=False).head(8)
                )
        else:
            visc_ok = False

    # KPIs
    a, b = st.columns(2)
    with a:
        st.markdown("#### 🧾 재고(최근 30일)")
        if not stock_ok:
            st.info("재고 파일이 없거나 읽지 못했습니다. (좌측 사이드바에서 재고 파일 경로/업로드 설정)")
        else:
            k1, k2, k3, k4, k5 = st.columns([1.2, 1.7, 1.7, 1.4, 1.8])
            k1.metric("최신일", stock_kpis["재고 최신일"])
            k2.metric("총 재고(kg)", f'{stock_kpis["현재 총 재고(kg)"]:,.1f}')
            k3.metric("30일 사용량(kg)", f'{stock_kpis["최근 30일 사용량(kg)"]:,.1f}')
            k4.metric("입고(건)", f'{stock_kpis["최근 30일 입고(건)"]:,}')
            k5.metric("일평균(kg/일)", f'{stock_kpis["평균 사용량(kg/일)"]:,.1f}')

    with b:
        st.markdown("#### 🧪 점도(최근 30일)")
        if not visc_ok:
            st.info("단일색 시트에 입고일/점도측정값/제품코드 컬럼이 필요합니다.")
        else:
            k1, k2, k3 = st.columns(3)
            k1.metric("최신일", visc_kpis["점도 최신일"])
            k2.metric("측정(건)", f'{visc_kpis["최근 30일 측정(건)"]:,}')
            k3.metric("부적합률(%)", f'{visc_kpis["부적합률(%)"]:.1f}')

    st.divider()
    st.markdown("#### 📊 한눈에 보는 그래프")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**재고(최신일) — 색상계열(BLACK/RED …)**")
        if stock_ok and not inv_color.empty:
            ch = alt.Chart(inv_color).mark_bar().encode(
                y=alt.Y("color_group:N", sort="-x", title=""),
                x=alt.X("kg:Q", title="재고(kg)"),
                color=alt.Color("color_group:N", scale=_color_scale_color_group(), legend=None),
                tooltip=[alt.Tooltip("color_group:N", title="색상계열"), alt.Tooltip("kg:Q", title="재고(kg)", format=",.1f")]
            ).properties(height=260)
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("재고 데이터가 없습니다.")

    with c2:
        st.markdown("**최근 30일 평균 점도(일별)**")
        if visc_ok and not daily_visc.empty:
            ch = alt.Chart(daily_visc).mark_line(point=True).encode(
                x=alt.X("date:T", title="날짜"),
                y=alt.Y("mean_visc:Q", title="평균 점도(cP)"),
                tooltip=[alt.Tooltip("date:T", title="날짜"),
                         alt.Tooltip("mean_visc:Q", title="평균점도", format=",.0f"),
                         alt.Tooltip("cnt:Q", title="측정(건)", format=",.0f")]
            ).properties(height=260)
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("점도 데이터가 없습니다.")

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**최근 30일 사용량 — 색상계열(BLACK/RED …)**")
        if stock_ok and not use_color.empty:
            ch = alt.Chart(use_color).mark_bar().encode(
                y=alt.Y("color_group:N", sort="-x", title=""),
                x=alt.X("kg:Q", title="사용량(kg)"),
                color=alt.Color("color_group:N", scale=_color_scale_color_group(), legend=None),
                tooltip=[alt.Tooltip("color_group:N", title="색상계열"), alt.Tooltip("kg:Q", title="사용량(kg)", format=",.1f")]
            ).properties(height=260)
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("사용량 데이터가 없습니다.")

    with c4:
        st.markdown("**부적합 Top 제품코드(최근 30일)**")
        if visc_ok and not top_ng.empty:
            ch = alt.Chart(top_ng).mark_bar().encode(
                y=alt.Y(f"{c_s_pc}:N", sort="-x", title=""),
                x=alt.X("ng_cnt:Q", title="부적합(건)"),
                tooltip=[alt.Tooltip(f"{c_s_pc}:N", title="제품코드"), alt.Tooltip("ng_cnt:Q", title="부적합(건)", format=",.0f")]
            ).properties(height=260)
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("부적합 데이터가 없습니다.")

    with st.expander("🔎 (상세) 커버리지 경보 Top10 보기"):
        if stock_ok and not cov_alert.empty:
            show = cov_alert.copy()
            show["stock_kg"] = show["stock_kg"].round(1)
            show["avg_daily_use"] = show["avg_daily_use"].round(2)
            show["cover_days"] = show["cover_days"].round(1)
            st.dataframe(show, use_container_width=True, height=320)
        else:
            st.info("커버리지 계산 데이터가 없습니다.")

# ==========================================================
# Stock tab
# ==========================================================
def render_stock_tab():
    st.markdown('<div class="section-title">📦 액상잉크 재고관리</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">재고(현재) · 입고(추정) · 사용량(일별)을 BLACK/RED 등 색상계열로 요약합니다.</div>', unsafe_allow_html=True)

    if not stock_xlsx_path or not Path(stock_xlsx_path).exists():
        st.error("재고 파일 경로가 올바르지 않습니다. (좌측 사이드바에서 재고 파일 경로/업로드 설정)")
        return

    product_to_color = build_product_to_color_map(spec_single, single_df)
    hist = load_stock_history(stock_xlsx_path, product_to_color)
    if hist.empty:
        st.error("재고 엑셀을 읽지 못했습니다. (시트명: 1.15 형식 / 컬럼: 품목명, 금일 재고(kg), 하루 사용량(kg) 확인)")
        return

    min_d = hist["date"].min().date()
    max_d = hist["date"].max().date()

    left, mid, right = st.columns([2.2, 2.8, 5.0])
    with left:
        quick = st.selectbox("기간(빠른 선택)", ["최근 7일", "최근 30일", "최근 90일", "전체", "직접 선택"], index=1)
    with mid:
        if quick == "직접 선택":
            start = st.date_input("시작일", value=max(min_d, max_d - dt.timedelta(days=30)), min_value=min_d, max_value=max_d)
            end = st.date_input("종료일", value=max_d, min_value=min_d, max_value=max_d)
        else:
            if quick == "최근 7일":
                start = max(min_d, max_d - dt.timedelta(days=6))
            elif quick == "최근 30일":
                start = max(min_d, max_d - dt.timedelta(days=29))
            elif quick == "최근 90일":
                start = max(min_d, max_d - dt.timedelta(days=89))
            else:
                start = min_d
            end = max_d
            st.write(f"**{start} ~ {end}**")
    with right:
        divisions = sorted([x for x in hist["division"].dropna().unique().tolist() if str(x).strip() and str(x).lower() not in ("nan", "none")])
        sel_div = st.multiselect("구분(PL/NPL/NSL 등)", divisions, default=divisions)

    if start > end:
        start, end = end, start

    filt = (hist["date"].dt.date >= start) & (hist["date"].dt.date <= end)
    if sel_div:
        filt = filt & (hist["division"].isin(sel_div))
    hist_f = hist[filt].copy()

    latest_date = hist["date"].max()
    latest_df = hist[hist["date"] == latest_date].copy()
    if sel_div:
        latest_df = latest_df[latest_df["division"].isin(sel_div)].copy()

    total_stock = float(latest_df["curr_stock_kg"].sum())
    total_used = float(hist_f["used_kg"].sum())
    inbound_events = int(hist_f["inbound_event"].sum())
    day_span = max(1, (end - start).days + 1)
    avg_daily_use = total_used / day_span if day_span else 0.0

    k1, k2, k3, k4, k5 = st.columns([1.4, 1.6, 1.6, 1.6, 1.8])
    k1.metric("재고 최신일", latest_date.date().isoformat())
    k2.metric("현재 총 재고(kg)", f"{total_stock:,.1f}")
    k3.metric("기간 총 사용량(kg)", f"{total_used:,.1f}")
    k4.metric("기간 입고(건)", f"{inbound_events:,}")
    k5.metric("평균 일 사용량(kg/일)", f"{avg_daily_use:,.1f}")

    st.markdown('<div class="kpi-note">※ 입고(kg/건)는 "하루 사용량"이 음수로 기입된 경우(재고 증가)를 입고로 추정합니다.</div>', unsafe_allow_html=True)
    st.divider()

    inv = latest_df.groupby("color_group", as_index=False)["curr_stock_kg"].sum().rename(columns={"curr_stock_kg":"kg"}).sort_values("kg", ascending=False)
    use = hist_f.groupby("color_group", as_index=False)["used_kg"].sum().rename(columns={"used_kg":"kg"}).sort_values("kg", ascending=False)

    def bar_chart(df: pd.DataFrame, value_title: str):
        if df.empty:
            return None
        return alt.Chart(df).mark_bar().encode(
            y=alt.Y("color_group:N", sort="-x", title="색상계열"),
            x=alt.X("kg:Q", title=value_title),
            color=alt.Color("color_group:N", scale=_color_scale_color_group(), legend=None),
            tooltip=[alt.Tooltip("color_group:N", title="색상계열"), alt.Tooltip("kg:Q", title=value_title, format=",.1f")],
        ).properties(height=240)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 1) 현재 재고(최신일) — 색상계열")
        ch = bar_chart(inv, "재고(kg)")
        st.altair_chart(ch, use_container_width=True) if ch else st.info("표시할 재고 데이터가 없습니다.")
    with c2:
        st.markdown("### 2) 기간 사용량 — 색상계열")
        ch = bar_chart(use, "사용량(kg)")
        st.altair_chart(ch, use_container_width=True) if ch else st.info("표시할 사용량 데이터가 없습니다.")

    st.divider()
    st.markdown("### 3) 일별 사용량 추이(kg)")

    present = [k for k in COLOR_KEYS if k in hist_f["color_group"].unique().tolist()]
    default_keys = [k for k in present if k != "OTHER"][:5] or present
    sel_keys = st.multiselect("표시할 색상계열", COLOR_KEYS, default=default_keys)

    daily = hist_f[hist_f["color_group"].isin(sel_keys)].groupby(["date","color_group"], as_index=False)["used_kg"].sum()
    total = hist_f.groupby("date", as_index=False)["used_kg"].sum().rename(columns={"used_kg":"TOTAL"})

    line = alt.Chart(daily).mark_line(point=True).encode(
        x=alt.X("date:T", title="날짜"),
        y=alt.Y("used_kg:Q", title="사용량(kg)"),
        color=alt.Color("color_group:N", scale=_color_scale_color_group(), legend=alt.Legend(title="색상계열")),
        tooltip=[alt.Tooltip("date:T", title="날짜"), alt.Tooltip("color_group:N", title="색상계열"), alt.Tooltip("used_kg:Q", title="사용량(kg)", format=",.1f")]
    )
    total_line = alt.Chart(total).mark_line(point=True, strokeDash=[6,3]).encode(
        x="date:T", y=alt.Y("TOTAL:Q", title="사용량(kg)"),
        tooltip=[alt.Tooltip("date:T", title="날짜"), alt.Tooltip("TOTAL:Q", title="TOTAL(kg)", format=",.1f")]
    )
    st.altair_chart((line + total_line).interactive(), use_container_width=True)

    st.divider()
    st.markdown("### 4) 재고 커버리지(일) 경보 (품목)")
    target_days = st.slider("목표 커버리지(일)", 3, 30, 14, 1)
    alert_days = st.slider("경보 기준(일)", 1, 21, 7, 1)

    use_by_product = hist_f.groupby("product_code", as_index=False)["used_kg"].sum()
    use_by_product["avg_daily_use"] = use_by_product["used_kg"] / day_span
    stock_by_product = latest_df.groupby("product_code", as_index=False)["curr_stock_kg"].sum().rename(columns={"curr_stock_kg":"stock_kg"})
    cov = stock_by_product.merge(use_by_product[["product_code","avg_daily_use"]], on="product_code", how="left")
    cov["avg_daily_use"] = cov["avg_daily_use"].fillna(0.0)
    cov["cover_days"] = cov.apply(lambda r: (r["stock_kg"]/r["avg_daily_use"]) if r["avg_daily_use"]>0 else None, axis=1)
    cov["need_order_kg"] = cov.apply(lambda r: max(0.0, target_days*r["avg_daily_use"]-r["stock_kg"]) if r["avg_daily_use"]>0 else None, axis=1)

    alert_df = cov[(cov["cover_days"].notna()) & (cov["cover_days"] <= float(alert_days))].sort_values("cover_days").head(20)
    if alert_df.empty:
        st.success("✅ 경보 기준 이하(커버리지 부족) 품목이 없습니다.")
    else:
        tmp = alert_df.copy()
        tmp["stock_kg"] = tmp["stock_kg"].round(1)
        tmp["avg_daily_use"] = tmp["avg_daily_use"].round(2)
        tmp["cover_days"] = tmp["cover_days"].round(1)
        tmp["need_order_kg"] = tmp["need_order_kg"].round(1)
        st.warning(f"⚠️ 커버리지 {alert_days}일 이하 품목(상위 20개)")
        st.dataframe(tmp, use_container_width=True, height=360)

# ==========================================================
# Dashboard tab
# ==========================================================
def render_dashboard():
    b_total = len(binder_df)
    s_total = len(single_df)

    c_b_judge = find_col(binder_df, "판정")
    b_ng = int((binder_df[c_b_judge] == "부적합").sum()) if c_b_judge and (c_b_judge in binder_df.columns) else 0
    s_ng = int((single_df[c_s_judge] == "부적합").sum()) if c_s_judge and (c_s_judge in single_df.columns) else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("바인더 기록", f"{b_total:,}")
    c2.metric("바인더 부적합", f"{b_ng:,}")
    c3.metric("단일색 기록", f"{s_total:,}")
    c4.metric("단일색(점도) 부적합", f"{s_ng:,}")

    st.divider()
    st.subheader("단일색 데이터 목록(필터)")

    if not (c_s_date and c_s_visc and c_s_pc and (c_s_date in single_df.columns) and (c_s_visc in single_df.columns) and (c_s_pc in single_df.columns)):
        st.warning("단일색 시트에서 입고일/점도측정값/제품코드 컬럼을 찾지 못했습니다.")
        return

    df = single_df.copy()
    df[c_s_date] = pd.to_datetime(df[c_s_date], errors="coerce")
    dmin, dmax = safe_date_bounds(df[c_s_date])

    f1, f2, f3 = st.columns([1.2, 1.2, 3.0])
    with f1:
        start = st.date_input("시작일", value=max(dmin, dmax - dt.timedelta(days=90)))
    with f2:
        end = st.date_input("종료일", value=dmax)
    with f3:
        pcs = sorted(df[c_s_pc].dropna().astype(str).unique().tolist())
        sel_pc = st.multiselect("제품코드", pcs, default=[])

    if start > end:
        start, end = end, start

    df = df[(df[c_s_date].dt.date >= start) & (df[c_s_date].dt.date <= end)]
    if sel_pc:
        df = df[df[c_s_pc].astype(str).isin(sel_pc)]

    view = pd.DataFrame({
        "입고일": df[c_s_date].dt.date,
        "색상군": df[c_s_cg].apply(normalize_color_group) if c_s_cg and (c_s_cg in df.columns) else None,
        "제품코드": df[c_s_pc],
        "단일색Lot": df[c_s_lot] if c_s_lot and (c_s_lot in df.columns) else None,
        "사용바인더Lot": df[c_s_blot] if c_s_blot and (c_s_blot in df.columns) else None,
        "점도(cP)": pd.to_numeric(df[c_s_visc].astype(str).str.replace(",", "", regex=False), errors="coerce"),
        "점도판정": df[c_s_judge] if c_s_judge and (c_s_judge in df.columns) else None,
    }).dropna(subset=["입고일"]).sort_values("입고일", ascending=False)

    st.dataframe(view, use_container_width=True, height=320)

# ==========================================================
# Binder IO tab
# ==========================================================
def render_binder_io():
    st.subheader("바인더 입출고 내역 (파일 업로드 / 구글시트)")
    st.caption("✅ 바인더 입출고 내역 파일(.xlsx)을 업로드하면, 업로드 즉시 아래에 입출고 표가 바로 표시됩니다.")

    binder_io_file = st.file_uploader("바인더 입출고 내역 파일 업로드(.xlsx)", type=["xlsx"], key="binder_io_upload")
    if binder_io_file is not None:
        try:
            io_data = load_binder_io_excel(binder_io_file.getvalue(), binder_io_file.name)
            st.success("업로드 파일을 불러왔습니다.")
            if "HEMA" in io_data and "Silicone" in io_data:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("### HEMA (파일)")
                    st.dataframe(io_data["HEMA"], use_container_width=True, height=420)
                with c2:
                    st.markdown("### Silicone (파일)")
                    st.dataframe(io_data["Silicone"], use_container_width=True, height=420)
            else:
                key = list(io_data.keys())[0]
                st.markdown(f"### {key} (파일)")
                st.dataframe(io_data[key], use_container_width=True, height=520)
        except Exception as e:
            st.error("업로드 파일을 읽지 못했습니다. (파일 형식/시트 구조 확인)")
            st.exception(e)

    st.divider()
    st.subheader("바인더 입출고 (Google Sheets 자동 반영)")
    st.caption("구글 시트를 수정하면 이 화면은 새로고침 시 자동 반영됩니다. (캐시 60초)")

    try:
        df_hema = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_HEMA)
        df_sil = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_SIL)
    except Exception as e:
        st.error("구글시트에서 데이터를 못 불러왔습니다. (시트 공유/웹게시/시트명/ID 확인)")
        st.exception(e)
        return

    for _df in [df_hema, df_sil]:
        dc = detect_date_col(_df)
        if dc:
            _df["_sort_date"] = pd.to_datetime(_df[dc], errors="coerce")
            _df.sort_values(by="_sort_date", ascending=False, inplace=True)
            _df.drop(columns=["_sort_date"], inplace=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### HEMA (구글시트)")
        st.dataframe(df_hema, use_container_width=True, height=420)
    with c2:
        st.markdown("### Silicone (구글시트)")
        st.dataframe(df_sil, use_container_width=True, height=420)

    if st.button("지금 최신값으로 다시 불러오기"):
        st.cache_data.clear()
        st.rerun()

# ==========================================================
# Search tab
# ==========================================================
def render_search():
    st.subheader("빠른검색")
    mode = st.selectbox("검색 종류", ["바인더 Lot", "단일색 Lot", "제품코드"])
    q = st.text_input("검색어", placeholder="예: PCB20250112-01 / PLB25041501 / PL-835-1 ...")

    # prep
    s_df = single_df.copy()
    if c_s_date and (c_s_date in s_df.columns):
        s_df[c_s_date] = pd.to_datetime(s_df[c_s_date], errors="coerce")
    b_df = binder_df.copy()
    if c_b_date and (c_b_date in b_df.columns):
        b_df[c_b_date] = pd.to_datetime(b_df[c_b_date], errors="coerce")

    def text_filter(df: pd.DataFrame, cols: list[str], text: str) -> pd.DataFrame:
        if not text:
            return df.iloc[0:0]
        t = str(text).strip()
        if not t:
            return df.iloc[0:0]
        mask = False
        for c in cols:
            if c and c in df.columns:
                mask = mask | df[c].astype(str).str.contains(t, case=False, na=False)
        return df[mask]

    if mode == "바인더 Lot":
        c_bl = find_col(b_df, "Lot(자동)")
        hit_b = text_filter(b_df, [c_bl], q)
        st.markdown("#### 바인더_제조_입고")
        st.dataframe(add_excel_row_number(hit_b), use_container_width=True)

        if q and c_s_blot and (c_s_blot in s_df.columns):
            hit_s = s_df[s_df[c_s_blot].astype(str).str.contains(str(q).strip(), case=False, na=False)]
            st.markdown("#### 연결된 단일색_수입검사 (사용된 바인더 Lot)")
            st.dataframe(add_excel_row_number(hit_s), use_container_width=True)

    elif mode == "단일색 Lot":
        hit = text_filter(s_df, [c_s_lot], q)
        st.dataframe(add_excel_row_number(hit), use_container_width=True)

    else:  # 제품코드
        hit = text_filter(s_df, [c_s_pc], q)
        st.dataframe(add_excel_row_number(hit), use_container_width=True)

# ==========================================================
# Render tabs
# ==========================================================
with tab_dash:
    render_dashboard()

with tab_summary:
    render_summary()

with tab_stock:
    render_stock_tab()

with tab_binder:
    render_binder_io()

with tab_search:
    render_search()
