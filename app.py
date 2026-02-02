# app.py
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
# Page / Style
# ==========================================================
st.set_page_config(page_title="액상 잉크 Lot 추적 관리", page_icon="🧪", layout="wide")

st.markdown(
    """
    <style>
      /* 전체 레이아웃 */
      .block-container { 
        padding-top: 1.5rem; 
        padding-bottom: 2rem; 
        max-width: 1400px;
      }
      
      /* 타이틀 */
      h1 {
        color: #1f2937;
        font-weight: 800;
        font-size: 2.2rem !important;
        margin-bottom: 0.5rem !important;
      }
      
      /* KPI 카드 스타일 */
      div[data-testid="metric-container"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 1rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
      }
      
      div[data-testid="metric-container"] label {
        color: white !important;
        font-weight: 700 !important;
        font-size: 0.85rem !important;
      }
      
      div[data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: white !important;
        font-size: 1.8rem !important;
        font-weight: 900 !important;
      }
      
      /* 섹션 타이틀 */
      .section-title { 
        font-size: 1.4rem; 
        font-weight: 900; 
        margin: 1.5rem 0 0.5rem 0;
        color: #1f2937;
        border-left: 5px solid #667eea;
        padding-left: 12px;
      }
      
      .section-sub { 
        color: #6b7280; 
        font-size: 0.95rem; 
        margin-bottom: 1rem;
        padding-left: 17px;
      }
      
      /* KPI 노트 */
      .kpi-note { 
        color: #6b7280; 
        font-size: 0.88rem; 
        margin-top: 0.5rem;
        padding: 0.8rem;
        background: #f3f4f6;
        border-radius: 8px;
        border-left: 3px solid #fbbf24;
      }
      
      /* Expander 스타일 */
      div[data-testid="stExpander"] {
        border: 2px solid #e5e7eb;
        border-radius: 10px;
        background: #ffffff;
      }
      
      div[data-testid="stExpander"] > details > summary { 
        font-weight: 800;
        color: #374151;
        font-size: 1.05rem;
      }
      
      /* 탭 스타일 */
      .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f9fafb;
        padding: 8px;
        border-radius: 10px;
      }
      
      .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 700;
        font-size: 0.95rem;
      }
      
      .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white !important;
      }
      
      /* 테이블 스타일 */
      .dataframe {
        border-radius: 10px !important;
        overflow: hidden;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
      }
      
      /* 버튼 스타일 */
      .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.5rem;
        font-weight: 700;
        transition: all 0.3s;
      }
      
      .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
      }
      
      /* 다운로드 버튼 */
      .stDownloadButton > button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 700;
      }
      
      /* 경고 박스 */
      .stAlert {
        border-radius: 10px;
        border-left: 5px solid;
      }
      
      /* 차트 컨테이너 */
      .chart-container {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin-bottom: 1rem;
      }
      
      /* 폼 스타일 */
      .stForm {
        background: #f9fafb;
        padding: 1.5rem;
        border-radius: 12px;
        border: 2px solid #e5e7eb;
      }
      
      /* 입력 필드 */
      .stTextInput > div > div > input,
      .stNumberInput > div > div > input,
      .stSelectbox > div > div > div {
        border-radius: 8px;
        border: 2px solid #e5e7eb;
      }
      
      /* 구분선 */
      hr {
        margin: 2rem 0;
        border: none;
        border-top: 2px solid #e5e7eb;
      }
      
      /* 사이드바 */
      .css-1d391kg, [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f9fafb 0%, #ffffff 100%);
      }
      
      /* 상태 배지 */
      .status-badge-success {
        background: #10b981;
        color: white;
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 700;
        display: inline-block;
      }
      
      .status-badge-warning {
        background: #f59e0b;
        color: white;
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 700;
        display: inline-block;
      }
      
      .status-badge-error {
        background: #ef4444;
        color: white;
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 700;
        display: inline-block;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# toast
if "_toast_msg" not in st.session_state:
    st.session_state["_toast_msg"] = None
if st.session_state.get("_toast_msg"):
    try:
        st.toast(st.session_state["_toast_msg"])
    except Exception:
        pass
    st.session_state["_toast_msg"] = None

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

COLOR_KEYS = ["BLACK", "BLUE", "GREEN", "YELLOW", "RED", "PINK", "WHITE", "OTHER"]

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

def normalize_date(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    if isinstance(x, (dt.date, dt.datetime)):
        return x.date() if isinstance(x, dt.datetime) else x
    try:
        return pd.to_datetime(x).date()
    except Exception:
        return None

def safe_date_bounds(series: pd.Series):
    s = pd.to_datetime(series, errors="coerce").dropna()
    if len(s) == 0:
        today = dt.date.today()
        return today, today
    return s.min().date(), s.max().date()

def detect_date_col(df: pd.DataFrame):
    for c in df.columns:
        ck = norm_key(c).lower()
        if any(k in ck for k in ["일자", "날짜", "date", "입고일", "출고일", "반입일", "반출일"]):
            return c
    return None

def file_sig(path: str):
    """캐시 무효화를 위한 시그니처"""
    try:
        p = Path(path)
        if not p.exists():
            return None
        stat = p.stat()
        return (str(p.resolve()), int(stat.st_size), int(stat.st_mtime))
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

# ==========================================================
# Color helpers (요청: BLACK/RED 등 대문자 가시화)
# ==========================================================
def normalize_color_group(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "OTHER"
    s = str(x).strip()
    if not s or s.lower() in ("nan", "none"):
        return "OTHER"

    u = s.upper()
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

def _color_scale_color_group():
    domain = ["BLACK", "BLUE", "GREEN", "YELLOW", "RED", "PINK", "WHITE", "OTHER"]
    rng = ["#111111", "#1f77b4", "#2ca02c", "#f1c40f", "#d62728", "#e377c2", "#dddddd", "#7f7f7f"]
    return alt.Scale(domain=domain, range=rng)

# ==========================================================
# Product -> ColorGroup mapping (Spec + Single)
# ==========================================================
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

# ==========================================================
# Excel append / download
# ==========================================================
def get_sheet_headers(xlsx_path: str, sheet_name: str) -> list[str]:
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    headers = []
    for cell in ws[1]:
        headers.append(None if cell.value is None else str(cell.value).strip())
    while headers and headers[-1] in (None, "", "nan"):
        headers.pop()
    return headers

def append_row_to_xlsx(xlsx_path: str, sheet_name: str, row_dict: dict):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"시트가 없습니다: {sheet_name}")

    ws = wb[sheet_name]
    headers = get_sheet_headers(xlsx_path, sheet_name)
    if not headers:
        raise ValueError(f"헤더(1행)를 찾지 못했습니다: {sheet_name}")

    row = []
    for h in headers:
        if h is None:
            row.append(None)
            continue
        v = row_dict.get(h, None)
        if isinstance(v, dt.datetime):
            v = v.date()
        row.append(v)

    ws.append(row)
    wb.save(xlsx_path)

def download_xlsx_button(xlsx_path: str, label: str = "업데이트된 엑셀 다운로드"):
    try:
        data = Path(xlsx_path).read_bytes()
        st.download_button(
            label,
            data=data,
            file_name=Path(xlsx_path).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error("다운로드 파일 생성 실패")
        st.exception(e)

# ==========================================================
# Lot auto generator (기존 Lot 스타일 추정)
# ==========================================================
def infer_lot_style(existing_lots: list[str], fallback_prefix: str):
    """
    returns (prefix, date_len, sep, seq_len)
    - prefix: leading letters from last lot (or fallback)
    - date_len: 6(YYMMDD) or 8(YYYYMMDD)
    - sep: "-" or ""
    - seq_len: digits
    """
    last = None
    for x in reversed(existing_lots):
        if x and str(x).strip() and str(x).lower() not in ("nan", "none"):
            last = str(x).strip()
            break

    if not last:
        return (fallback_prefix, 6, "-", 2)

    m = re.match(r"^([A-Za-z]+)(.*)$", last)
    if not m:
        return (fallback_prefix, 6, "-", 2)

    prefix = m.group(1) or fallback_prefix
    rest = (m.group(2) or "").strip()

    sep = "-" if "-" in rest else ""
    if sep:
        parts = rest.split("-", 1)
        date_part = parts[0]
        seq_part = parts[1] if len(parts) > 1 else ""
        date_len = 8 if (date_part.startswith("20") and len(date_part) >= 8) else 6
        seq_len = max(2, len(seq_part)) if seq_part else 2
        return (prefix, date_len, sep, seq_len)

    # no sep
    date_len = 8 if (rest.startswith("20") and len(rest) >= 8) else 6
    seq_len = max(2, len(rest) - date_len)
    return (prefix, date_len, "", seq_len)

def next_lot(existing_lots: list[str], date_value: dt.date, fallback_prefix: str):
    prefix, date_len, sep, seq_len = infer_lot_style(existing_lots, fallback_prefix)
    date_str = date_value.strftime("%Y%m%d") if date_len == 8 else date_value.strftime("%y%m%d")

    pat = re.compile(rf"^{re.escape(prefix)}{re.escape(date_str)}{re.escape(sep)}(\d+)$")
    max_seq = 0
    for x in existing_lots:
        if not x:
            continue
        s = str(x).strip()
        mm = pat.match(s)
        if not mm:
            continue
        try:
            max_seq = max(max_seq, int(mm.group(1)))
        except Exception:
            pass

    seq = max_seq + 1
    seq_str = str(seq).zfill(seq_len)
    return f"{prefix}{date_str}{sep}{seq_str}"

# ==========================================================
# Stock history loader (재고 파일: 시트명 1.15 등)
# ==========================================================
def _parse_stock_sheet_date(sheet_name: str, today: dt.date):
    s = str(sheet_name).strip()
    m = re.match(r"^(\d{1,2})\.(\d{1,2})$", s)  # 예: 1.15
    if not m:
        return None
    month = int(m.group(1))
    day = int(m.group(2))
    year = today.year
    # 연말/연초 경계 보정
    if month > (today.month + 1):
        year -= 1
    try:
        return dt.date(year, month, day)
    except ValueError:
        return None

@st.cache_data(show_spinner=False)
def load_stock_history(stock_xlsx_path: str, stock_sig, product_to_color: dict[str, str]) -> pd.DataFrame:
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

        out = pd.DataFrame()
        out["division"] = df[c_div].astype(str).str.strip() if c_div else ""
        out["product_code"] = df[c_item].apply(normalize_product_code)

        out["curr_stock_kg"] = pd.to_numeric(
            df[c_curr].astype(str).str.replace(",", "", regex=False), errors="coerce"
        ).fillna(0)

        used_raw = pd.to_numeric(df[c_used].astype(str).str.replace(",", "", regex=False), errors="coerce")
        # 사용량: 양수 / 입고(추정): 음수(재고 증가로 입력되는 경우) 처리
        out["used_kg"] = used_raw.clip(lower=0).fillna(0)
        out["inbound_kg"] = (-used_raw).clip(lower=0).fillna(0)
        out["inbound_event"] = (out["inbound_kg"] > 0).astype(int)

        out = out.dropna(subset=["product_code"])
        out["color_group"] = out["product_code"].map(product_to_color).fillna("OTHER").apply(normalize_color_group)
        out["date"] = pd.to_datetime(d)
        frames.append(out[["date", "division", "product_code", "color_group", "curr_stock_kg", "used_kg", "inbound_kg", "inbound_event"]])

    if not frames:
        return pd.DataFrame()

    hist = pd.concat(frames, ignore_index=True)
    hist = hist.sort_values(["date", "division", "product_code"]).reset_index(drop=True)
    return hist

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
# Load Lot excel sheets (Lot 관리)
# ==========================================================
@st.cache_data(show_spinner=False)
def load_dataframes(xlsx_path: str, lot_sig) -> dict[str, pd.DataFrame]:
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
# Binder IO file upload (엑셀 업로드 즉시 표시)
# ==========================================================
def _guess_hema_sil_sheets(sheet_names: list[str]):
    hema = None
    sil = None
    for s in sheet_names:
        u = str(s).upper()
        if hema is None and ("HEMA" in u or "헤마" in str(s)):
            hema = s
        if sil is None and (("SIL" in u) or ("SILIC" in u) or ("실리" in str(s)) or ("실리콘" in str(s))):
            sil = s
    return hema, sil

@st.cache_data(show_spinner=False)
def load_binder_io_excel(xlsx_bytes: bytes, filename: str) -> dict[str, pd.DataFrame]:
    tmp = Path(f".binder_io_{re.sub(r'[^A-Za-z0-9_.-]', '_', filename)}")
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

    # 날짜 컬럼 감지 후 최근순 정렬
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
st.title("🧪 액상 잉크 Lot 추적 관리 대시보드")
st.markdown(
    """
    <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                padding: 1rem 1.5rem; 
                border-radius: 12px; 
                margin-bottom: 1.5rem;
                color: white;'>
        <h3 style='margin: 0; color: white; font-size: 1.1rem;'>
            📊 실시간 재고 · 점도 · Lot 추적 통합 관리 시스템
        </h3>
        <p style='margin: 0.5rem 0 0 0; font-size: 0.9rem; opacity: 0.9;'>
            ✅ 대시보드 | ✅ 요약 | ✅ 재고관리 | ✅ 바인더 입출고 | ✅ 신규 입력 | ✅ 빠른검색
        </p>
    </div>
    """,
    unsafe_allow_html=True
)

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
    st.sidebar.info("업로드 파일(Lot 관리)로 실행 중입니다. (원본 자동 저장이 아니라, 업데이트 후 '다운로드'로 받는 방식이 안전합니다.)")

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
# Load Lot excel (없으면 빈 데이터라도 화면 표시)
# ==========================================================
lot_sig = file_sig(xlsx_path)
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
        headers=["일자", "바인더타입", "바인더명", "바인더 Lot", "반환량(kg)", "비고"],
    )
    data = load_dataframes(xlsx_path, lot_sig)
    binder_df = data["binder"].copy()
    single_df = data["single"].copy()
    spec_binder = data["spec_binder"].copy()
    spec_single = data["spec_single"].copy()
    base_lab = data["base_lab"].copy()
    binder_return_df = data["binder_return"].copy()

# normalize dates (Lot)
c_b_date = find_col(binder_df, "제조/입고일")
c_s_date = find_col(single_df, "입고일")
if c_b_date and c_b_date in binder_df.columns:
    binder_df[c_b_date] = binder_df[c_b_date].apply(normalize_date)
if c_s_date and c_s_date in single_df.columns:
    single_df[c_s_date] = single_df[c_s_date].apply(normalize_date)

# common cols (Single)
c_s_visc = find_col(single_df, "점도측정값(cP)")
c_s_judge = find_col(single_df, "점도판정")
c_s_lot = find_col(single_df, "단일색잉크 Lot")
c_s_blot = find_col(single_df, "사용된 바인더 Lot")
c_s_cg = find_col(single_df, "색상군")
c_s_pc = find_col(single_df, "제품코드")

# ==========================================================
# Tabs
# ==========================================================
tab_dash, tab_summary, tab_stock, tab_binder, tab_input, tab_search = st.tabs(
    ["📊 대시보드", "📌 요약", "📦 액상잉크 재고관리", "📦 바인더 입출고", "📝 신규 입력", "🔎 빠른검색"]
)

# ==========================================================
# Render: Summary (상사용 1장 요약)
# ==========================================================
def render_summary():
    st.markdown('<div class="section-title">📌 경영진 요약 리포트</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">핵심 KPI와 주요 지표를 한눈에 확인할 수 있습니다</div>', unsafe_allow_html=True)

    stock_ok = bool(stock_xlsx_path and Path(stock_xlsx_path).exists())
    product_to_color = build_product_to_color_map(spec_single, single_df)

    # ---------- 재고 KPI ----------
    inv_color = pd.DataFrame()
    use_color = pd.DataFrame()
    cov_alert = pd.DataFrame()
    stock_kpis = {}

    if stock_ok:
        stock_sig = file_sig(stock_xlsx_path)
        hist = load_stock_history(stock_xlsx_path, stock_sig, product_to_color)
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

            # 커버리지 (stock / avg daily use)
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

    # ---------- 점도 KPI ----------
    visc_ok = bool(
        c_s_date and c_s_visc and c_s_pc
        and (c_s_date in single_df.columns)
        and (c_s_visc in single_df.columns)
        and (c_s_pc in single_df.columns)
    )

    visc_kpis = {}
    daily_visc = pd.DataFrame()
    top_ng = pd.DataFrame()

    if visc_ok:
        df = single_df.copy()
        df[c_s_date] = pd.to_datetime(df[c_s_date], errors="coerce")
        df["_점도"] = pd.to_numeric(df[c_s_visc].astype(str).str.replace(",", "", regex=False), errors="coerce")
        df[c_s_pc] = df[c_s_pc].apply(normalize_product_code)

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

            # ✅ 수정된 부분: groupby 오류 해결
            daily_visc = (
                df30.groupby(df30[c_s_date].dt.date)
                .agg(mean_visc=("_점도", "mean"), cnt=("_점도", "size"))
                .reset_index()
                .rename(columns={c_s_date: "date"})
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

    # ---------- KPI Row ----------
    st.markdown("### 📊 핵심 성과 지표 (KPI)")
    
    a, b = st.columns(2)
    with a:
        st.markdown("#### 📦 재고 현황 (최근 30일)")
        if not stock_ok:
            st.info("💡 재고 파일이 없거나 읽지 못했습니다. 좌측 사이드바에서 재고 파일을 설정해주세요.")
        else:
            k1, k2 = st.columns(2)
            k1.metric("📅 최신 업데이트", stock_kpis["재고 최신일"])
            k2.metric("📊 총 재고량", f'{stock_kpis["현재 총 재고(kg)"]:,.1f} kg')
            
            k3, k4, k5 = st.columns(3)
            k3.metric("📉 30일 사용량", f'{stock_kpis["최근 30일 사용량(kg)"]:,.1f} kg')
            k4.metric("📥 입고 건수", f'{stock_kpis["최근 30일 입고(건)"]:,}')
            k5.metric("⚡ 일평균 사용", f'{stock_kpis["평균 사용량(kg/일)"]:,.1f} kg/일')

    with b:
        st.markdown("#### 🧪 품질 현황 (최근 30일)")
        if not visc_ok:
            st.info("💡 단일색 시트에 입고일/점도측정값/제품코드 컬럼이 필요합니다.")
        else:
            k1, k2 = st.columns(2)
            k1.metric("📅 최신 측정일", visc_kpis["점도 최신일"])
            k2.metric("🔬 총 측정 건수", f'{visc_kpis["최근 30일 측정(건)"]:,}')
            
            k3, k4 = st.columns(2)
            k3.metric("❌ 부적합", f'{visc_kpis["부적합(건)"]:,} 건')
            
            ng_rate = visc_kpis["부적합률(%)"]
            if ng_rate > 10:
                k4.markdown(f'<div class="status-badge-error">⚠️ 부적합률 {ng_rate:.1f}%</div>', unsafe_allow_html=True)
            elif ng_rate > 5:
                k4.markdown(f'<div class="status-badge-warning">⚠️ 부적합률 {ng_rate:.1f}%</div>', unsafe_allow_html=True)
            else:
                k4.markdown(f'<div class="status-badge-success">✅ 부적합률 {ng_rate:.1f}%</div>', unsafe_allow_html=True)

    st.divider()
    st.markdown("### 📈 핵심 시각화 차트")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown("**📦 현재 재고 현황 (색상별)**")
        if stock_ok and not inv_color.empty:
            ch = alt.Chart(inv_color).mark_bar().encode(
                y=alt.Y("color_group:N", sort="-x", title="색상 계열"),
                x=alt.X("kg:Q", title="재고량 (kg)"),
                color=alt.Color("color_group:N", scale=_color_scale_color_group(), legend=None),
                tooltip=[alt.Tooltip("color_group:N", title="색상계열"), alt.Tooltip("kg:Q", title="재고(kg)", format=",.1f")],
            ).properties(height=280)
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("📊 재고 데이터가 없습니다.")
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown("**🧪 점도 추이 (일별 평균)**")
        if visc_ok and not daily_visc.empty:
            ch = alt.Chart(daily_visc).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X("date:T", title="날짜"),
                y=alt.Y("mean_visc:Q", title="평균 점도 (cP)"),
                tooltip=[
                    alt.Tooltip("date:T", title="날짜"),
                    alt.Tooltip("mean_visc:Q", title="평균점도", format=",.0f"),
                    alt.Tooltip("cnt:Q", title="측정(건)", format=",.0f"),
                ],
            ).properties(height=280)
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("📊 점도 데이터가 없습니다.")
        st.markdown('</div>', unsafe_allow_html=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown("**📉 30일 사용량 (색상별)**")
        if stock_ok and not use_color.empty:
            ch = alt.Chart(use_color).mark_bar().encode(
                y=alt.Y("color_group:N", sort="-x", title="색상 계열"),
                x=alt.X("kg:Q", title="사용량 (kg)"),
                color=alt.Color("color_group:N", scale=_color_scale_color_group(), legend=None),
                tooltip=[alt.Tooltip("color_group:N", title="색상계열"), alt.Tooltip("kg:Q", title="사용량(kg)", format=",.1f")],
            ).properties(height=280)
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("📊 사용량 데이터가 없습니다.")
        st.markdown('</div>', unsafe_allow_html=True)

    with c4:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown("**⚠️ 부적합 다발 제품 (Top 8)**")
        if visc_ok and not top_ng.empty:
            ch = alt.Chart(top_ng).mark_bar(color='#ef4444').encode(
                y=alt.Y(f"{c_s_pc}:N", sort="-x", title="제품 코드"),
                x=alt.X("ng_cnt:Q", title="부적합 건수"),
                tooltip=[alt.Tooltip(f"{c_s_pc}:N", title="제품코드"), alt.Tooltip("ng_cnt:Q", title="부적합(건)", format=",.0f")],
            ).properties(height=280)
            st.altair_chart(ch, use_container_width=True)
        else:
            st.success("✅ 부적합 데이터가 없습니다.")
        st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("🔍 상세 분석: 재고 부족 경보 (커버리지 Top 10)"):
        if stock_ok and not cov_alert.empty:
            st.warning("⚠️ 다음 제품들은 재고 소진 위험이 있습니다. 발주를 검토해주세요.")
            show = cov_alert.copy()
            show["stock_kg"] = show["stock_kg"].round(1)
            show["avg_daily_use"] = show["avg_daily_use"].round(2)
            show["cover_days"] = show["cover_days"].round(1)
            st.dataframe(show, use_container_width=True, height=320)
        else:
            st.success("✅ 현재 재고 부족 위험 제품이 없습니다.")

# ==========================================================
# Render: Stock tab (재고/입고/사용량을 한 탭에서 보기 좋게)
# ==========================================================
def render_stock_tab():
    st.markdown('<div class="section-title">📦 액상잉크 재고관리</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">재고 현황 · 입고 추정 · 사용량 추이를 색상별로 분석합니다</div>', unsafe_allow_html=True)

    if not stock_xlsx_path or not Path(stock_xlsx_path).exists():
        st.error("❌ 재고 파일 경로가 올바르지 않습니다. 좌측 사이드바에서 재고 파일을 설정해주세요.")
        return

    product_to_color = build_product_to_color_map(spec_single, single_df)
    stock_sig = file_sig(stock_xlsx_path)
    hist = load_stock_history(stock_xlsx_path, stock_sig, product_to_color)
    if hist.empty:
        st.error("❌ 재고 엑셀을 읽지 못했습니다. (시트명: 1.15 형식 / 컬럼: 품목명, 금일 재고(kg), 하루 사용량(kg) 확인)")
        return

    min_d = hist["date"].min().date()
    max_d = hist["date"].max().date()

    # 필터 섹션
    st.markdown("### 🔍 조회 기간 및 필터")
    left, mid, right = st.columns([2.2, 2.8, 5.0])
    with left:
        quick = st.selectbox("📅 기간 선택", ["최근 7일", "최근 30일", "최근 90일", "전체", "직접 선택"], index=1)
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
            st.write(f"**📅 {start} ~ {end}**")
    with right:
        divisions = sorted([x for x in hist["division"].dropna().unique().tolist() if str(x).strip() and str(x).lower() not in ("nan", "none")])
        sel_div = st.multiselect("🏭 제품군 (PL/NPL/NSL 등)", divisions, default=divisions)

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
    inbound_kg = float(hist_f["inbound_kg"].sum())
    day_span = max(1, (end - start).days + 1)
    avg_daily_use = total_used / day_span if day_span else 0.0

    # KPI 카드
    st.markdown("### 📊 주요 지표")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("📅 최신 업데이트", latest_date.date().isoformat())
    k2.metric("📦 현재 총 재고", f"{total_stock:,.1f} kg")
    k3.metric("📉 기간 사용량", f"{total_used:,.1f} kg")
    k4.metric("📥 입고 건수", f"{inbound_events:,}")
    k5.metric("📦 입고량", f"{inbound_kg:,.1f} kg")
    k6.metric("⚡ 일평균 사용", f"{avg_daily_use:,.1f} kg")

    st.markdown('<div class="kpi-note">💡 입고는 "하루 사용량"이 음수로 기입된 경우(재고 증가)를 입고로 추정합니다.</div>', unsafe_allow_html=True)
    st.divider()

    # ---------- 1) 색상계열 바차트 ----------
    inv = latest_df.groupby("color_group", as_index=False)["curr_stock_kg"].sum().rename(columns={"curr_stock_kg": "kg"}).sort_values("kg", ascending=False)
    use = hist_f.groupby("color_group", as_index=False)["used_kg"].sum().rename(columns={"used_kg": "kg"}).sort_values("kg", ascending=False)
    inbound = hist_f.groupby("color_group", as_index=False)["inbound_kg"].sum().rename(columns={"inbound_kg": "kg"}).sort_values("kg", ascending=False)

    def bar_chart(df: pd.DataFrame, value_title: str):
        if df.empty:
            return None
        
        # 막대 차트
        bars = alt.Chart(df).mark_bar().encode(
            y=alt.Y("color_group:N", sort="-x", title=""),
            x=alt.X("kg:Q", title=value_title),
            color=alt.Color("color_group:N", scale=_color_scale_color_group(), legend=None),
            tooltip=[alt.Tooltip("color_group:N", title="색상계열"), alt.Tooltip("kg:Q", title=value_title, format=",.1f")],
        )
        
        # 텍스트 레이블 (막대 끝에 값 표시)
        text = alt.Chart(df).mark_text(
            align='left',
            baseline='middle',
            dx=3,  # 막대 끝에서 3px 오른쪽
            fontSize=11,
            fontWeight='bold'
        ).encode(
            y=alt.Y("color_group:N", sort="-x", title=""),
            x=alt.X("kg:Q"),
            text=alt.Text("kg:Q", format=",.1f"),
            color=alt.value('#333333')
        )
        
        return (bars + text).properties(height=260)

    st.markdown("### 📊 색상별 재고 · 사용량 · 입고 현황")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown("**1) 현재 재고 (최신일)**")
        ch = bar_chart(inv, "재고(kg)")
        if ch is not None:
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("📊 표시할 재고 데이터가 없습니다.")
        st.markdown('</div>', unsafe_allow_html=True)
        
    with c2:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown("**2) 기간 사용량**")
        ch = bar_chart(use, "사용량(kg)")
        if ch is not None:
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("📊 표시할 사용량 데이터가 없습니다.")
        st.markdown('</div>', unsafe_allow_html=True)
        
    with c3:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown("**3) 기간 입고 (추정)**")
        ch = bar_chart(inbound, "입고(kg)")
        if ch is not None:
            st.altair_chart(ch, use_container_width=True)
        else:
            st.info("📊 표시할 입고 데이터가 없습니다.")
        st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    # ---------- 2) 일별 사용량 추이 ----------
    st.markdown("### 📈 일별 사용량 추이 분석")
    present = [k for k in COLOR_KEYS if k in hist_f["color_group"].unique().tolist()]
    default_keys = [k for k in present if k != "OTHER"][:5] or present
    sel_keys = st.multiselect("🎨 표시할 색상계열", COLOR_KEYS, default=default_keys)

    daily = (
        hist_f[hist_f["color_group"].isin(sel_keys)]
        .groupby(["date", "color_group"], as_index=False)["used_kg"].sum()
    )
    total = hist_f.groupby("date", as_index=False)["used_kg"].sum().rename(columns={"used_kg": "TOTAL"})

    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    line = alt.Chart(daily).mark_line(point=True, strokeWidth=2.5).encode(
        x=alt.X("date:T", title="날짜"),
        y=alt.Y("used_kg:Q", title="사용량(kg)"),
        color=alt.Color("color_group:N", scale=_color_scale_color_group(), legend=alt.Legend(title="색상계열")),
        tooltip=[
            alt.Tooltip("date:T", title="날짜"),
            alt.Tooltip("color_group:N", title="색상계열"),
            alt.Tooltip("used_kg:Q", title="사용량(kg)", format=",.1f"),
        ],
    )
    total_line = alt.Chart(total).mark_line(point=True, strokeDash=[6, 3], strokeWidth=3, color='#374151').encode(
        x="date:T",
        y=alt.Y("TOTAL:Q", title="사용량(kg)"),
        tooltip=[alt.Tooltip("date:T", title="날짜"), alt.Tooltip("TOTAL:Q", title="TOTAL(kg)", format=",.1f")],
    )
    st.altair_chart((line + total_line).interactive(), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    # ---------- 3) 커버리지(발주 판단용) ----------
    st.markdown("### ⚠️ 재고 커버리지 분석 (발주 판단)")
    st.caption("💡 커버리지 = (최신 재고 kg) / (선택기간 평균 일사용량). 평균 사용량이 0이면 커버리지 계산 제외.")

    use_by_product = hist_f.groupby("product_code", as_index=False)["used_kg"].sum()
    use_by_product["avg_daily_use"] = use_by_product["used_kg"] / day_span
    stock_by_product = latest_df.groupby("product_code", as_index=False)["curr_stock_kg"].sum().rename(columns={"curr_stock_kg": "stock_kg"})

    cov = stock_by_product.merge(use_by_product[["product_code", "avg_daily_use"]], on="product_code", how="left")
    cov["avg_daily_use"] = cov["avg_daily_use"].fillna(0.0)
    cov["cover_days"] = cov.apply(lambda r: (r["stock_kg"] / r["avg_daily_use"]) if r["avg_daily_use"] > 0 else None, axis=1)
    cov["color_group"] = cov["product_code"].map(product_to_color).fillna("OTHER").apply(normalize_color_group)

    warn_days = st.slider("⚠️ 경보 기준 (일)", min_value=1, max_value=60, value=14, step=1)
    show_cov = cov[cov["cover_days"].notna()].copy()
    show_cov["cover_days"] = show_cov["cover_days"].round(1)
    show_cov["stock_kg"] = show_cov["stock_kg"].round(1)
    show_cov["avg_daily_use"] = show_cov["avg_daily_use"].round(2)

    alert = show_cov[show_cov["cover_days"] <= warn_days].sort_values("cover_days").head(30)
    if alert.empty:
        st.success(f"✅ 커버리지 {warn_days}일 이하 품목이 없습니다. 재고 상태가 양호합니다.")
    else:
        st.warning(f"⚠️ 커버리지 {warn_days}일 이하 품목이 **{len(alert)}개** 있습니다. 발주를 검토해주세요. (상위 30개 표시)")
        st.dataframe(alert, use_container_width=True, height=360)

# ==========================================================
# Render: Dashboard tab (Lot 쪽 전반 현황)
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

    f1, f2, f3, f4 = st.columns([1.2, 1.2, 2.2, 2.4])
    with f1:
        start = st.date_input("시작일", value=max(dmin, dmax - dt.timedelta(days=90)))
    with f2:
        end = st.date_input("종료일", value=dmax)
    with f3:
        pcs = sorted(df[c_s_pc].dropna().astype(str).unique().tolist())
        sel_pc = st.multiselect("제품코드", pcs, default=[])
    with f4:
        cg = sorted({normalize_color_group(x) for x in df[c_s_cg].dropna().tolist()}) if c_s_cg and c_s_cg in df.columns else []
        sel_cg = st.multiselect("색상계열", COLOR_KEYS, default=[])

    if start > end:
        start, end = end, start

    df = df[(df[c_s_date].dt.date >= start) & (df[c_s_date].dt.date <= end)]
    if sel_pc:
        df = df[df[c_s_pc].astype(str).isin(sel_pc)]
    if sel_cg and c_s_cg and c_s_cg in df.columns:
        df = df[df[c_s_cg].apply(normalize_color_group).isin(sel_cg)]

    view = pd.DataFrame({
        "입고일": df[c_s_date].dt.date,
        "색상군": df[c_s_cg].apply(normalize_color_group) if c_s_cg and (c_s_cg in df.columns) else None,
        "제품코드": df[c_s_pc].apply(normalize_product_code),
        "단일색Lot": df[c_s_lot] if c_s_lot and (c_s_lot in df.columns) else None,
        "사용바인더Lot": df[c_s_blot] if c_s_blot and (c_s_blot in df.columns) else None,
        "점도(cP)": pd.to_numeric(df[c_s_visc].astype(str).str.replace(",", "", regex=False), errors="coerce"),
        "점도판정": df[c_s_judge] if c_s_judge and (c_s_judge in df.columns) else None,
    }).dropna(subset=["입고일"]).sort_values("입고일", ascending=False)

    st.dataframe(view, use_container_width=True, height=360)

# ==========================================================
# Render: Binder IO tab (파일 업로드 + 구글시트)
# ==========================================================
def render_binder_io():
    st.subheader("바인더 입출고 내역 (파일 업로드 / 구글시트)")
    st.caption("✅ 바인더 입출고 내역 파일(.xlsx)을 업로드하면 업로드 즉시 표가 표시됩니다. 구글시트는 새로고침 시 자동 반영(캐시 60초).")

    # ---- file upload ----
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

    # ---- google sheets ----
    st.subheader("바인더 입출고 (Google Sheets 자동 반영)")
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
# Render: Input tab (엑셀에 저장)
# ==========================================================
def render_input_tab():
    st.markdown('<div class="section-title">📝 신규 입력</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">여기서 입력한 값은 해당 엑셀 시트에 바로 추가(append)됩니다.</div>', unsafe_allow_html=True)

    if not Path(xlsx_path).exists():
        st.error("Lot 관리 엑셀 파일이 없어서 저장할 수 없습니다. 좌측에서 업로드하거나 경로를 설정해 주세요.")
        return

    st.info("⚠️ 엑셀이 PC에서 열려 있으면 저장이 실패할 수 있습니다. 실패 시 엑셀을 닫고 다시 시도하거나, 업로드 파일로 실행 후 '다운로드'로 받는 방식이 안전합니다.")

    subt1, subt2, subt3 = st.tabs(["🧪 단일색 잉크 신규 입력", "🧴 바인더 제조/입고 신규 입력", "↩️ 바인더 업체반환 입력"])

    # ---------- 단일색 ----------
    with subt1:
        headers = get_sheet_headers(xlsx_path, SHEET_SINGLE)
        if not headers:
            st.error(f"시트/헤더를 읽지 못했습니다: {SHEET_SINGLE}")
            return

        existing_lots = []
        if c_s_lot and c_s_lot in single_df.columns:
            existing_lots = single_df[c_s_lot].dropna().astype(str).tolist()

        colA, colB = st.columns([1.6, 1.4])
        with colA:
            st.markdown("#### 입력")
            with st.form("form_single"):
                in_date = st.date_input("입고일", value=dt.date.today())
                product_code = st.text_input("제품코드", value="")
                color_group = st.selectbox("색상군", COLOR_KEYS, index=COLOR_KEYS.index("OTHER"))
                binder_lot = st.text_input("사용된 바인더 Lot", value="")

                auto_lot = st.checkbox("단일색잉크 Lot 자동 생성", value=True)
                if auto_lot:
                    lot_preview = next_lot(existing_lots, in_date, fallback_prefix="PCB")
                    single_lot = st.text_input("단일색잉크 Lot", value=lot_preview)
                else:
                    single_lot = st.text_input("단일색잉크 Lot", value="")

                visc = st.number_input("점도측정값(cP)", min_value=0, value=0, step=100)
                judge = st.selectbox("점도판정", ["적합", "부적합", ""], index=2)
                note = st.text_input("비고(있으면)", value="")

                submit = st.form_submit_button("✅ 저장(단일색_수입검사에 추가)", use_container_width=True)

        with colB:
            st.markdown("#### 저장될 형태(미리보기)")
            preview = {
                "입고일": in_date,
                "제품코드": normalize_product_code(product_code),
                "색상군": normalize_color_group(color_group),
                "사용된 바인더 Lot": binder_lot,
                "단일색잉크 Lot": single_lot,
                "점도측정값(cP)": visc if visc != 0 else None,
                "점도판정": judge if judge else None,
                "비고": note if note else None,
            }
            st.dataframe(pd.DataFrame([preview]), use_container_width=True, height=220)

        if submit:
            if not preview["입고일"] or not preview["제품코드"] or not preview["단일색잉크 Lot"]:
                st.error("입고일 / 제품코드 / 단일색잉크 Lot 는 필수입니다.")
            else:
                try:
                    append_row_to_xlsx(xlsx_path, SHEET_SINGLE, preview)
                    st.session_state["_toast_msg"] = "단일색_수입검사에 저장 완료"
                    st.cache_data.clear()
                    st.success("저장 완료 ✅")
                    download_xlsx_button(xlsx_path)
                    st.rerun()
                except PermissionError:
                    st.error("저장 실패: 엑셀 파일이 열려 있거나 잠겨 있습니다. 엑셀을 닫고 다시 시도하세요.")
                    download_xlsx_button(xlsx_path)
                except Exception as e:
                    st.error("저장 실패")
                    st.exception(e)

    # ---------- 바인더 제조/입고 ----------
    with subt2:
        headers = get_sheet_headers(xlsx_path, SHEET_BINDER)
        if not headers:
            st.error(f"시트/헤더를 읽지 못했습니다: {SHEET_BINDER}")
            return

        c_lot_b = find_col(binder_df, "Lot(자동)")
        existing_b_lots = []
        if c_lot_b and c_lot_b in binder_df.columns:
            existing_b_lots = binder_df[c_lot_b].dropna().astype(str).tolist()

        st.markdown("#### 입력")
        with st.form("form_binder"):
            b_date = st.date_input("제조/입고일", value=dt.date.today(), key="b_date_in")
            auto_lot_b = st.checkbox("바인더 Lot(자동) 자동 생성", value=True)
            if auto_lot_b:
                b_lot = st.text_input("Lot(자동)", value=next_lot(existing_b_lots, b_date, fallback_prefix="PLB"))
            else:
                b_lot = st.text_input("Lot(자동)", value="")

            b_judge = st.selectbox("판정", ["적합", "부적합", ""], index=2, key="b_judge_in")
            b_note = st.text_input("비고(있으면)", value="", key="b_note_in")

            with st.expander("추가 항목(시트에 컬럼이 있으면 같이 저장됨)"):
                extras = {}
                candidates = ["바인더타입", "바인더명", "제조량(kg)", "점도(cP)", "점도", "투입량(kg)", "담당", "원료 Lot", "원료Lot"]
                for name in candidates:
                    if name in headers:
                        if ("kg" in name.lower()) or ("량" in name) or ("투입" in name):
                            extras[name] = st.number_input(name, min_value=0.0, value=0.0, step=1.0)
                        else:
                            extras[name] = st.text_input(name, value="")

            submit_b = st.form_submit_button("✅ 저장(바인더_제조_입고에 추가)", use_container_width=True)

        if submit_b:
            row = {
                "제조/입고일": b_date,
                "Lot(자동)": b_lot,
                "판정": b_judge if b_judge else None,
                "비고": b_note if b_note else None,
            }
            for k, v in extras.items():
                if isinstance(v, (int, float)) and v == 0:
                    row[k] = None
                else:
                    row[k] = v if str(v).strip() else None

            if not row.get("제조/입고일") or not row.get("Lot(자동)"):
                st.error("제조/입고일 / Lot(자동) 은 필수입니다.")
            else:
                try:
                    append_row_to_xlsx(xlsx_path, SHEET_BINDER, row)
                    st.session_state["_toast_msg"] = "바인더_제조_입고에 저장 완료"
                    st.cache_data.clear()
                    st.success("저장 완료 ✅")
                    download_xlsx_button(xlsx_path)
                    st.rerun()
                except PermissionError:
                    st.error("저장 실패: 엑셀 파일이 열려 있거나 잠겨 있습니다. 엑셀을 닫고 다시 시도하세요.")
                    download_xlsx_button(xlsx_path)
                except Exception as e:
                    st.error("저장 실패")
                    st.exception(e)

        st.divider()
        st.markdown("#### 최근 바인더 기록(상위 30)")
        st.dataframe(binder_df.tail(30).iloc[::-1], use_container_width=True, height=320)

    # ---------- 바인더 업체반환 ----------
    with subt3:
        headers = get_sheet_headers(xlsx_path, SHEET_BINDER_RETURN)
        if not headers:
            st.error(f"시트/헤더를 읽지 못했습니다: {SHEET_BINDER_RETURN}")
            return

        st.markdown("#### 입력")
        with st.form("form_return"):
            r_date = st.date_input("일자", value=dt.date.today(), key="r_date")
            r_type = st.text_input("바인더타입", value="")
            r_name = st.text_input("바인더명", value="")
            r_lot = st.text_input("바인더 Lot", value="")
            r_qty = st.number_input("반환량(kg)", min_value=0.0, value=0.0, step=1.0)
            r_note = st.text_input("비고", value="")

            submit_r = st.form_submit_button("✅ 저장(바인더_업체반환에 추가)", use_container_width=True)

        if submit_r:
            row = {
                "일자": r_date,
                "바인더타입": r_type if r_type else None,
                "바인더명": r_name if r_name else None,
                "바인더 Lot": r_lot if r_lot else None,
                "반환량(kg)": r_qty if r_qty != 0 else None,
                "비고": r_note if r_note else None,
            }
            if not row["일자"] or not row["바인더 Lot"] or row["반환량(kg)"] is None:
                st.error("일자 / 바인더 Lot / 반환량(kg)은 필수입니다.")
            else:
                try:
                    append_row_to_xlsx(xlsx_path, SHEET_BINDER_RETURN, row)
                    st.session_state["_toast_msg"] = "바인더_업체반환에 저장 완료"
                    st.cache_data.clear()
                    st.success("저장 완료 ✅")
                    download_xlsx_button(xlsx_path)
                    st.rerun()
                except PermissionError:
                    st.error("저장 실패: 엑셀 파일이 열려 있거나 잠겨 있습니다. 엑셀을 닫고 다시 시도하세요.")
                    download_xlsx_button(xlsx_path)
                except Exception as e:
                    st.error("저장 실패")
                    st.exception(e)

        st.divider()
        st.markdown("#### 최근 반환 기록(상위 30)")
        st.dataframe(binder_return_df.tail(30).iloc[::-1], use_container_width=True, height=320)

# ==========================================================
# Render: Search tab
# ==========================================================
def render_search():
    st.subheader("빠른검색")
    mode = st.selectbox("검색 종류", ["바인더 Lot", "단일색 Lot", "제품코드"])
    q = st.text_input("검색어", placeholder="예: PCB20250112-01 / PLB25041501 / PL-835-1 ...")

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
        mask = None
        for c in cols:
            if c and c in df.columns:
                m = df[c].astype(str).str.contains(t, case=False, na=False)
                mask = m if mask is None else (mask | m)
        if mask is None:
            return df.iloc[0:0]
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

with tab_input:
    render_input_tab()

with tab_search:
    render_search()
