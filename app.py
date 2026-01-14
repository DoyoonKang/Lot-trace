import altair as alt
import streamlit as st
import pandas as pd
import datetime as dt
import re
from pathlib import Path
from openpyxl import load_workbook
import requests
from io import StringIO

# ==========================================================
# Page Config (딱 1번만)
# ==========================================================
st.set_page_config(
    page_title="액상 잉크 Lot 추적 관리",
    page_icon="🧪",
    layout="wide",
)

# ==========================================================
# Google Sheets (Public) Reader
# ==========================================================
@st.cache_data(ttl=60, show_spinner=False)  # 60초마다 최신값 갱신
def read_gsheet_csv(sheet_id: str, sheet_name: str) -> pd.DataFrame:
    """Public/Link-shared Google Sheet 를 CSV로 읽어옵니다."""
    base = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
    r = requests.get(base, params={"tqx": "out:csv", "sheet": sheet_name}, timeout=20)
    r.raise_for_status()
    r.encoding = "utf-8"
    return pd.read_csv(StringIO(r.text))

# ==========================================================
# Config
# ==========================================================
DEFAULT_XLSX = "액상잉크_Lot추적관리_FINAL.xlsx"

SHEET_BINDER = "바인더_제조_입고"
SHEET_SINGLE = "단일색_수입검사"
SHEET_SPEC_BINDER = "Spec_Binder"
SHEET_SPEC_SINGLE = "Spec_Single_H&S"
SHEET_BASE_LAB = "기준LAB"
SHEET_BINDER_RETURN = "바인더_업체반환"  # kg 단위 반환 기록(없으면 자동 생성)

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

# ==========================================================
# Helpers
# ==========================================================
def norm_key(x) -> str:
    """헤더/컬럼 비교용: 줄바꿈 제거 + 공백 정리"""
    if x is None:
        return ""
    s = str(x)
    s = s.replace("\n", " ").replace("\r", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def find_col(df: pd.DataFrame, want: str) -> str | None:
    """df에서 want(줄바꿈/공백 무시)와 동일한 컬럼명을 찾아 반환"""
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

def delta_e76(lab1, lab2) -> float:
    return float(((lab1[0]-lab2[0])**2 + (lab1[1]-lab2[1])**2 + (lab1[2]-lab2[2])**2) ** 0.5)

def ensure_sheet_exists(xlsx_path: str, sheet_name: str, headers: list[str]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        wb.save(xlsx_path)

def append_row_to_sheet(xlsx_path: str, sheet_name: str, row: dict):
    """
    엑셀 헤더(1행) 기준으로 append.
    row dict는 '헤더 원문' 또는 norm_key(헤더) 키로 값이 있으면 채움.
    """
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    values = []
    for h in headers:
        if h is None:
            values.append(None)
            continue
        v = row.get(h, None)
        if v is None:
            v = row.get(norm_key(h), None)
        values.append(v)
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
            if h is None:
                values.append(None)
                continue
            v = row.get(h, None)
            if v is None:
                v = row.get(norm_key(h), None)
            values.append(v)
        ws.append(values)
    wb.save(xlsx_path)

def update_sheet_cells(xlsx_path: str, sheet_name: str, updates: list[tuple[int, str, object]]):
    """
    updates: (excel_row_number, header_name, value)
    header_name은 시트 1행에 있는 헤더와 동일해야 함.
    """
    if not updates:
        return
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    header_map = {}
    for j, cell in enumerate(ws[1], start=1):
        header_map[str(cell.value)] = j

    for r, h, v in updates:
        if h not in header_map:
            # 헤더가 완전히 동일하지 않은 경우 norm_key로 한 번 더 매칭
            for hh, col_j in header_map.items():
                if norm_key(hh) == norm_key(h):
                    header_map[h] = col_j
                    break
        if h not in header_map:
            continue
        col = header_map[h]

        # 날짜/시간 처리
        if isinstance(v, pd.Timestamp):
            v = v.to_pydatetime()
        ws.cell(row=int(r), column=int(col)).value = v

    wb.save(xlsx_path)

@st.cache_data(show_spinner=False)
def load_dataframes(xlsx_path: str) -> dict[str, pd.DataFrame]:
    """pandas로 시트 읽기(표시/분석용)."""
    def read(name: str) -> pd.DataFrame:
        return pd.read_excel(xlsx_path, sheet_name=name)

    out = {
        "binder": read(SHEET_BINDER),
        "single": read(SHEET_SINGLE),
        "spec_binder": read(SHEET_SPEC_BINDER),
        "spec_single": read(SHEET_SPEC_SINGLE),
        "base_lab": read(SHEET_BASE_LAB),
    }
    # 반환 시트는 없을 수도 있음
    try:
        out["binder_return"] = read(SHEET_BINDER_RETURN)
    except Exception:
        out["binder_return"] = pd.DataFrame(columns=["일자", "바인더타입", "바인더명", "바인더 Lot", "반환량(kg)", "비고"])
    return out

def infer_binder_name_from_lot(spec_binder: pd.DataFrame, binder_lot: str):
    """Lot prefix 규칙으로 바인더명(=BinderType 역할) 추정."""
    if not binder_lot:
        return None
    lot = str(binder_lot).strip()
    c_name = find_col(spec_binder, "바인더명")
    c_rule = find_col(spec_binder, "Lot부여규칙")
    if not c_name or not c_rule:
        return None

    rules = spec_binder[[c_name, c_rule]].dropna().drop_duplicates().to_dict("records")
    for r in rules:
        rule = str(r[c_rule])
        m = re.match(r"^([A-Za-z0-9]+)\+", rule)
        if m and lot.startswith(m.group(1)):
            return r[c_name]
    return None

def get_binder_limits(spec_binder: pd.DataFrame, binder_name: str):
    c_name = find_col(spec_binder, "바인더명")
    c_item = find_col(spec_binder, "시험항목")
    c_lo = find_col(spec_binder, "하한")
    c_hi = find_col(spec_binder, "상한")
    c_rule = find_col(spec_binder, "Lot부여규칙")
    if not all([c_name, c_item, c_lo, c_hi]):
        return None, None, None, None

    df = spec_binder[spec_binder[c_name] == binder_name].copy()
    visc = df[df[c_item].astype(str).str.contains("점도", na=False)]
    uv = df[df[c_item].astype(str).str.contains("UV", na=False)]

    visc_lo = safe_to_float(visc[c_lo].dropna().iloc[0]) if len(visc[c_lo].dropna()) else None
    visc_hi = safe_to_float(visc[c_hi].dropna().iloc[0]) if len(visc[c_hi].dropna()) else None
    uv_hi = safe_to_float(uv[c_hi].dropna().iloc[0]) if len(uv[c_hi].dropna()) else None
    rule = df[c_rule].dropna().iloc[0] if c_rule and len(df[c_rule].dropna()) else None
    return visc_lo, visc_hi, uv_hi, rule

def next_seq_for_pattern(existing_lots: pd.Series, prefix: str, date_str: str, sep: str = "-") -> int:
    lots = existing_lots.dropna().astype(str).tolist()
    seqs = []
    for lot in lots:
        lot = str(lot).strip()
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
        seq = next_seq_for_pattern(existing_binder_lots, prefix, date_str, sep="-")
        return f"{prefix}{date_str}-{seq:02d}"
    return f"{prefix}{date_str}"

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

    c_lot = find_col(single_df, "단일색잉크 Lot")
    lots = single_df[c_lot].dropna().astype(str).tolist() if c_lot else []
    seqs = []
    for lot in lots:
        lot = str(lot).strip()
        if lot.startswith(patt_prefix):
            rest = lot[len(patt_prefix):]
            m = re.match(r"^(\d{2,})", rest)
            if m:
                seqs.append(int(m.group(1)))
    seq = (max(seqs) + 1) if seqs else 1
    return f"{patt_prefix}{seq:02d}"

def judge_range(value, lo, hi):
    v = safe_to_float(value)
    if v is None:
        return None
    if lo is not None and v < float(lo):
        return "부적합"
    if hi is not None and v > float(hi):
        return "부적합"
    return "적합"

def extract_de76_from_note(note: str | None):
    if not note:
        return None
    m = re.search(r"\[\s*ΔE76\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*\]", str(note))
    if not m:
        return None
    try:
        return float(m.group(1))
    except Exception:
        return None

def compute_de76_series(single_df: pd.DataFrame, base_lab: pd.DataFrame) -> pd.Series:
    """비고의 [ΔE76=..] 또는 착색력/기준LAB로 ΔE76 계산."""
    c_note = find_col(single_df, "비고")
    out = pd.Series([None] * len(single_df), index=single_df.index, dtype="float")

    if c_note:
        for idx, val in single_df[c_note].items():
            de = extract_de76_from_note(None if pd.isna(val) else str(val))
            if de is not None:
                out.loc[idx] = de

    # 착색력 기반 계산
    c_pc = find_col(single_df, "제품코드")
    cL = find_col(single_df, "착색력_L*")
    ca = find_col(single_df, "착색력_a*")
    cb = find_col(single_df, "착색력_b*")
    b_pc = find_col(base_lab, "제품코드")
    bL = find_col(base_lab, "기준_L*")
    ba = find_col(base_lab, "기준_a*")
    bb = find_col(base_lab, "기준_b*")
    if not all([c_pc, cL, ca, cb, b_pc, bL, ba, bb]):
        return out

    base = base_lab.copy()
    base[b_pc] = base[b_pc].astype(str).str.strip()
    base_map = base.set_index(b_pc)[[bL, ba, bb]].to_dict("index")

    for idx, row in single_df.iterrows():
        if pd.notna(out.loc[idx]):
            continue
        pc = row.get(c_pc, None)
        if pd.isna(pc):
            continue
        pc = str(pc).strip()
        if pc not in base_map:
            continue
        L = safe_to_float(row.get(cL))
        a = safe_to_float(row.get(ca))
        b = safe_to_float(row.get(cb))
        if None in (L, a, b):
            continue
        ref = base_map[pc]
        rL = safe_to_float(ref[bL]); ra = safe_to_float(ref[ba]); rb = safe_to_float(ref[bb])
        if None in (rL, ra, rb):
            continue
        out.loc[idx] = delta_e76((L, a, b), (rL, ra, rb))
    return out

def safe_date_bounds(series: pd.Series):
    s = pd.to_datetime(series, errors="coerce").dropna()
    if len(s) == 0:
        today = dt.date.today()
        return today, today
    return s.min().date(), s.max().date()

def detect_date_col(df: pd.DataFrame):
    for c in df.columns:
        ck = norm_key(c)
        if any(k in ck for k in ["일자", "날짜", "date", "입고일", "출고일"]):
            return c
    return None

def add_excel_row_number(df: pd.DataFrame) -> pd.DataFrame:
    """엑셀 1행이 헤더라고 가정할 때, 데이터 row 번호 = index+2."""
    df = df.copy()
    df["_excel_row"] = df.index + 2
    return df

# ==========================================================
# UI Header
# ==========================================================
st.title("액상 잉크 Lot 추적 관리 대시보드")
st.caption("✅ 대시보드(목록/평균/추이)  |  ✅ 잉크 입고(엑셀 누적)  |  ✅ 바인더 입출고(구글시트 최신순)  |  ✅ 반품(kg) 기록  |  ✅ 빠른검색/수정")

# ==========================================================
# Data file selection
# ==========================================================
with st.sidebar:
    st.header("데이터 파일")
    xlsx_path = st.text_input("엑셀 파일 경로", value=DEFAULT_XLSX)
    uploaded = st.file_uploader("또는 엑셀 업로드(업로드 모드: 서버 저장 보장 X)", type=["xlsx"])

# 업로드 파일은 "처음 1회만" tmp로 복사 (저장한 내용이 rerun 때 덮어써지는 문제 방지)
if uploaded is not None:
    upload_sig = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.get("_uploaded_sig") != upload_sig:
        tmp_path = Path(".streamlit_tmp.xlsx")
        tmp_path.write_bytes(uploaded.getvalue())
        st.session_state["_uploaded_sig"] = upload_sig
        st.session_state["_tmp_xlsx_path"] = str(tmp_path)
    xlsx_path = st.session_state.get("_tmp_xlsx_path", xlsx_path)
    st.sidebar.info("업로드 파일로 실행 중입니다. (서버 재시작 시 누적이 보장되지 않습니다.)")

if not Path(xlsx_path).exists():
    st.error(f"엑셀 파일을 찾을 수 없습니다: {xlsx_path}")
    st.stop()

# 반환 시트 없으면 생성
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

# 컬럼 참조(실제 이름)
c_b_date = find_col(binder_df, "제조/입고일")
c_s_date = find_col(single_df, "입고일")

# 날짜 정리
if c_b_date:
    binder_df[c_b_date] = binder_df[c_b_date].apply(normalize_date)
if c_s_date:
    single_df[c_s_date] = single_df[c_s_date].apply(normalize_date)

# 대시보드 파생
c_s_visc = find_col(single_df, "점도측정값(cP)")
c_s_lot = find_col(single_df, "단일색잉크 Lot")
c_s_blot = find_col(single_df, "사용된 바인더 Lot")
c_s_cg = find_col(single_df, "색상군")
c_s_pc = find_col(single_df, "제품코드")

# ΔE76
single_df["_ΔE76"] = compute_de76_series(single_df, base_lab)

# tabs
tab_dash, tab_ink_in, tab_binder, tab_search = st.tabs(
    ["📊 대시보드", "✍️ 잉크 입고", "📦 바인더 입출고", "🔎 빠른검색/수정"]
)

# ==========================================================
# Dashboard
# ==========================================================
with tab_dash:
    # KPI
    b_total = len(binder_df)
    s_total = len(single_df)
    c_b_judge = find_col(binder_df, "판정")
    c_s_judge = find_col(single_df, "점도판정")
    b_ng = int((binder_df[c_b_judge] == "부적합").sum()) if c_b_judge else 0
    s_ng = int((single_df[c_s_judge] == "부적합").sum()) if c_s_judge else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("바인더 기록", f"{b_total:,}")
    c2.metric("바인더 부적합", f"{b_ng:,}")
    c3.metric("단일색 기록", f"{s_total:,}")
    c4.metric("단일색(점도) 부적합", f"{s_ng:,}")

    st.divider()

    # 1) 목록(엑셀형)
    st.subheader("1) 단일색 데이터 목록 (엑셀형)")
    need = [c_s_date, c_s_cg, c_s_pc, c_s_blot, c_s_visc]
    if any(c is None for c in need):
        st.warning("단일색 시트에서 필요한 컬럼을 찾지 못했습니다. (입고일/색상군/제품코드/사용된 바인더 Lot/점도측정값)")
    else:
        df_list = single_df.copy()
        df_list[c_s_date] = pd.to_datetime(df_list[c_s_date], errors="coerce")
        dmin, dmax = safe_date_bounds(df_list[c_s_date])

        f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
        with f1:
            start = st.date_input("시작일(목록)", value=max(dmin, dmax - dt.timedelta(days=90)), key="list_start")
        with f2:
            end = st.date_input("종료일(목록)", value=dmax, key="list_end")
        with f3:
            cg_opts = sorted([x for x in df_list[c_s_cg].dropna().unique().tolist()]) if c_s_cg else []
            cg = st.multiselect("색상군(목록)", cg_opts, key="list_cg")
        with f4:
            pc_opts = sorted([x for x in df_list[c_s_pc].dropna().unique().tolist()]) if c_s_pc else []
            pc = st.multiselect("제품코드(목록)", pc_opts, key="list_pc")

        if start > end:
            start, end = end, start

        df_list = df_list[(df_list[c_s_date].dt.date >= start) & (df_list[c_s_date].dt.date <= end)]
        if cg and c_s_cg:
            df_list = df_list[df_list[c_s_cg].isin(cg)]
        if pc and c_s_pc:
            df_list = df_list[df_list[c_s_pc].isin(pc)]

        view = pd.DataFrame({
            "제조일자": df_list[c_s_date].dt.date,
            "색상군": df_list[c_s_cg] if c_s_cg else None,
            "제품코드": df_list[c_s_pc] if c_s_pc else None,
            "사용된바인더": df_list[c_s_blot] if c_s_blot else None,
            "점도(cP)": pd.to_numeric(df_list[c_s_visc].astype(str).str.replace(",", "", regex=False), errors="coerce") if c_s_visc else None,
            "색차(ΔE76)": df_list["_ΔE76"],
        }).dropna(subset=["제조일자"]).sort_values(by="제조일자", ascending=False)

        st.dataframe(view, use_container_width=True, height=320)

        st.divider()

        st.subheader("1-1) 색상군별 평균 점도 (점 + 값)")
        mean_df = (
            view.dropna(subset=["색상군", "점도(cP)"])
            .groupby("색상군", as_index=False)["점도(cP)"]
            .mean()
            .rename(columns={"점도(cP)": "평균점도(cP)"})
        )
        if len(mean_df) == 0:
            st.info("표시할 평균 점도 데이터가 없습니다.")
        else:
            mean_df["표시"] = mean_df["평균점도(cP)"].round(0).astype("Int64").astype(str)
            base = alt.Chart(mean_df).encode(
                x=alt.X("색상군:N", sort=sorted(mean_df["색상군"].unique().tolist()), title="색상군"),
                y=alt.Y("평균점도(cP):Q", title="평균 점도(cP)"),
                tooltip=["색상군:N", "평균점도(cP):Q"],
            )
            pts = base.mark_circle(size=240)
            lbl = base.mark_text(dx=10, dy=-10).encode(text="표시:N")
            st.altair_chart((pts + lbl).interactive(), use_container_width=True)

    st.divider()

    # 2) Lot별 추이
    st.subheader("2) 단일색 점도 변화 추이 (Lot별)")
    st.caption("선택한 Lot별로 입고일 기준 점도 변화를 확인합니다. (점/라벨 크게 표시)")

    if not all([c_s_date, c_s_lot, c_s_visc]):
        st.warning("단일색 시트에서 추이 그래프에 필요한 컬럼(입고일/단일색잉크 Lot/점도측정값)을 찾지 못했습니다.")
    else:
        df = single_df.copy()
        df[c_s_date] = pd.to_datetime(df[c_s_date], errors="coerce")
        df["_점도"] = pd.to_numeric(df[c_s_visc].astype(str).str.replace(",", "", regex=False), errors="coerce")
        df["_Lot"] = df[c_s_lot].astype(str).replace("nan", "").replace("None", "").str.strip()
        df = df.dropna(subset=[c_s_date, "_점도"])
        df = df[df["_Lot"] != ""]

        if len(df) == 0:
            st.info("입고일/점도/Lot 값이 비어있어 추이 그래프를 표시할 수 없습니다.")
        else:
            dmin, dmax = safe_date_bounds(df[c_s_date])

            f1, f2, f3, f4, f5 = st.columns([1.2, 1.2, 1.6, 2.0, 1.0])
            with f1:
                start = st.date_input("시작일(추이)", value=max(dmin, dmax - dt.timedelta(days=90)), key="trend_start")
            with f2:
                end = st.date_input("종료일(추이)", value=dmax, key="trend_end")
            with f3:
                cg_opts = sorted([x for x in df[c_s_cg].dropna().unique().tolist()]) if c_s_cg else []
                cg = st.multiselect("색상군(추이)", cg_opts, key="trend_cg")
            with f4:
                pc_opts = sorted([x for x in df[c_s_pc].dropna().unique().tolist()]) if c_s_pc else []
                pc = st.multiselect("제품코드(추이)", pc_opts, key="trend_pc")
            with f5:
                show_labels = st.checkbox("라벨 표시", value=True, key="trend_labels")

            if start > end:
                start, end = end, start
            df = df[(df[c_s_date].dt.date >= start) & (df[c_s_date].dt.date <= end)]
            if cg and c_s_cg:
                df = df[df[c_s_cg].isin(cg)]
            if pc and c_s_pc:
                df = df[df[c_s_pc].isin(pc)]

            lot_list = sorted(df["_Lot"].dropna().unique().tolist())
            default_pick = lot_list[-5:] if len(lot_list) > 5 else lot_list
            pick = st.multiselect("표시할 단일색 Lot(복수 선택)", lot_list, default=default_pick, key="trend_lots")
            if pick:
                df = df[df["_Lot"].isin(pick)]

            if len(df) == 0:
                st.info("선택한 조건에 해당하는 데이터가 없습니다.")
            else:
                df = df.sort_values(c_s_date)
                df["_표시"] = df["_점도"].round(0).astype("Int64").astype(str)

                tooltip = ["입고일:T", "_Lot:N", "_점도:Q"]
                if c_s_pc:
                    tooltip.insert(2, f"{c_s_pc}:N")
                if c_s_cg:
                    tooltip.insert(3, f"{c_s_cg}:N")
                if c_s_blot:
                    tooltip.append(f"{c_s_blot}:N")

                base = alt.Chart(df).encode(
                    x=alt.X(f"{c_s_date}:T", title="입고일"),
                    y=alt.Y("_점도:Q", title="점도(cP)"),
                    tooltip=tooltip,
                )
                line = base.mark_line()
                points = base.mark_point(size=260).encode(color=alt.Color("_Lot:N", title="Lot"))
                if show_labels:
                    labels = base.mark_text(dy=-12).encode(text="_표시:N", color=alt.Color("_Lot:N", legend=None))
                    chart = (line + points + labels).interactive()
                else:
                    chart = (line + points).interactive()

                st.altair_chart(chart, use_container_width=True)

    st.divider()

    # 3) 제품(단일색)별 트렌드 + 스펙선 + 스펙 수정
    st.subheader("3) 제품별 점도 트렌드 (스펙 상/하한 빨간선 + 대시보드에서 스펙 수정)")
    st.caption("예: GREEN(PL-435)처럼 제품코드 기준으로 점도 추이를 보면서, 스펙 상/하한을 바로 조정할 수 있습니다.")

    if not all([c_s_date, c_s_visc, c_s_pc]):
        st.info("제품별 트렌드를 만들기 위해서는 단일색 시트에 입고일/점도측정값/제품코드 컬럼이 필요합니다.")
    else:
        dfp = single_df.copy()
        dfp[c_s_date] = pd.to_datetime(dfp[c_s_date], errors="coerce")
        dfp["_점도"] = pd.to_numeric(dfp[c_s_visc].astype(str).str.replace(",", "", regex=False), errors="coerce")
        dfp = dfp.dropna(subset=[c_s_date, "_점도", c_s_pc])

        prod_opts = sorted(dfp[c_s_pc].astype(str).dropna().unique().tolist())
        if len(prod_opts) == 0:
            st.info("제품코드 데이터가 없습니다.")
        else:
            cA, cB, cC = st.columns([1.4, 1.2, 1.4])
            with cA:
                prod = st.selectbox("제품코드 선택", prod_opts, key="prod_trend_pc")
            with cB:
                cg_val = None
                if c_s_cg:
                    cg_opts = sorted(dfp[dfp[c_s_pc].astype(str) == str(prod)][c_s_cg].dropna().unique().tolist())
                    cg_val = st.selectbox("색상군(선택)", ["(전체)"] + cg_opts, key="prod_trend_cg")
            with cC:
                # binder type(=바인더명 추정) 기반으로 스펙 선택 가능
                btypes = []
                if c_s_blot:
                    for x in dfp[dfp[c_s_pc].astype(str) == str(prod)][c_s_blot].dropna().astype(str).tolist():
                        bt = infer_binder_name_from_lot(spec_binder, x)
                        if bt:
                            btypes.append(bt)
                btypes = sorted(set(btypes))
                bt_val = st.selectbox("BinderType(자동/선택)", ["(자동/전체)"] + btypes, key="prod_trend_bt")

            dfp2 = dfp[dfp[c_s_pc].astype(str) == str(prod)].copy()
            if c_s_cg and cg_val and cg_val != "(전체)":
                dfp2 = dfp2[dfp2[c_s_cg] == cg_val]

            # 스펙 조회
            c_sp_cg = find_col(spec_single, "색상군")
            c_sp_pc = find_col(spec_single, "제품코드")
            c_sp_bt = find_col(spec_single, "BinderType")
            c_sp_lo = find_col(spec_single, "하한")
            c_sp_hi = find_col(spec_single, "상한")

            spec_lo = None
            spec_hi = None
            spec_row_excel = None

            if all([c_sp_pc, c_sp_lo, c_sp_hi]) and len(spec_single):
                hit = spec_single.copy()
                hit[c_sp_pc] = hit[c_sp_pc].astype(str).str.strip()
                hit = hit[hit[c_sp_pc] == str(prod).strip()]

                if c_s_cg and cg_val and cg_val != "(전체)" and c_sp_cg:
                    hit = hit[hit[c_sp_cg] == cg_val]

                if bt_val != "(자동/전체)" and c_sp_bt:
                    hit = hit[hit[c_sp_bt] == bt_val]

                if len(hit) >= 1:
                    # 첫 행 기준 (동일 조건 여러 행이면 사용자 데이터 구조상 1행이 정상)
                    spec_lo = safe_to_float(hit.iloc[0][c_sp_lo])
                    spec_hi = safe_to_float(hit.iloc[0][c_sp_hi])
                    # 엑셀 row 번호 (pandas index+2)
                    spec_row_excel = int(hit.index[0]) + 2

            # 차트
            if len(dfp2) == 0:
                st.info("선택 조건에 해당하는 데이터가 없습니다.")
            else:
                dfp2 = dfp2.sort_values(c_s_date)
                dfp2["_표시"] = dfp2["_점도"].round(0).astype("Int64").astype(str)

                base = alt.Chart(dfp2).encode(
                    x=alt.X(f"{c_s_date}:T", title="입고일"),
                    y=alt.Y("_점도:Q", title="점도(cP)"),
                    tooltip=[f"{c_s_date}:T", f"{c_s_pc}:N", "_점도:Q"] + ([f"{c_s_cg}:N"] if c_s_cg else []) + ([f"{c_s_blot}:N"] if c_s_blot else []),
                )
                line = base.mark_line()
                pts = base.mark_point(size=260)
                lbl = base.mark_text(dy=-12).encode(text="_표시:N")

                layers = [line, pts, lbl]

                # 스펙선(빨간선)
                if spec_lo is not None:
                    lo_df = pd.DataFrame({"y": [spec_lo]})
                    lo_rule = alt.Chart(lo_df).mark_rule(color="red").encode(y="y:Q")
                    layers.append(lo_rule)
                if spec_hi is not None:
                    hi_df = pd.DataFrame({"y": [spec_hi]})
                    hi_rule = alt.Chart(hi_df).mark_rule(color="red").encode(y="y:Q")
                    layers.append(hi_rule)

                st.altair_chart(alt.layer(*layers).interactive(), use_container_width=True)

            # 스펙 수정 UI
            with st.expander("스펙 상/하한 수정(Excel: Spec_Single_H&S)"):
                if spec_row_excel is None:
                    st.info("현재 선택 조건으로 Spec_Single_H&S에서 스펙 행을 찾지 못했습니다. (제품코드/색상군/BinderType 매칭 확인)")
                else:
                    cX, cY = st.columns(2)
                    with cX:
                        new_lo = st.number_input("새 하한", value=float(spec_lo) if spec_lo is not None else 0.0, step=10.0, format="%.1f", key="spec_edit_lo")
                    with cY:
                        new_hi = st.number_input("새 상한", value=float(spec_hi) if spec_hi is not None else 0.0, step=10.0, format="%.1f", key="spec_edit_hi")

                    if st.button("스펙 저장", type="primary", key="spec_save_btn"):
                        # 어떤 헤더명을 업데이트할지: 실제 엑셀 헤더 사용
                        updates = []
                        if c_sp_lo:
                            updates.append((spec_row_excel, c_sp_lo, float(new_lo)))
                        if c_sp_hi:
                            updates.append((spec_row_excel, c_sp_hi, float(new_hi)))
                        try:
                            update_sheet_cells(xlsx_path, SHEET_SPEC_SINGLE, updates)
                            st.success("스펙 저장 완료! (다시 계산/표시됩니다)")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"스펙 저장 실패: {e}")

    st.divider()

    st.subheader("최근 20건 (단일색)")
    show = single_df.copy()
    if c_s_date:
        show[c_s_date] = pd.to_datetime(show[c_s_date], errors="coerce")
        show = show.sort_values(by=c_s_date, ascending=False)
    st.dataframe(show.head(20), use_container_width=True)

    with st.expander("최근 데이터(단일색) 수정하기 (최대 50건)"):
        st.caption("실수로 입력된 값이 있으면 여기서 바로 수정 → '변경사항 저장'을 누르시면 엑셀에 반영됩니다.")
        if len(single_df) == 0:
            st.info("단일색 데이터가 없습니다.")
        else:
            edit_base = add_excel_row_number(show.head(50)).copy()
            # 편집용 컬럼만 노출(필요하면 추가)
            editable_cols = []
            for w in ["입고일", "잉크타입\n(HEMA/Silicone)", "색상군", "제품코드", "단일색잉크 Lot", "사용된 바인더 Lot",
                      "바인더제조처\n(내부/외주)", "BinderType(자동)", "점도측정값(cP)", "착색력_L*", "착색력_a*", "착색력_b*", "비고"]:
                c = find_col(edit_base, w)
                if c:
                    editable_cols.append(c)
            show_cols = ["_excel_row"] + editable_cols
            original = edit_base[show_cols].copy()

            edited = st.data_editor(
                original,
                use_container_width=True,
                num_rows="fixed",
                key="edit_recent_single",
                disabled=["_excel_row"],
            )

            if st.button("변경사항 저장(최근 50건)", type="primary", key="save_recent_single"):
                updates = []
                for i in range(len(original)):
                    excel_row = int(original.iloc[i]["_excel_row"])
                    for col in editable_cols:
                        before = original.iloc[i][col]
                        after = edited.iloc[i][col]
                        if (pd.isna(before) and pd.isna(after)) or (str(before) == str(after)):
                            continue
                        # 날짜 변환
                        if "일" in norm_key(col) and after is not None:
                            after = normalize_date(after)
                        updates.append((excel_row, col, after))
                if not updates:
                    st.info("변경된 값이 없습니다.")
                else:
                    try:
                        update_sheet_cells(xlsx_path, SHEET_SINGLE, updates)
                        st.success("수정사항 저장 완료!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"저장 실패: {e}")

# ==========================================================
# 잉크 입고 (단일색 입력)
# ==========================================================
with tab_ink_in:
    st.subheader("단일색 잉크 입력(입고)")
    st.info("이 탭은 **단일색_수입검사** 시트에 행을 추가(Append)하여 누적합니다. (동시 사용 시 충돌 가능)")

    ink_types = ["HEMA", "Silicone"]
    cg_col = find_col(spec_single, "색상군")
    pc_col = find_col(spec_single, "제품코드")
    bt_col = find_col(spec_single, "BinderType")
    lo_col = find_col(spec_single, "하한")
    hi_col = find_col(spec_single, "상한")

    color_groups = sorted(spec_single[cg_col].dropna().unique().tolist()) if cg_col else []
    product_codes = sorted(spec_single[pc_col].dropna().unique().tolist()) if pc_col else []

    c_blot = find_col(binder_df, "Lot(자동)")
    binder_lots = binder_df[c_blot].dropna().astype(str).tolist() if c_blot else []
    binder_lots = sorted(set([x.strip() for x in binder_lots if x.strip()]), reverse=True)

    with st.form("single_form", clear_on_submit=True):
        col1, col2, col3, col4 = st.columns([1.2, 1.3, 1.5, 2.0])
        with col1:
            in_date = st.date_input("입고일", value=dt.date.today(), key="single_in_date")
            ink_type = st.selectbox("잉크타입", ink_types, key="single_ink_type")
            color_group = st.selectbox("색상군", color_groups, key="single_cg") if color_groups else st.text_input("색상군", key="single_cg_text")
        with col2:
            product_code = st.selectbox("제품코드", product_codes, key="single_pc") if product_codes else st.text_input("제품코드", key="single_pc_text")
            binder_lot = st.selectbox("사용된 바인더 Lot", binder_lots, key="single_blot") if binder_lots else st.text_input("사용된 바인더 Lot", key="single_blot_text")
        with col3:
            visc_meas = st.number_input("점도측정값(cP)", min_value=0.0, step=1.0, format="%.1f", key="single_visc")
            supplier = st.selectbox("바인더제조처", ["내부", "외주"], index=0, key="single_supplier")
        with col4:
            st.caption("선택: 착색력(L*a*b*) 입력 시, 기준LAB이 있으면 ΔE(76)을 자동 계산해 '비고'에 기록합니다.")
            L = st.number_input("착색력_L*", value=0.0, step=0.1, format="%.2f", key="single_L")
            a = st.number_input("착색력_a*", value=0.0, step=0.1, format="%.2f", key="single_a")
            b = st.number_input("착색력_b*", value=0.0, step=0.1, format="%.2f", key="single_b")
            lab_enabled = st.checkbox("L*a*b* 입력함", value=False, key="single_lab_en")

        note = st.text_input("비고", value="", key="single_note")
        submit_s = st.form_submit_button("저장(단일색)")

    if submit_s:
        # binder type 추정
        binder_type = infer_binder_name_from_lot(spec_binder, binder_lot)

        # spec 조회 (color+product (+binderType))
        lo, hi = None, None
        visc_judge = None
        if all([cg_col, pc_col, lo_col, hi_col]) and len(spec_single):
            hit = spec_single[(spec_single[cg_col] == color_group) & (spec_single[pc_col] == product_code)].copy()
            if binder_type and bt_col and bt_col in hit.columns:
                hit = hit[hit[bt_col] == binder_type]
            if len(hit):
                lo = safe_to_float(hit.iloc[0][lo_col])
                hi = safe_to_float(hit.iloc[0][hi_col])
                visc_judge = judge_range(visc_meas, lo, hi)

        # lot 생성
        new_lot = generate_single_lot(single_df, product_code, color_group, in_date)
        if new_lot is None:
            st.error("단일색 Lot 자동 생성에 실패했습니다. (색상군 매핑 확인 필요)")
        else:
            note2 = note
            if lab_enabled:
                base_pc = find_col(base_lab, "제품코드")
                if base_pc:
                    base_hit = base_lab[base_lab[base_pc].astype(str).str.strip() == str(product_code).strip()]
                else:
                    base_hit = pd.DataFrame()

                bL = find_col(base_lab, "기준_L*")
                ba = find_col(base_lab, "기준_a*")
                bb = find_col(base_lab, "기준_b*")
                if len(base_hit) == 1 and all([bL, ba, bb]):
                    base_vals = (safe_to_float(base_hit.iloc[0][bL]), safe_to_float(base_hit.iloc[0][ba]), safe_to_float(base_hit.iloc[0][bb]))
                    if None not in base_vals:
                        de = delta_e76((float(L), float(a), float(b)), base_vals)
                        note2 = (note2 + " " if note2 else "") + f"[ΔE76={de:.2f}]"
                    else:
                        note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"
                else:
                    note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"

            # 엑셀 헤더에 맞춰 저장 (헤더 norm_key로 키 제공)
            row = {
                norm_key("입고일"): in_date,
                norm_key("잉크타입\n(HEMA/Silicone)"): ink_type,
                norm_key("색상군"): color_group,
                norm_key("제품코드"): product_code,
                norm_key("단일색잉크 Lot"): new_lot,
                norm_key("사용된 바인더 Lot"): binder_lot,
                norm_key("바인더제조처\n(내부/외주)"): supplier,
                norm_key("BinderType(자동)"): binder_type,
                norm_key("점도측정값(cP)"): float(visc_meas),
                norm_key("점도하한"): lo,
                norm_key("점도상한"): hi,
                norm_key("점도판정"): visc_judge,
                norm_key("착색력_L*"): float(L) if lab_enabled else None,
                norm_key("착색력_a*"): float(a) if lab_enabled else None,
                norm_key("착색력_b*"): float(b) if lab_enabled else None,
                norm_key("비고"): note2,
            }

            try:
                append_row_to_sheet(xlsx_path, SHEET_SINGLE, row)
                st.success(f"저장 완료! 단일색 Lot = {new_lot} / 점도판정 = {visc_judge}")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

# ==========================================================
# 바인더 입출고
# ==========================================================
with tab_binder:
    st.subheader("업체반환(반품) 입력 (kg 단위)")
    st.caption("※ 20kg(1통) 기준이더라도, 실제 반환량은 kg 단위로 입력합니다.")

    bname_col = find_col(spec_binder, "바인더명")
    binder_names = sorted(spec_binder[bname_col].dropna().unique().tolist()) if bname_col else []
    blot_col = find_col(binder_df, "Lot(자동)")
    binder_lots = binder_df[blot_col].dropna().astype(str).tolist() if blot_col else []
    binder_lots = sorted(set([x.strip() for x in binder_lots if x.strip()]), reverse=True)

    with st.form("binder_return_form", clear_on_submit=True):
        c1, c2, c3 = st.columns([1.2, 1.2, 2.6])
        with c1:
            r_date = st.date_input("반환일자", value=dt.date.today(), key="ret_date")
        with c2:
            r_type = st.selectbox("바인더타입", ["HEMA", "Silicone"], key="ret_type")
        with c3:
            r_name = st.selectbox("바인더명", binder_names, key="ret_name") if binder_names else st.text_input("바인더명", key="ret_name_text")

        c4, c5, c6 = st.columns([2.0, 1.2, 2.8])
        with c4:
            r_lot = st.selectbox("바인더 Lot(선택)", ["(직접입력)"] + binder_lots, key="ret_lot_sel")
            r_lot_text = st.text_input("바인더 Lot 직접입력", value="", key="ret_lot_text") if r_lot == "(직접입력)" else ""
            final_lot = r_lot_text.strip() if r_lot == "(직접입력)" else r_lot
        with c5:
            r_kg = st.number_input("반환량(kg)", min_value=0.0, step=0.5, format="%.1f", key="ret_kg")
        with c6:
            r_note = st.text_input("비고", value="", key="ret_note")

        submit_ret = st.form_submit_button("반품 저장")

    if submit_ret:
        if r_kg <= 0:
            st.error("반환량(kg)은 0보다 커야 합니다.")
        else:
            row = {
                "일자": r_date,
                "바인더타입": r_type,
                "바인더명": r_name,
                "바인더 Lot": final_lot,
                "반환량(kg)": float(r_kg),
                "비고": r_note,
            }
            try:
                append_row_to_sheet(xlsx_path, SHEET_BINDER_RETURN, row)
                st.success("반품 저장 완료!")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"반품 저장 실패: {e}")

    st.divider()
    st.subheader("바인더 입력 (제조/입고) — 여러 날짜/수량 묶음 입력 지원")

    input_mode = st.radio("입력 방식", ["개별 입력", "묶음 입력(여러 날짜/수량)"], horizontal=True, key="binder_input_mode")

    if input_mode == "개별 입력":
        with st.form("binder_form_single", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                mfg_date = st.date_input("제조/입고일", value=dt.date.today(), key="b_single_date")
                b_name = st.selectbox("바인더명", binder_names, key="b_single_name") if binder_names else st.text_input("바인더명", key="b_single_name_text")
            with col2:
                visc = st.number_input("점도(cP)", min_value=0.0, step=1.0, format="%.1f", key="b_single_visc")
                uv = st.number_input("UV흡광도(선택)", min_value=0.0, step=0.01, format="%.3f", key="b_single_uv")
                uv_enabled = st.checkbox("UV 값 입력함", value=False, key="b_single_uv_en")
            with col3:
                note = st.text_input("비고", value="", key="b_single_note")
                submit_b = st.form_submit_button("저장(바인더)")

        if submit_b:
            visc_lo, visc_hi, uv_hi, _ = get_binder_limits(spec_binder, b_name)
            lot = generate_binder_lot(spec_binder, b_name, mfg_date, binder_df.get(blot_col, pd.Series(dtype=str)) if blot_col else pd.Series(dtype=str))

            judge_v = judge_range(visc, visc_lo, visc_hi)
            judge_u = judge_range(uv if uv_enabled else None, None, uv_hi)
            judge = "부적합" if (judge_v == "부적합" or judge_u == "부적합") else "적합"

            row = {
                "제조/입고일": mfg_date,
                "바인더명": b_name,
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
        st.caption("표에 날짜/바인더명/수량(통)/점도/UV/비고를 입력하고, 한 번에 저장합니다.")
        base_rows = st.session_state.get("binder_batch_rows")
        if base_rows is None:
            d0 = dt.date.today()
            first_name = binder_names[0] if binder_names else ""
            base_rows = [
                {"제조/입고일": d0, "바인더명": first_name, "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
                {"제조/입고일": d0 - dt.timedelta(days=1), "바인더명": first_name, "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
            ]
        edit_df = st.data_editor(pd.DataFrame(base_rows), use_container_width=True, num_rows="dynamic", key="binder_batch_editor")
        submit_batch = st.button("묶음 저장(바인더)", type="primary", key="binder_batch_submit")

        if submit_batch:
            tmp = edit_df.copy()
            tmp["제조/입고일"] = tmp["제조/입고일"].apply(normalize_date)
            tmp["수량(통)"] = pd.to_numeric(tmp["수량(통)"], errors="coerce").fillna(0).astype(int)
            tmp["점도(cP)"] = pd.to_numeric(tmp["점도(cP)"].astype(str).str.replace(",", "", regex=False), errors="coerce")
            tmp = tmp.dropna(subset=["제조/입고일", "바인더명", "점도(cP)"])
            tmp = tmp[tmp["수량(통)"] > 0]
            if len(tmp) == 0:
                st.error("저장할 행이 없습니다. (날짜/바인더명/수량/점도 입력 확인)")
                st.stop()

            existing = binder_df.get(blot_col, pd.Series(dtype=str)) if blot_col else pd.Series(dtype=str)
            existing_list = existing.dropna().astype(str).tolist()
            seq_counters = {}
            rows_out = []

            tmp = tmp.sort_values(by="제조/입고일")
            for _, r in tmp.iterrows():
                mfg_date = r["제조/입고일"]
                b_name = str(r["바인더명"]).strip()
                qty = int(r["수량(통)"])
                visc = safe_to_float(r["점도(cP)"])
                uv_enabled = bool(r.get("UV입력", False))
                uv_val = safe_to_float(r.get("UV흡광도(선택)", None)) if uv_enabled else None
                note = str(r.get("비고", "")).strip()

                visc_lo, visc_hi, uv_hi, rule = get_binder_limits(spec_binder, b_name)
                m = re.match(r"^([A-Za-z0-9]+)\+YYYYMMDD(-##)?$", str(rule).strip()) if rule else None
                if not m:
                    st.error(f"[{b_name}] Lot부여규칙을 해석할 수 없습니다. (Spec_Binder 확인 필요)")
                    st.stop()

                prefix = m.group(1)
                has_seq = bool(m.group(2))
                date_str = mfg_date.strftime("%Y%m%d")

                if (not has_seq) and qty > 1:
                    st.error(f"[{b_name}] Lot부여규칙에 순번(-##)이 없어 여러 통(수량={qty})을 자동 생성할 수 없습니다.")
                    st.stop()

                key = (prefix, date_str)
                if key not in seq_counters:
                    seq_counters[key] = next_seq_for_pattern(pd.Series(existing_list), prefix, date_str, sep="-")

                for _i in range(qty):
                    if has_seq:
                        seq = seq_counters[key]
                        seq_counters[key] += 1
                        lot = f"{prefix}{date_str}-{seq:02d}"
                    else:
                        lot = f"{prefix}{date_str}"

                    judge_v = judge_range(visc, visc_lo, visc_hi)
                    judge_u = judge_range(uv_val, None, uv_hi) if uv_enabled else None
                    judge = "부적합" if (judge_v == "부적합" or judge_u == "부적합") else "적합"

                    rows_out.append({
                        "제조/입고일": mfg_date,
                        "바인더명": b_name,
                        "Lot(자동)": lot,
                        "점도(cP)": float(visc) if visc is not None else None,
                        "UV흡광도(선택)": float(uv_val) if uv_enabled and uv_val is not None else None,
                        "판정": judge,
                        "비고": note,
                    })
                    existing_list.append(lot)

            st.write("저장 미리보기(상위 50건)")
            st.dataframe(pd.DataFrame(rows_out).tail(50), use_container_width=True)

            try:
                append_rows_to_sheet(xlsx_path, SHEET_BINDER, rows_out)
                st.success(f"묶음 저장 완료! 총 {len(rows_out)}건 입력했습니다.")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    st.divider()
    st.subheader("바인더 입출고 (Google Sheets 자동 반영, 최신순)")
    st.caption("구글 시트를 수정하면 이 화면은 새로고침 시 자동 반영됩니다. (캐시 60초)")

    try:
        df_hema = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_HEMA)
        df_sil = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_SIL)
    except Exception as e:
        st.error("구글시트에서 데이터를 못 불러왔습니다. (시트 공유/웹게시/시트명/ID 확인)")
        st.exception(e)
        st.stop()

    for _df in [df_hema, df_sil]:
        dc = detect_date_col(_df)
        if dc:
            _df["_sort_date"] = pd.to_datetime(_df[dc], errors="coerce")
            _df.sort_values(by="_sort_date", ascending=False, inplace=True)
            _df.drop(columns=["_sort_date"], inplace=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### HEMA (최신순)")
        st.dataframe(df_hema, use_container_width=True, height=420)
    with c2:
        st.markdown("### Silicone (최신순)")
        st.dataframe(df_sil, use_container_width=True, height=420)

    if st.button("지금 최신값으로 다시 불러오기", key="binder_refresh"):
        st.cache_data.clear()
        st.rerun()

# ==========================================================
# 빠른검색 / 수정
# ==========================================================
with tab_search:
    st.subheader("빠른검색")
    st.caption("검색 결과를 바로 '수정 모드'로 열어서 잘못 입력된 데이터를 고칠 수 있습니다. (엑셀에 직접 반영)")

    edit_mode = st.checkbox("🔧 수정 모드 켜기(검색 결과를 편집 가능하게)", value=False, key="qs_edit_mode")

    c1, c2, c3 = st.columns([2, 2, 3])
    with c1:
        mode = st.selectbox("검색 종류", ["바인더 Lot", "단일색 잉크 Lot", "제품코드", "색상군", "기간(입고일)"])
    with c2:
        q = st.text_input("검색어", placeholder="예: PCB20250112-01 / PLB25041501 / PL-835-1 ...")
    with c3:
        st.write("")
        st.caption("💡 바인더 Lot 검색: 바인더 정보 + 해당 Lot를 사용한 단일색 잉크 목록")

    # 단일색 검색용 df
    s_df = single_df.copy()
    if c_s_date:
        s_df[c_s_date] = pd.to_datetime(s_df[c_s_date], errors="coerce")

    # 바인더 검색용 df
    b_df = binder_df.copy()
    if c_b_date:
        b_df[c_b_date] = pd.to_datetime(b_df[c_b_date], errors="coerce")

    def text_filter(df: pd.DataFrame, cols: list[str], text: str) -> pd.DataFrame:
        if not text:
            return df.iloc[0:0]  # 빈 결과
        t = str(text).strip()
        if not t:
            return df.iloc[0:0]
        mask = False
        for c in cols:
            if c and c in df.columns:
                mask = mask | df[c].astype(str).str.contains(t, case=False, na=False)
        return df[mask]

    if mode == "기간(입고일)":
        dmin, dmax = safe_date_bounds(s_df[c_s_date]) if c_s_date else (dt.date.today(), dt.date.today())
        d1, d2 = st.columns(2)
        with d1:
            start = st.date_input("시작일", value=max(dmin, dmax - dt.timedelta(days=30)), key="qs_start")
        with d2:
            end = st.date_input("종료일", value=dmax, key="qs_end")
        if start > end:
            start, end = end, start
        df_hit = s_df.copy()
        if c_s_date:
            df_hit = df_hit[(df_hit[c_s_date].dt.date >= start) & (df_hit[c_s_date].dt.date <= end)]
        st.subheader("단일색_수입검사")
        df_hit_show = add_excel_row_number(df_hit.sort_values(by=c_s_date, ascending=False) if c_s_date else df_hit)
        st.dataframe(df_hit_show, use_container_width=True)

        if edit_mode and len(df_hit_show) > 0:
            st.markdown("#### 🔧 검색 결과 수정")
            edited = st.data_editor(
                df_hit_show,
                use_container_width=True,
                num_rows="fixed",
                disabled=["_excel_row"],
                key="qs_edit_period",
            )
            if st.button("변경사항 저장(기간검색)", type="primary", key="qs_save_period"):
                # diff 저장
                updates = []
                # 비교는 문자열 기반으로 단순 비교(필요 시 강화 가능)
                for i in range(len(df_hit_show)):
                    excel_row = int(df_hit_show.iloc[i]["_excel_row"])
                    for col in df_hit_show.columns:
                        if col == "_excel_row":
                            continue
                        before = df_hit_show.iloc[i][col]
                        after = edited.iloc[i][col]
                        if (pd.isna(before) and pd.isna(after)) or (str(before) == str(after)):
                            continue
                        if "일" in norm_key(col):
                            after = normalize_date(after)
                        updates.append((excel_row, col, after))
                if not updates:
                    st.info("변경된 값이 없습니다.")
                else:
                    try:
                        update_sheet_cells(xlsx_path, SHEET_SINGLE, updates)
                        st.success("저장 완료!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"저장 실패: {e}")

    elif mode == "바인더 Lot":
        c_bl = find_col(b_df, "Lot(자동)")
        c_bn = find_col(b_df, "바인더명")
        c_bnote = find_col(b_df, "비고")
        hit_b = text_filter(b_df, [c_bl, c_bn, c_bnote], q)
        st.subheader("바인더_제조_입고")
        hit_b_show = add_excel_row_number(hit_b.sort_values(by=c_b_date, ascending=False) if c_b_date else hit_b)
        st.dataframe(hit_b_show, use_container_width=True)

        if q and c_s_blot:
            hit_s = s_df[s_df[c_s_blot].astype(str).str.contains(str(q).strip(), case=False, na=False)]
            st.subheader("연결된 단일색_수입검사 (사용된 바인더 Lot)")
            hit_s_show = add_excel_row_number(hit_s.sort_values(by=c_s_date, ascending=False) if c_s_date else hit_s)
            st.dataframe(hit_s_show, use_container_width=True)

        if edit_mode:
            if len(hit_b_show) > 0:
                st.markdown("#### 🔧 바인더 결과 수정")
                edited_b = st.data_editor(hit_b_show, use_container_width=True, num_rows="fixed", disabled=["_excel_row"], key="qs_edit_binder")
                if st.button("변경사항 저장(바인더)", type="primary", key="qs_save_binder"):
                    updates = []
                    for i in range(len(hit_b_show)):
                        excel_row = int(hit_b_show.iloc[i]["_excel_row"])
                        for col in hit_b_show.columns:
                            if col == "_excel_row":
                                continue
                            before = hit_b_show.iloc[i][col]
                            after = edited_b.iloc[i][col]
                            if (pd.isna(before) and pd.isna(after)) or (str(before) == str(after)):
                                continue
                            if "일" in norm_key(col):
                                after = normalize_date(after)
                            updates.append((excel_row, col, after))
                    if not updates:
                        st.info("변경된 값이 없습니다.")
                    else:
                        try:
                            update_sheet_cells(xlsx_path, SHEET_BINDER, updates)
                            st.success("저장 완료!")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"저장 실패: {e}")

            # 연결 단일색도 수정 허용
            if q and c_s_blot and 'hit_s_show' in locals() and len(hit_s_show) > 0:
                st.markdown("#### 🔧 연결된 단일색 결과 수정")
                edited_s = st.data_editor(hit_s_show, use_container_width=True, num_rows="fixed", disabled=["_excel_row"], key="qs_edit_single_by_binder")
                if st.button("변경사항 저장(연결 단일색)", type="primary", key="qs_save_single_by_binder"):
                    updates = []
                    for i in range(len(hit_s_show)):
                        excel_row = int(hit_s_show.iloc[i]["_excel_row"])
                        for col in hit_s_show.columns:
                            if col == "_excel_row":
                                continue
                            before = hit_s_show.iloc[i][col]
                            after = edited_s.iloc[i][col]
                            if (pd.isna(before) and pd.isna(after)) or (str(before) == str(after)):
                                continue
                            if "일" in norm_key(col):
                                after = normalize_date(after)
                            updates.append((excel_row, col, after))
                    if not updates:
                        st.info("변경된 값이 없습니다.")
                    else:
                        try:
                            update_sheet_cells(xlsx_path, SHEET_SINGLE, updates)
                            st.success("저장 완료!")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"저장 실패: {e}")

    elif mode == "단일색 잉크 Lot":
        hit = text_filter(s_df, [c_s_lot, c_s_pc, c_s_blot, c_s_cg, find_col(s_df, "비고")], q)
        st.subheader("단일색_수입검사")
        hit_show = add_excel_row_number(hit.sort_values(by=c_s_date, ascending=False) if c_s_date else hit)
        st.dataframe(hit_show, use_container_width=True)

        # 연결 바인더
        if len(hit) == 1 and c_s_blot:
            b_lot = str(hit.iloc[0].get(c_s_blot, "")).strip()
            if b_lot:
                c_bl = find_col(b_df, "Lot(자동)")
                hit_b = b_df[b_df[c_bl].astype(str) == b_lot] if c_bl else b_df.iloc[0:0]
                if len(hit_b):
                    st.subheader("연결된 바인더_제조_입고")
                    st.dataframe(add_excel_row_number(hit_b), use_container_width=True)

        if edit_mode and len(hit_show) > 0:
            st.markdown("#### 🔧 검색 결과 수정")
            edited = st.data_editor(hit_show, use_container_width=True, num_rows="fixed", disabled=["_excel_row"], key="qs_edit_single_lot")
            if st.button("변경사항 저장(단일색 Lot 검색)", type="primary", key="qs_save_single_lot"):
                updates = []
                for i in range(len(hit_show)):
                    excel_row = int(hit_show.iloc[i]["_excel_row"])
                    for col in hit_show.columns:
                        if col == "_excel_row":
                            continue
                        before = hit_show.iloc[i][col]
                        after = edited.iloc[i][col]
                        if (pd.isna(before) and pd.isna(after)) or (str(before) == str(after)):
                            continue
                        if "일" in norm_key(col):
                            after = normalize_date(after)
                        updates.append((excel_row, col, after))
                if not updates:
                    st.info("변경된 값이 없습니다.")
                else:
                    try:
                        update_sheet_cells(xlsx_path, SHEET_SINGLE, updates)
                        st.success("저장 완료!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"저장 실패: {e}")

    elif mode == "제품코드":
        hit = text_filter(s_df, [c_s_pc], q)
        st.subheader("단일색_수입검사")
        hit_show = add_excel_row_number(hit.sort_values(by=c_s_date, ascending=False) if c_s_date else hit)
        st.dataframe(hit_show, use_container_width=True)

        if edit_mode and len(hit_show) > 0:
            st.markdown("#### 🔧 검색 결과 수정")
            edited = st.data_editor(hit_show, use_container_width=True, num_rows="fixed", disabled=["_excel_row"], key="qs_edit_pc")
            if st.button("변경사항 저장(제품코드 검색)", type="primary", key="qs_save_pc"):
                updates = []
                for i in range(len(hit_show)):
                    excel_row = int(hit_show.iloc[i]["_excel_row"])
                    for col in hit_show.columns:
                        if col == "_excel_row":
                            continue
                        before = hit_show.iloc[i][col]
                        after = edited.iloc[i][col]
                        if (pd.isna(before) and pd.isna(after)) or (str(before) == str(after)):
                            continue
                        if "일" in norm_key(col):
                            after = normalize_date(after)
                        updates.append((excel_row, col, after))
                if not updates:
                    st.info("변경된 값이 없습니다.")
                else:
                    try:
                        update_sheet_cells(xlsx_path, SHEET_SINGLE, updates)
                        st.success("저장 완료!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"저장 실패: {e}")

    elif mode == "색상군":
        hit = text_filter(s_df, [c_s_cg], q)
        st.subheader("단일색_수입검사")
        hit_show = add_excel_row_number(hit.sort_values(by=c_s_date, ascending=False) if c_s_date else hit)
        st.dataframe(hit_show, use_container_width=True)

        if edit_mode and len(hit_show) > 0:
            st.markdown("#### 🔧 검색 결과 수정")
            edited = st.data_editor(hit_show, use_container_width=True, num_rows="fixed", disabled=["_excel_row"], key="qs_edit_cg")
            if st.button("변경사항 저장(색상군 검색)", type="primary", key="qs_save_cg"):
                updates = []
                for i in range(len(hit_show)):
                    excel_row = int(hit_show.iloc[i]["_excel_row"])
                    for col in hit_show.columns:
                        if col == "_excel_row":
                            continue
                        before = hit_show.iloc[i][col]
                        after = edited.iloc[i][col]
                        if (pd.isna(before) and pd.isna(after)) or (str(before) == str(after)):
                            continue
                        if "일" in norm_key(col):
                            after = normalize_date(after)
                        updates.append((excel_row, col, after))
                if not updates:
                    st.info("변경된 값이 없습니다.")
                else:
                    try:
                        update_sheet_cells(xlsx_path, SHEET_SINGLE, updates)
                        st.success("저장 완료!")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"저장 실패: {e}")

