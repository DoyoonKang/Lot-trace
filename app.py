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
# Page Config
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
SHEET_BASE_LAB = "기준LAB"

# 업체반환(kg 단위 기록용) - 없으면 자동 생성
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
# Utils
# =========================
def norm_key(x) -> str:
    if x is None:
        return ""
    s = str(x).replace("\n", " ").replace("\r", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def safe_to_float(x):
    if x is None:
        return None
    if isinstance(x, float) and pd.isna(x):
        return None
    if isinstance(x, str):
        if x.strip() == "":
            return None
        x = x.replace(",", "")
    try:
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

def delta_e76(lab1, lab2):
    return float(((lab1[0]-lab2[0])**2 + (lab1[1]-lab2[1])**2 + (lab1[2]-lab2[2])**2) ** 0.5)

def drop_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    """pandas read_excel 시 생기는 Unnamed 컬럼 제거 + 전부 NaN 컬럼 제거"""
    df = df.copy()
    df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed:")]]
    df = df.dropna(axis=1, how="all")
    return df

def get_col(df: pd.DataFrame, wanted: str):
    """줄바꿈/공백 차이가 있어도 컬럼을 찾아줌"""
    w = norm_key(wanted)
    for c in df.columns:
        if norm_key(c) == w:
            return c
    return None

def safe_date_bounds(series) -> tuple[dt.date, dt.date]:
    s = pd.to_datetime(series, errors="coerce").dropna()
    if len(s) == 0:
        today = dt.date.today()
        return today, today
    return s.min().date(), s.max().date()

def ensure_sheet_exists(xlsx_path: str, sheet_name: str, headers: list[str]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        wb.save(xlsx_path)

def get_sheet_header_map(xlsx_path: str, sheet_name: str):
    """
    엑셀 1행 헤더 기준으로
    - headers: 마지막 유효 헤더까지의 리스트
    - idx: header -> column index(0-based)
    """
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    raw_headers = [c.value for c in ws[1]]
    # 마지막 유효 헤더까지 자르기(뒤쪽 None들 제거)
    last = -1
    for i, h in enumerate(raw_headers):
        if h is not None and str(h).strip() != "":
            last = i
    if last < 0:
        raise ValueError(f"{sheet_name} 시트의 1행 헤더가 비어 있습니다.")
    headers = raw_headers[: last + 1]
    idx = {h: i for i, h in enumerate(headers) if h is not None and str(h).strip() != ""}
    return headers, idx

def append_row_by_headers(xlsx_path: str, sheet_name: str, row: dict):
    """
    ✅ 핵심: '엑셀 시트 헤더' 기준으로 컬럼 위치 고정 append.
    (컬럼명 줄바꿈/공백/Unnamed 컬럼 때문에 값이 밀려 들어가는 문제 방지)
    """
    headers, idx = get_sheet_header_map(xlsx_path, sheet_name)
    values = [None] * len(headers)
    for h, i in idx.items():
        # row에서 동일 헤더 우선, 없으면 norm_key 비교로 보조 매칭
        v = row.get(h, None)
        if v is None:
            # 보조 매칭
            nh = norm_key(h)
            for k in row.keys():
                if norm_key(k) == nh:
                    v = row.get(k)
                    break
        values[i] = v
    wb = load_workbook(xlsx_path)
    ws = wb[sheet_name]
    ws.append(values)
    wb.save(xlsx_path)

def append_rows_by_headers(xlsx_path: str, sheet_name: str, rows: list[dict]):
    headers, idx = get_sheet_header_map(xlsx_path, sheet_name)
    wb = load_workbook(xlsx_path)
    ws = wb[sheet_name]
    for row in rows:
        values = [None] * len(headers)
        for h, i in idx.items():
            v = row.get(h, None)
            if v is None:
                nh = norm_key(h)
                for k in row.keys():
                    if norm_key(k) == nh:
                        v = row.get(k)
                        break
            values[i] = v
        ws.append(values)
    wb.save(xlsx_path)

def detect_date_col(df: pd.DataFrame):
    for c in df.columns:
        ck = norm_key(c)
        if any(k in ck for k in ["일자", "날짜", "date", "입고일", "출고일"]):
            return c
    return None

def get_binder_limits(spec_binder: pd.DataFrame, binder_name: str):
    c_name = get_col(spec_binder, "바인더명")
    c_item = get_col(spec_binder, "시험항목")
    c_lo = get_col(spec_binder, "하한")
    c_hi = get_col(spec_binder, "상한")
    c_rule = get_col(spec_binder, "Lot부여규칙")

    if not all([c_name, c_item, c_lo, c_hi]):
        return None, None, None, None

    df = spec_binder[spec_binder[c_name].astype(str).str.strip() == str(binder_name).strip()].copy()
    visc = df[df[c_item].astype(str).str.contains("점도", na=False)]
    uv = df[df[c_item].astype(str).str.contains("UV", na=False)]

    visc_lo = safe_to_float(visc[c_lo].dropna().iloc[0]) if len(visc[c_lo].dropna()) else None
    visc_hi = safe_to_float(visc[c_hi].dropna().iloc[0]) if len(visc[c_hi].dropna()) else None
    uv_hi = safe_to_float(uv[c_hi].dropna().iloc[0]) if len(uv[c_hi].dropna()) else None
    rule = df[c_rule].dropna().iloc[0] if c_rule and (c_rule in df.columns) and len(df[c_rule].dropna()) else None
    return visc_lo, visc_hi, uv_hi, rule

def infer_binder_type_from_lot(spec_binder: pd.DataFrame, binder_lot: str):
    if not binder_lot:
        return None
    c_name = get_col(spec_binder, "바인더명")
    c_rule = get_col(spec_binder, "Lot부여규칙")
    if not c_name or not c_rule:
        return None
    rules = spec_binder[[c_name, c_rule]].dropna().drop_duplicates().to_dict("records")
    for r in rules:
        rule = str(r[c_rule])
        m = re.match(r"^([A-Za-z0-9]+)\+", rule)
        if m:
            prefix = m.group(1)
            if str(binder_lot).strip().startswith(prefix):
                return str(r[c_name]).strip()
    return None

def next_seq_for_pattern(existing_lots: pd.Series, prefix: str, date_str: str, sep: str = "-"):
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
        code = re.sub(r"\W+", "", str(binder_name))[:6].upper()
        return f"{code}{mfg_date.strftime('%Y%m%d')}-01"

    m = re.match(r"^([A-Za-z0-9]+)\+YYYYMMDD(-##)?$", str(rule).strip())
    if not m:
        code = re.sub(r"\W+", "", str(binder_name))[:6].upper()
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

    c_lot = get_col(single_df, "단일색잉크 Lot")
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

def extract_or_compute_de76(single_df: pd.DataFrame, base_lab: pd.DataFrame) -> pd.Series:
    out = pd.Series([None] * len(single_df), index=single_df.index, dtype="float")

    c_note = get_col(single_df, "비고")
    if c_note:
        pat = re.compile(r"\[\s*ΔE76\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*\]")
        for idx, val in single_df[c_note].items():
            if pd.isna(val):
                continue
            m = pat.search(str(val))
            if m:
                try:
                    out.loc[idx] = float(m.group(1))
                except Exception:
                    pass

    # Lab 기반 계산(가능할 때만)
    c_pc = get_col(single_df, "제품코드")
    c_L = get_col(single_df, "착색력_L*")
    c_a = get_col(single_df, "착색력_a*")
    c_b = get_col(single_df, "착색력_b*")

    b_pc = get_col(base_lab, "제품코드")
    b_L = get_col(base_lab, "기준_L*")
    b_a = get_col(base_lab, "기준_a*")
    b_b = get_col(base_lab, "기준_b*")

    if all([c_pc, c_L, c_a, c_b, b_pc, b_L, b_a, b_b]):
        base = base_lab.copy()
        base[b_pc] = base[b_pc].astype(str).str.strip()
        base_map = base.set_index(b_pc)[[b_L, b_a, b_b]].to_dict("index")

        for idx, row in single_df.iterrows():
            if pd.notna(out.loc[idx]):
                continue
            pc = row.get(c_pc, None)
            if pd.isna(pc):
                continue
            pc = str(pc).strip()
            if pc not in base_map:
                continue
            L = safe_to_float(row.get(c_L, None))
            a = safe_to_float(row.get(c_a, None))
            b = safe_to_float(row.get(c_b, None))
            if None in (L, a, b):
                continue
            ref = base_map[pc]
            ref_lab = (safe_to_float(ref[b_L]), safe_to_float(ref[b_a]), safe_to_float(ref[b_b]))
            if None in ref_lab:
                continue
            out.loc[idx] = delta_e76((L, a, b), ref_lab)

    return out

def get_single_spec(spec_single: pd.DataFrame, color_group: str, product_code: str, binder_type: str | None):
    c_cg = get_col(spec_single, "색상군")
    c_pc = get_col(spec_single, "제품코드")
    c_lo = get_col(spec_single, "하한")
    c_hi = get_col(spec_single, "상한")
    c_bt = get_col(spec_single, "BinderType")

    if not all([c_cg, c_pc, c_lo, c_hi]):
        return None, None, 0

    hit = spec_single[
        (spec_single[c_cg].astype(str).str.strip() == str(color_group).strip())
        & (spec_single[c_pc].astype(str).str.strip() == str(product_code).strip())
    ].copy()

    if binder_type and c_bt and (c_bt in hit.columns):
        hit = hit[hit[c_bt].astype(str).str.strip() == str(binder_type).strip()]

    if len(hit) == 0:
        return None, None, 0

    lo = safe_to_float(hit[c_lo].iloc[0])
    hi = safe_to_float(hit[c_hi].iloc[0])
    return lo, hi, len(hit)

def update_spec_single_limits(xlsx_path: str, color_group: str, product_code: str, binder_type: str | None, new_lo, new_hi):
    """Spec_Single_H&S 시트에서 조건에 맞는 행(들)의 하한/상한을 업데이트"""
    wb = load_workbook(xlsx_path)
    if SHEET_SPEC_SINGLE not in wb.sheetnames:
        return 0, "Spec_Single_H&S 시트가 없습니다."
    ws = wb[SHEET_SPEC_SINGLE]

    headers = [c.value for c in ws[1]]
    # 필요한 컬럼 index 찾기
    def find_idx(name):
        for i, h in enumerate(headers):
            if norm_key(h) == norm_key(name):
                return i + 1  # openpyxl는 1-based
        return None

    i_cg = find_idx("색상군")
    i_pc = find_idx("제품코드")
    i_lo = find_idx("하한")
    i_hi = find_idx("상한")
    i_bt = find_idx("BinderType")  # 있을 수도/없을 수도

    if not all([i_cg, i_pc, i_lo, i_hi]):
        return 0, "Spec_Single_H&S 헤더(색상군/제품코드/하한/상한)를 찾지 못했습니다."

    updated = 0
    for r in range(2, ws.max_row + 1):
        v_cg = ws.cell(r, i_cg).value
        v_pc = ws.cell(r, i_pc).value
        if norm_key(v_cg) != norm_key(color_group):
            continue
        if norm_key(v_pc) != norm_key(product_code):
            continue

        if binder_type and i_bt:
            v_bt = ws.cell(r, i_bt).value
            if norm_key(v_bt) != norm_key(binder_type):
                continue

        # 업데이트
        ws.cell(r, i_lo).value = float(new_lo) if new_lo is not None else None
        ws.cell(r, i_hi).value = float(new_hi) if new_hi is not None else None
        updated += 1

    wb.save(xlsx_path)
    return updated, None


# =========================
# Header
# =========================
st.title("액상 잉크 Lot 추적 관리 대시보드")
st.caption("✅ 빠른 검색 | ✅ 잉크 입고(엑셀 누적) | ✅ 대시보드(목록/평균/추이) | ✅ 바인더 입출고(구글시트 자동 반영)")


# =========================
# Sidebar: Excel file
# =========================
with st.sidebar:
    st.header("데이터 파일")
    xlsx_path = st.text_input("엑셀 파일 경로", value=DEFAULT_XLSX)
    uploaded = st.file_uploader("또는 엑셀 업로드(업로드 모드: 서버 저장 보장 X)", type=["xlsx"])

# 업로드 파일은 "처음 업로드 시"만 tmp로 복사 (rerun마다 원본으로 덮어쓰기 방지)
if uploaded is not None:
    upload_sig = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.get("_uploaded_sig") != upload_sig:
        tmp_path = Path(".streamlit_tmp.xlsx")
        tmp_path.write_bytes(uploaded.getvalue())
        st.session_state["_uploaded_sig"] = upload_sig
        st.session_state["_tmp_xlsx_path"] = str(tmp_path)
    xlsx_path = st.session_state.get("_tmp_xlsx_path", xlsx_path)
    st.sidebar.info("업로드 파일로 실행 중입니다. (이 모드에서는 서버 재시작 시 누적 보장이 어렵습니다.)")

if not Path(xlsx_path).exists():
    st.error(f"엑셀 파일을 찾을 수 없습니다: {xlsx_path}")
    st.stop()

# 업체반환 시트 없으면 생성
ensure_sheet_exists(
    xlsx_path,
    SHEET_BINDER_RETURN,
    headers=["일자", "바인더타입", "바인더명", "바인더 Lot", "반환량(kg)", "비고"],
)

# =========================
# Load Excel -> pandas
# =========================
@st.cache_data(show_spinner=False)
def load_excel_all(xlsx_path: str):
    binder = drop_unnamed(pd.read_excel(xlsx_path, sheet_name=SHEET_BINDER))
    single = drop_unnamed(pd.read_excel(xlsx_path, sheet_name=SHEET_SINGLE))
    spec_binder = drop_unnamed(pd.read_excel(xlsx_path, sheet_name=SHEET_SPEC_BINDER))
    spec_single = drop_unnamed(pd.read_excel(xlsx_path, sheet_name=SHEET_SPEC_SINGLE))
    base_lab = drop_unnamed(pd.read_excel(xlsx_path, sheet_name=SHEET_BASE_LAB))
    # 반환 시트는 없어도 자동 생성했으니 읽기
    binder_return = drop_unnamed(pd.read_excel(xlsx_path, sheet_name=SHEET_BINDER_RETURN))
    return binder, single, spec_binder, spec_single, base_lab, binder_return

binder_df, single_df, spec_binder, spec_single, base_lab, binder_return_df = load_excel_all(xlsx_path)

# 날짜 정규화
c_b_date = get_col(binder_df, "제조/입고일")
if c_b_date:
    binder_df[c_b_date] = binder_df[c_b_date].apply(normalize_date)

c_s_date = get_col(single_df, "입고일")
if c_s_date:
    single_df[c_s_date] = single_df[c_s_date].apply(normalize_date)

# ΔE76 파생
single_df["_ΔE76"] = extract_or_compute_de76(single_df, base_lab)

# rerun 키 안정화용(데이터가 갱신되면 자동으로 date_input key가 바뀌도록)
single_ver = str(pd.to_datetime(single_df[c_s_date], errors="coerce").max()) if c_s_date else "na"


# =========================
# Tabs
# =========================
tab_dash, tab_ink_in, tab_binder, tab_search = st.tabs(
    ["📊 대시보드", "✍️ 잉크 입고", "📦 바인더 입출고", "🔎 빠른검색"]
)

# =========================
# DASHBOARD
# =========================
with tab_dash:
    # KPI
    b_total = len(binder_df)
    s_total = len(single_df)

    c_b_judge = get_col(binder_df, "판정")
    c_s_vjudge = get_col(single_df, "점도판정")
    b_ng = int((binder_df[c_b_judge] == "부적합").sum()) if c_b_judge else 0
    s_ng = int((single_df[c_s_vjudge] == "부적합").sum()) if c_s_vjudge else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("바인더 기록", f"{b_total:,}")
    c2.metric("바인더 부적합", f"{b_ng:,}")
    c3.metric("단일색 기록", f"{s_total:,}")
    c4.metric("단일색(점도) 부적합", f"{s_ng:,}")

    st.divider()

    # ---- Spec 관리(상사가 봐도 이해되게, 상단 고정)
    with st.expander("🛠️ 단일색 점도 스펙(하한/상한) 관리 (대시보드에서 바로 수정)", expanded=False):
        c_cg = get_col(spec_single, "색상군")
        c_pc = get_col(spec_single, "제품코드")
        c_bt = get_col(spec_single, "BinderType")
        cg_opts = sorted(spec_single[c_cg].dropna().astype(str).unique().tolist()) if c_cg else []
        pc_opts = sorted(spec_single[c_pc].dropna().astype(str).unique().tolist()) if c_pc else []
        bt_opts = sorted(spec_single[c_bt].dropna().astype(str).unique().tolist()) if c_bt else []

        colA, colB, colC, colD, colE = st.columns([1.4, 1.4, 1.2, 1.0, 1.0])
        with colA:
            sel_cg = st.selectbox("색상군", cg_opts, index=0 if cg_opts else None)
        with colB:
            sel_pc = st.selectbox("제품코드", pc_opts, index=0 if pc_opts else None)
        with colC:
            sel_bt = st.selectbox("BinderType(있을 때만)", ["(미사용)"] + bt_opts, index=0)
            sel_bt = None if sel_bt == "(미사용)" else sel_bt

        cur_lo, cur_hi, hit_n = get_single_spec(spec_single, sel_cg, sel_pc, sel_bt)
        with colD:
            new_lo = st.number_input("하한(cP)", value=float(cur_lo) if cur_lo is not None else 0.0, step=10.0)
        with colE:
            new_hi = st.number_input("상한(cP)", value=float(cur_hi) if cur_hi is not None else 0.0, step=10.0)

        st.caption(f"현재 매칭 행 수: {hit_n} (여러 행이면 전부 동일 값으로 업데이트됩니다)")
        if st.button("스펙 저장(엑셀 반영)", type="primary"):
            if new_hi is not None and new_lo is not None and float(new_lo) > float(new_hi):
                st.error("하한이 상한보다 큽니다. 값을 확인해주세요.")
            else:
                updated, err = update_spec_single_limits(xlsx_path, sel_cg, sel_pc, sel_bt, new_lo, new_hi)
                if err:
                    st.error(err)
                else:
                    st.success(f"스펙 저장 완료! 업데이트 행 수: {updated}")
                    st.cache_data.clear()
                    st.rerun()

    st.divider()

    # ---- 1) 엑셀형 리스트(요청)
    st.subheader("1) 단일색 데이터 목록 (엑셀형 보기)")

    c_s_cg = get_col(single_df, "색상군")
    c_s_pc = get_col(single_df, "제품코드")
    c_s_lot = get_col(single_df, "단일색잉크 Lot")
    c_s_blot = get_col(single_df, "사용된 바인더 Lot")
    c_s_visc = get_col(single_df, "점도측정값(cP)")

    needed = [c_s_date, c_s_cg, c_s_pc, c_s_blot, c_s_visc]
    if any(x is None for x in needed):
        st.warning("단일색 시트에서 필요한 컬럼(입고일/색상군/제품코드/사용된 바인더 Lot/점도측정값)을 찾지 못했습니다.")
    else:
        df_list = single_df.copy()
        df_list[c_s_date] = pd.to_datetime(df_list[c_s_date], errors="coerce")
        dmin, dmax = safe_date_bounds(df_list[c_s_date])

        f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
        with f1:
            start = st.date_input("시작일(목록)", value=max(dmin, dmax - dt.timedelta(days=90)), key=f"list_start_{single_ver}")
        with f2:
            end = st.date_input("종료일(목록)", value=dmax, key=f"list_end_{single_ver}")
        with f3:
            cg_opts = sorted(df_list[c_s_cg].dropna().astype(str).unique().tolist())
            cg = st.multiselect("색상군(목록)", cg_opts, key=f"list_cg_{single_ver}")
        with f4:
            pc_opts = sorted(df_list[c_s_pc].dropna().astype(str).unique().tolist())
            pc = st.multiselect("제품코드(목록)", pc_opts, key=f"list_pc_{single_ver}")

        if start > end:
            start, end = end, start

        df_list = df_list[(df_list[c_s_date].dt.date >= start) & (df_list[c_s_date].dt.date <= end)]
        if cg:
            df_list = df_list[df_list[c_s_cg].astype(str).isin(cg)]
        if pc:
            df_list = df_list[df_list[c_s_pc].astype(str).isin(pc)]

        view = pd.DataFrame({
            "제조일자": df_list[c_s_date].dt.date,
            "색상군": df_list[c_s_cg],
            "제품코드": df_list[c_s_pc],
            "사용된바인더": df_list[c_s_blot],
            "단일색Lot": df_list[c_s_lot] if c_s_lot else None,
            "점도(cP)": pd.to_numeric(df_list[c_s_visc].astype(str).str.replace(",", "", regex=False), errors="coerce"),
            "색차(ΔE76)": df_list["_ΔE76"],
        }).sort_values(by="제조일자", ascending=False)

        st.dataframe(view, use_container_width=True, height=330)

        st.divider()

        # ---- 1-1) 평균점도 점+값
        st.subheader("1-1) 색상군별 평균 점도 (점 + 값 표시)")
        mean_df = (
            view.dropna(subset=["색상군", "점도(cP)"])
            .groupby("색상군", as_index=False)["점도(cP)"]
            .mean()
            .rename(columns={"점도(cP)": "평균점도(cP)"})
        )
        if len(mean_df) == 0:
            st.info("평균 점도 그래프를 만들 데이터가 없습니다.")
        else:
            mean_df["표시"] = mean_df["평균점도(cP)"].round(0).astype("Int64").astype(str)
            base = alt.Chart(mean_df).encode(
                x=alt.X("색상군:N", sort=sorted(mean_df["색상군"].astype(str).unique().tolist()), title="색상군"),
                y=alt.Y("평균점도(cP):Q", title="평균 점도(cP)"),
                tooltip=["색상군:N", "평균점도(cP):Q"]
            )
            points = base.mark_circle(size=260)
            labels = base.mark_text(dx=10, dy=-10).encode(text="표시:N")
            st.altair_chart((points + labels).interactive(), use_container_width=True)

    st.divider()

    # ---- 2) 추이 (Lot별) + 스펙선
    st.subheader("2) 단일색 점도 변화 추이 (Lot별)")
    st.caption("선택한 Lot별로 입고일 기준 점도 변화를 확인합니다. (점 크게 + 라벨 표시 + 스펙선 빨간색)")

    if all([c_s_date, c_s_visc]) and c_s_lot:
        df = single_df.copy()
        df[c_s_date] = pd.to_datetime(df[c_s_date], errors="coerce")
        df["점도"] = pd.to_numeric(df[c_s_visc].astype(str).str.replace(",", "", regex=False), errors="coerce")
        df["Lot"] = df[c_s_lot].astype(str)
        df = df.dropna(subset=[c_s_date, "점도"])
        df = df[df["Lot"].str.strip().ne("") & df["Lot"].str.lower().ne("nan")]

        if len(df) == 0:
            st.info("입고일/점도/Lot 값이 비어있어 추이 그래프를 표시할 수 없습니다.")
        else:
            dmin, dmax = safe_date_bounds(df[c_s_date])

            f1, f2, f3, f4, f5 = st.columns([1.2, 1.2, 1.6, 2.0, 1.0])
            with f1:
                start = st.date_input("시작일(추이)", value=max(dmin, dmax - dt.timedelta(days=90)), key=f"trend_start_{single_ver}")
            with f2:
                end = st.date_input("종료일(추이)", value=dmax, key=f"trend_end_{single_ver}")
            with f3:
                cg_opts = sorted(df[c_s_cg].dropna().astype(str).unique().tolist()) if c_s_cg else []
                cg = st.multiselect("색상군(추이)", cg_opts, key=f"trend_cg_{single_ver}")
            with f4:
                pc_opts = sorted(df[c_s_pc].dropna().astype(str).unique().tolist()) if c_s_pc else []
                pc = st.multiselect("제품코드(추이)", pc_opts, key=f"trend_pc_{single_ver}")
            with f5:
                show_labels = st.checkbox("라벨 표시", value=True, key=f"trend_labels_{single_ver}")

            if start > end:
                start, end = end, start

            df = df[(df[c_s_date].dt.date >= start) & (df[c_s_date].dt.date <= end)]
            if cg and c_s_cg:
                df = df[df[c_s_cg].astype(str).isin(cg)]
            if pc and c_s_pc:
                df = df[df[c_s_pc].astype(str).isin(pc)]

            lot_list = sorted(df["Lot"].dropna().unique().tolist())
            default_pick = lot_list[-5:] if len(lot_list) > 5 else lot_list
            pick = st.multiselect("표시할 단일색 Lot(복수 선택)", lot_list, default=default_pick, key=f"trend_lots_{single_ver}")
            if pick:
                df = df[df["Lot"].isin(pick)]

            if len(df) == 0:
                st.info("선택한 조건에 해당하는 데이터가 없습니다. (기간/색상군/제품코드/로트 필터 확인)")
            else:
                df = df.sort_values(c_s_date)
                df["점도표시"] = df["점도"].round(0).astype("Int64").astype(str)

                tooltip_cols = [f"{c_s_date}:T", "Lot:N", "점도:Q"]
                if c_s_pc:
                    tooltip_cols.insert(2, f"{c_s_pc}:N")
                if c_s_cg:
                    tooltip_cols.insert(3, f"{c_s_cg}:N")
                if c_s_blot:
                    tooltip_cols.append(f"{c_s_blot}:N")

                base = alt.Chart(df).encode(
                    x=alt.X(f"{c_s_date}:T", title="입고일"),
                    y=alt.Y("점도:Q", title="점도(cP)"),
                    tooltip=tooltip_cols
                )

                line = base.mark_line()
                points = base.mark_point(size=260).encode(color=alt.Color("Lot:N", title="Lot"))

                layers = [line, points]

                if show_labels:
                    labels = base.mark_text(dy=-12).encode(
                        color=alt.Color("Lot:N", legend=None),
                        text="점도표시:N"
                    )
                    layers.append(labels)

                # ✅ 스펙선(빨간선): 선택 조건이 좁혀졌을 때만 정확히 그리기
                spec_lo = None
                spec_hi = None
                if c_s_pc and c_s_cg:
                    uniq_pc = df[c_s_pc].dropna().astype(str).unique().tolist()
                    uniq_cg = df[c_s_cg].dropna().astype(str).unique().tolist()
                    if len(uniq_pc) == 1 and len(uniq_cg) == 1:
                        # binder_type은 단일 값일 때만 적용(없으면 None)
                        c_bt_auto = get_col(df, "BinderType(자동)")
                        bt = None
                        if c_bt_auto:
                            uniq_bt = df[c_bt_auto].dropna().astype(str).unique().tolist()
                            bt = uniq_bt[0] if len(uniq_bt) == 1 else None

                        spec_lo, spec_hi, _ = get_single_spec(spec_single, uniq_cg[0], uniq_pc[0], bt)

                if spec_lo is not None:
                    rule_lo = alt.Chart(pd.DataFrame({"y": [spec_lo]})).mark_rule().encode(y="y:Q")
                    layers.append(rule_lo)
                if spec_hi is not None:
                    rule_hi = alt.Chart(pd.DataFrame({"y": [spec_hi]})).mark_rule().encode(y="y:Q")
                    layers.append(rule_hi)

                st.altair_chart(alt.layer(*layers).interactive(), use_container_width=True)

                if spec_lo is not None or spec_hi is not None:
                    st.caption(f"적용 스펙: 하한={spec_lo if spec_lo is not None else '-'} cP / 상한={spec_hi if spec_hi is not None else '-'} cP (빨간선)")

    else:
        st.warning("단일색 시트에서 추이 그래프에 필요한 컬럼(입고일/단일색잉크 Lot/점도측정값)을 찾지 못했습니다.")

    st.divider()
    st.subheader("최근 20건 (단일색)")
    show = single_df.copy()
    if c_s_date:
        show[c_s_date] = pd.to_datetime(show[c_s_date], errors="coerce")
        show = show.sort_values(by=c_s_date, ascending=False)
    st.dataframe(show.head(20), use_container_width=True)


# =========================
# 잉크 입고 (단일색 입력)
# =========================
with tab_ink_in:
    st.subheader("단일색 잉크 입력(입고)")
    st.info("이 탭은 **단일색_수입검사** 시트에 행을 추가(Append)하여 누적합니다. (동시 사용 시 충돌 가능)")

    # 옵션 목록
    ink_types = ["HEMA", "Silicone"]

    # spec_single에서 옵션 추출
    sp_cg = get_col(spec_single, "색상군")
    sp_pc = get_col(spec_single, "제품코드")
    color_groups = sorted(spec_single[sp_cg].dropna().astype(str).unique().tolist()) if sp_cg else []
    product_codes = sorted(spec_single[sp_pc].dropna().astype(str).unique().tolist()) if sp_pc else []

    # binder lot 목록
    b_lot_col = get_col(binder_df, "Lot(자동)")
    binder_lots = binder_df[b_lot_col].dropna().astype(str).tolist() if b_lot_col else []
    binder_lots = sorted(set([x.strip() for x in binder_lots if x.strip()]), reverse=True)

    with st.form("single_form", clear_on_submit=True):
        col1, col2, col3, col4 = st.columns([1.2, 1.3, 1.5, 2.0])
        with col1:
            in_date = st.date_input("입고일", value=dt.date.today(), key="single_in_date")
            ink_type = st.selectbox("잉크타입", ink_types, key="single_ink_type")
            color_group = st.selectbox("색상군", color_groups, key="single_cg")
        with col2:
            product_code = st.selectbox("제품코드", product_codes, key="single_pc")
            binder_lot = st.selectbox("사용된 바인더 Lot", binder_lots, key="single_blot")
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
        binder_type = infer_binder_type_from_lot(spec_binder, binder_lot)

        # spec 조회
        lo, hi, _ = get_single_spec(spec_single, color_group, product_code, binder_type)
        visc_judge = judge_range(visc_meas, lo, hi) if (lo is not None or hi is not None) else None

        new_lot = generate_single_lot(single_df, product_code, color_group, in_date)
        if new_lot is None:
            st.error("단일색 Lot 자동 생성에 실패했습니다. (색상군 매핑 확인 필요)")
        else:
            note2 = note
            if lab_enabled:
                b_pc = get_col(base_lab, "제품코드")
                b_L = get_col(base_lab, "기준_L*")
                b_a = get_col(base_lab, "기준_a*")
                b_b = get_col(base_lab, "기준_b*")
                if all([b_pc, b_L, b_a, b_b]):
                    base_hit = base_lab[base_lab[b_pc].astype(str).str.strip() == str(product_code).strip()]
                    if len(base_hit) == 1:
                        ref = (
                            safe_to_float(base_hit.iloc[0][b_L]),
                            safe_to_float(base_hit.iloc[0][b_a]),
                            safe_to_float(base_hit.iloc[0][b_b]),
                        )
                        if None not in ref:
                            de = delta_e76((float(L), float(a), float(b)), ref)
                            note2 = (note2 + " " if note2 else "") + f"[ΔE76={de:.2f}]"
                        else:
                            note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"
                    else:
                        note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"
                else:
                    note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"

            # ✅ 반드시 '엑셀 헤더명'으로 저장(컬럼 밀림 방지)
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
                append_row_by_headers(xlsx_path, SHEET_SINGLE, row)
                st.success(f"저장 완료! 단일색 Lot = {new_lot} / 점도판정 = {visc_judge}")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")


# =========================
# 바인더 입출고
# =========================
with tab_binder:
    # 1) 업체반환(탭 최상단, 재고요약 제거 요청 반영)
    st.subheader("업체반환(반품) 입력 (kg 단위)")
    st.caption("※ 20kg(1통) 기준이더라도, 실제 반환량은 kg 단위로 입력합니다.")

    # binder명 목록
    sb_name = get_col(spec_binder, "바인더명")
    binder_names = sorted(spec_binder[sb_name].dropna().astype(str).unique().tolist()) if sb_name else []

    # binder lot 목록
    b_lot_col = get_col(binder_df, "Lot(자동)")
    binder_lots = binder_df[b_lot_col].dropna().astype(str).tolist() if b_lot_col else []
    binder_lots = sorted(set([x.strip() for x in binder_lots if x.strip()]), reverse=True)

    with st.form("binder_return_form", clear_on_submit=True):
        c1, c2, c3 = st.columns([1.2, 1.2, 2.6])
        with c1:
            r_date = st.date_input("반환일자", value=dt.date.today(), key="ret_date")
        with c2:
            r_type = st.selectbox("바인더타입", ["HEMA", "Silicone"], key="ret_type")
        with c3:
            r_name = st.selectbox("바인더명", binder_names, key="ret_name")

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
                append_row_by_headers(xlsx_path, SHEET_BINDER_RETURN, row)
                st.success("반품 저장 완료!")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"반품 저장 실패: {e}")

    st.divider()

    # 2) 바인더 제조/입고 입력
    st.subheader("바인더 입력 (제조/입고) — 여러 Lot/날짜 묶음 입력 지원")
    st.caption("※ 바인더는 여러 날짜의 Lot가 한 번에 입고될 수 있어, 날짜별/수량별 묶음 입력을 지원합니다.")

    input_mode = st.radio("입력 방식", ["개별 입력", "묶음 입력(여러 날짜/수량)"], horizontal=True, key="binder_input_mode")

    if input_mode == "개별 입력":
        with st.form("binder_form_single", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                mfg_date = st.date_input("제조/입고일", value=dt.date.today(), key="b_single_date")
                b_name = st.selectbox("바인더명", binder_names, key="b_single_name")
            with col2:
                visc = st.number_input("점도(cP)", min_value=0.0, step=1.0, format="%.1f", key="b_single_visc")
                uv = st.number_input("UV흡광도(선택)", min_value=0.0, step=0.01, format="%.3f", key="b_single_uv")
                uv_enabled = st.checkbox("UV 값 입력함", value=False, key="b_single_uv_en")
            with col3:
                note = st.text_input("비고", value="", key="b_single_note")
                submit_b = st.form_submit_button("저장(바인더)")

        if submit_b:
            visc_lo, visc_hi, uv_hi, _ = get_binder_limits(spec_binder, b_name)
            lot = generate_binder_lot(spec_binder, b_name, mfg_date, binder_df.get(get_col(binder_df, "Lot(자동)"), pd.Series(dtype=str)))

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
                append_row_by_headers(xlsx_path, SHEET_BINDER, row)
                st.success(f"저장 완료! 바인더 Lot = {lot}")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    else:
        st.caption("아래 표에 날짜/바인더명/수량(통)/점도/UV/비고를 입력하고, 한 번에 저장하세요.")

        # 기본 3행 제공
        base_rows = [
            {"제조/입고일": dt.date.today(), "바인더명": (binder_names[0] if binder_names else ""), "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
            {"제조/입고일": dt.date.today() - dt.timedelta(days=1), "바인더명": (binder_names[0] if binder_names else ""), "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
            {"제조/입고일": dt.date.today() - dt.timedelta(days=2), "바인더명": (binder_names[0] if binder_names else ""), "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
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

            existing_lot_col = get_col(binder_df, "Lot(자동)")
            existing = binder_df[existing_lot_col] if existing_lot_col else pd.Series(dtype=str)
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
                    st.error(f"[{b_name}] Lot부여규칙 해석 실패 (Spec_Binder 확인 필요)")
                    st.stop()

                prefix = m.group(1)
                has_seq = bool(m.group(2))
                date_str = mfg_date.strftime("%Y%m%d")

                if (not has_seq) and qty > 1:
                    st.error(f"[{b_name}] 규칙에 순번(-##)이 없어 여러 통(수량={qty}) 자동 생성 불가")
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
                        "점도(cP)": float(visc),
                        "UV흡광도(선택)": float(uv_val) if uv_enabled and uv_val is not None else None,
                        "판정": judge,
                        "비고": note,
                    })
                    existing_list.append(lot)

            st.write("저장 미리보기(상위 30건)")
            st.dataframe(pd.DataFrame(rows_out).tail(30), use_container_width=True)

            try:
                append_rows_by_headers(xlsx_path, SHEET_BINDER, rows_out)
                st.success(f"묶음 저장 완료! 총 {len(rows_out)}건 입력했습니다.")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    st.divider()

    # 3) Google Sheet 보기(최신순)
    st.subheader("바인더 입출고 (Google Sheets 자동 반영, 최신순)")
    st.caption("구글 시트를 수정하면 새로고침 시 자동 반영됩니다. (캐시 60초)")

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


# =========================
# Search (원하시면 여기 확장)
# =========================
with tab_search:
    st.info("빠른검색은 원하시는 조건(바인더 Lot → 연결 단일색, 기간+색상군+제품코드 복합 등)으로 확장 가능합니다. 말씀만 주세요.")
