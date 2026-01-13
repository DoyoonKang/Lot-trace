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
# Page Config (딱 1번만)
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

# 새로 추가(없으면 자동 생성)
SHEET_BINDER_RETURN = "바인더_업체반환"  # kg 단위 반환 기록

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
def norm_key(x) -> str:
    if x is None:
        return ""
    s = str(x)
    s = s.replace("\n", " ").replace("\r", " ").strip()
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
    if x is None:
        return None
    if isinstance(x, float) and pd.isna(x):
        return None
    if isinstance(x, (dt.date, dt.datetime)):
        return x.date() if isinstance(x, dt.datetime) else x
    try:
        return pd.to_datetime(x, errors="coerce").date()
    except Exception:
        return None

def safe_date_bounds(s: pd.Series):
    s2 = pd.to_datetime(s, errors="coerce").dropna()
    if len(s2) == 0:
        today = dt.date.today()
        return today, today
    return s2.min().date(), s2.max().date()

def delta_e76(lab1, lab2):
    return float(((lab1[0]-lab2[0])**2 + (lab1[1]-lab2[1])**2 + (lab1[2]-lab2[2])**2) ** 0.5)

def judge_range(value, lo, hi):
    v = safe_to_float(value)
    if v is None:
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
        try:
            wb.calculation.calcMode = "auto"
            wb.calculation.fullCalcOnLoad = True
        except Exception:
            pass
        wb.save(xlsx_path)

def _read_excel_from_path(xlsx_path: str) -> dict[str, pd.DataFrame]:
    def read(name: str) -> pd.DataFrame:
        return pd.read_excel(xlsx_path, sheet_name=name)
    return {
        "binder": read(SHEET_BINDER),
        "single": read(SHEET_SINGLE),
        "spec_binder": read(SHEET_SPEC_BINDER),
        "spec_single": read(SHEET_SPEC_SINGLE),
        "base_lab": read(SHEET_BASE_LAB),
    }

@st.cache_data(show_spinner=False)
def load_data(xlsx_path: str) -> dict[str, pd.DataFrame]:
    return _read_excel_from_path(xlsx_path)

def set_excel_recalc_on_open(wb):
    # openpyxl로 저장하면 수식 캐시값이 날아가 Streamlit(pandas)이 None으로 읽는 경우가 많아서,
    # Excel에서 파일 열 때 자동 재계산되도록 설정
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

def append_row_to_sheet(xlsx_path: str, sheet_name: str, row_by_normkey: dict):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]

    values = []
    for h in headers:
        nk = norm_key(h)
        values.append(row_by_normkey.get(nk, None))
    ws.append(values)

    set_excel_recalc_on_open(wb)
    wb.save(xlsx_path)

def append_rows_to_sheet(xlsx_path: str, sheet_name: str, rows_by_normkey: list[dict]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    norm_headers = [norm_key(h) for h in headers]

    for row in rows_by_normkey:
        ws.append([row.get(nh, None) for nh in norm_headers])

    set_excel_recalc_on_open(wb)
    wb.save(xlsx_path)

def detect_date_col(df: pd.DataFrame):
    # 구글시트 컬럼명 다양성 대응
    for c in df.columns:
        ck = norm_key(c)
        if any(k in ck for k in ["일자", "날짜", "date", "입고일", "출고일"]):
            return c
    return None

# ===== Spec Helpers =====
def get_binder_limits(spec_binder: pd.DataFrame, binder_name: str):
    df = spec_binder[spec_binder["바인더명"] == binder_name].copy()
    visc = df[df["시험항목"].astype(str).str.contains("점도", na=False)]
    uv = df[df["시험항목"].astype(str).str.contains("UV", na=False)]

    visc_lo = safe_to_float(visc["하한"].dropna().iloc[0]) if len(visc["하한"].dropna()) else None
    visc_hi = safe_to_float(visc["상한"].dropna().iloc[0]) if len(visc["상한"].dropna()) else None
    uv_hi = safe_to_float(uv["상한"].dropna().iloc[0]) if len(uv["상한"].dropna()) else None
    rule = df["Lot부여규칙"].dropna().iloc[0] if "Lot부여규칙" in df.columns and len(df["Lot부여규칙"].dropna()) else None
    return visc_lo, visc_hi, uv_hi, rule

def infer_binder_type_from_lot(spec_binder: pd.DataFrame, binder_lot: str):
    # "Lot부여규칙"의 prefix로 바인더명(BinderType)을 역추론
    if not binder_lot or (isinstance(binder_lot, float) and pd.isna(binder_lot)):
        return None
    binder_lot = str(binder_lot).strip()
    rules = (
        spec_binder[["바인더명", "Lot부여규칙"]]
        .dropna()
        .drop_duplicates()
        .to_dict("records")
    )
    for r in rules:
        rule = str(r["Lot부여규칙"]).strip()
        m = re.match(r"^([A-Za-z0-9]+)\+", rule)
        if m:
            prefix = m.group(1)
            if binder_lot.startswith(prefix):
                return r["바인더명"]
    return None

def next_seq_for_pattern(existing_lots: list[str], prefix: str, date_str: str, sep: str = "-"):
    seqs = []
    for lot in existing_lots:
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

def generate_binder_lot(spec_binder: pd.DataFrame, binder_name: str, mfg_date: dt.date, existing_binder_lots: list[str]):
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

def lot_prefix_from_single(product_code: str, color_group: str, in_date: dt.date):
    pc = (product_code or "").strip()
    cc = COLOR_CODE.get(color_group)
    if not cc or not in_date:
        return None
    if pc.startswith("NPL"):
        prefix = "NPL"
    elif pc.startswith("PL"):
        prefix = "PL"
    elif pc.startswith("SL") or pc.startswith("NSL"):
        prefix = "SL"
    else:
        prefix = "PL"
    date_str = in_date.strftime("%y%m%d")
    return f"{prefix}{cc}{date_str}"

def fill_missing_single_lots(df_single: pd.DataFrame) -> pd.Series:
    """
    단일색잉크 Lot이 수식/캐시 문제로 None으로 읽히는 경우,
    앱에서 다시 lot을 복원(추이 그래프/검색 정상화)
    """
    s = df_single.copy()
    # 필요한 컬럼
    if not all(c in s.columns for c in ["입고일", "제품코드", "색상군", "단일색잉크 Lot"]):
        return s.get("단일색잉크 Lot", pd.Series([None]*len(s), index=s.index))

    s["_in_date"] = pd.to_datetime(s["입고일"], errors="coerce").dt.date
    s["_lot_raw"] = s["단일색잉크 Lot"].astype(str)
    s["_lot_raw"] = s["_lot_raw"].replace(["nan", "None", "NaT"], "").str.strip()

    # 기존 lot에서 prefix별 max seq 추출
    max_seq = {}
    patt = re.compile(r"^(NPL|PL|SL)([BWUGYRP])(\d{6})(\d{2,})$")
    for lot in s["_lot_raw"]:
        if not lot:
            continue
        m = patt.match(lot)
        if not m:
            continue
        pfx = f"{m.group(1)}{m.group(2)}{m.group(3)}"
        seq = int(m.group(4))
        max_seq[pfx] = max(max_seq.get(pfx, 0), seq)

    # 결측 lot 채우기(날짜→원본순)
    out = s["_lot_raw"].copy()
    for idx, row in s.sort_values(by=["_in_date"]).iterrows():
        if out.loc[idx]:
            continue
        in_date = row["_in_date"]
        pfx = lot_prefix_from_single(str(row.get("제품코드", "")).strip(), str(row.get("색상군", "")).strip(), in_date)
        if not pfx:
            continue
        next_seq = max_seq.get(pfx, 0) + 1
        max_seq[pfx] = next_seq
        out.loc[idx] = f"{pfx}{next_seq:02d}"
    return out

def compute_single_derived(df_single: pd.DataFrame, spec_binder: pd.DataFrame, spec_single: pd.DataFrame, base_lab: pd.DataFrame) -> pd.DataFrame:
    """
    Streamlit에서 None으로 읽히는(=엑셀 수식 캐시 문제) 컬럼들을 앱에서 안전하게 재계산:
    - BinderType(자동)
    - 점도하한/점도상한/점도판정
    - 단일색잉크 Lot (없으면 복원)
    - ΔE76(가능하면)
    """
    s = df_single.copy()

    # 날짜/점도 파싱
    if "입고일" in s.columns:
        s["_입고일_dt"] = pd.to_datetime(s["입고일"], errors="coerce")
    else:
        s["_입고일_dt"] = pd.NaT

    if "점도측정값(cP)" in s.columns:
        s["_점도"] = pd.to_numeric(s["점도측정값(cP)"].astype(str).str.replace(",", "", regex=False), errors="coerce")
    else:
        s["_점도"] = pd.NA

    # Lot 복원(없으면)
    if "단일색잉크 Lot" in s.columns:
        fixed_lot = fill_missing_single_lots(s)
        s["_Lot_fix"] = fixed_lot
    else:
        s["_Lot_fix"] = ""

    # BinderType(자동) 보정
    if "사용된 바인더 Lot" in s.columns:
        s["_BinderLot"] = s["사용된 바인더 Lot"].astype(str).replace(["nan", "None"], "").str.strip()
    else:
        s["_BinderLot"] = ""

    def _infer_bt(x):
        bt = infer_binder_type_from_lot(spec_binder, x)
        return bt

    s["_BinderType_fix"] = s.get("BinderType(자동)", pd.Series([None]*len(s), index=s.index))
    # 값이 None/NaN인 곳만 채우기
    mask_bt = s["_BinderType_fix"].isna() | (s["_BinderType_fix"].astype(str).str.strip().isin(["", "None", "nan"]))
    s.loc[mask_bt, "_BinderType_fix"] = s.loc[mask_bt, "_BinderLot"].apply(_infer_bt)

    # 점도 기준(하한/상한/판정) 보정
    # spec_single: 색상군, 제품코드, 하한, 상한, BinderType
    for c in ["색상군", "제품코드"]:
        if c in s.columns:
            s[c] = s[c].astype(str).str.strip()

    spec_single2 = spec_single.copy()
    for c in ["색상군", "제품코드", "BinderType"]:
        if c in spec_single2.columns:
            spec_single2[c] = spec_single2[c].astype(str).str.strip()

    def _lookup_limits(row):
        cg = row.get("색상군", "")
        pc = row.get("제품코드", "")
        bt = row.get("_BinderType_fix", None)
        if not cg or not pc:
            return None, None
        hit = spec_single2[(spec_single2["색상군"] == cg) & (spec_single2["제품코드"] == pc)].copy()
        if bt and "BinderType" in hit.columns and len(hit["BinderType"].dropna()):
            hit2 = hit[hit["BinderType"] == str(bt).strip()]
            if len(hit2) > 0:
                hit = hit2
        if len(hit) == 0:
            return None, None
        lo = safe_to_float(hit.iloc[0].get("하한", None))
        hi = safe_to_float(hit.iloc[0].get("상한", None))
        return lo, hi

    s["_점도하한_fix"] = s.get("점도하한", pd.Series([None]*len(s), index=s.index))
    s["_점도상한_fix"] = s.get("점도상한", pd.Series([None]*len(s), index=s.index))
    s["_점도판정_fix"] = s.get("점도판정", pd.Series([None]*len(s), index=s.index))

    for idx, row in s.iterrows():
        # 하한/상한/판정이 이미 숫자/값으로 있으면 존중
        lo0 = safe_to_float(row.get("_점도하한_fix", None))
        hi0 = safe_to_float(row.get("_점도상한_fix", None))
        judge0 = str(row.get("_점도판정_fix", "")).strip()
        need = (lo0 is None and hi0 is None) or (judge0 in ["", "None", "nan"])
        if not need:
            continue

        lo, hi = _lookup_limits(row)
        if lo0 is None:
            s.at[idx, "_점도하한_fix"] = lo
        if hi0 is None:
            s.at[idx, "_점도상한_fix"] = hi

        visc = row.get("_점도", None)
        s.at[idx, "_점도판정_fix"] = judge_range(visc, lo, hi)

    # ΔE76(가능하면)
    s["_ΔE76_fix"] = None
    if "제품코드" in base_lab.columns:
        base2 = base_lab.copy()
        base2["제품코드"] = base2["제품코드"].astype(str).str.strip()
        base_map = {}
        if all(c in base2.columns for c in ["기준_L*", "기준_a*", "기준_b*"]):
            for _, r in base2.iterrows():
                pc = str(r.get("제품코드", "")).strip()
                if not pc:
                    continue
                base_map[pc] = (safe_to_float(r.get("기준_L*", None)),
                                safe_to_float(r.get("기준_a*", None)),
                                safe_to_float(r.get("기준_b*", None)))

        if all(c in s.columns for c in ["착색력_L*", "착색력_a*", "착색력_b*"]):
            for idx, row in s.iterrows():
                pc = str(row.get("제품코드", "")).strip()
                if pc not in base_map:
                    continue
                ref = base_map[pc]
                if None in ref:
                    continue
                L = safe_to_float(row.get("착색력_L*", None))
                a = safe_to_float(row.get("착색력_a*", None))
                b = safe_to_float(row.get("착색력_b*", None))
                if None in (L, a, b):
                    continue
                s.at[idx, "_ΔE76_fix"] = delta_e76((L, a, b), ref)

    return s

# =========================
# UI Header
# =========================
st.title("액상 잉크 Lot 추적 관리 대시보드")
st.caption("✅ 빠른 검색  |  ✅ 잉크 입고(엑셀 누적)  |  ✅ 대시보드(목록/평균/추이)  |  ✅ 바인더 입출고(구글시트 자동 반영)")

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
    uploaded = st.file_uploader("또는 엑셀 업로드(업로드 모드: 서버 저장 보장 X)", type=["xlsx"])

# ✅ 업로드 파일은 '처음 1회만' tmp로 복사 (저장한 내용이 rerun 때 덮어써지는 문제 방지)
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
    headers=["일자", "바인더타입", "바인더명", "바인더 Lot", "반환량(kg)", "비고"]
)

# 파일 시그니처(위젯 key 충돌 방지)
file_sig = f"{Path(xlsx_path).name}:{Path(xlsx_path).stat().st_mtime_ns}"

# Load
raw = load_data(xlsx_path)
binder_df = raw["binder"].copy()
single_df_raw = raw["single"].copy()
spec_binder = raw["spec_binder"].copy()
spec_single = raw["spec_single"].copy()
base_lab = raw["base_lab"].copy()

# 단일색 파생/보정(중요!)
single_df = compute_single_derived(single_df_raw, spec_binder, spec_single, base_lab)

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
    # KPI(간단)
    b_total = len(binder_df)
    s_total = len(single_df)
    b_ng = int((binder_df.get("판정", pd.Series(dtype=str)) == "부적합").sum()) if "판정" in binder_df.columns else 0
    s_ng = int((single_df.get("_점도판정_fix", pd.Series(dtype=str)) == "부적합").sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("바인더 기록", f"{b_total:,}")
    c2.metric("바인더 부적합", f"{b_ng:,}")
    c3.metric("단일색 기록", f"{s_total:,}")
    c4.metric("단일색(점도) 부적합", f"{s_ng:,}")

    st.divider()

    # ---- 1) 목록(엑셀형)
    st.subheader("1) 단일색 데이터 목록 (엑셀형 보기)")
    need_cols = ["_입고일_dt", "색상군", "제품코드", "사용된 바인더 Lot", "_점도"]
    miss = [c for c in need_cols if c not in single_df.columns]
    if miss:
        st.warning(f"단일색 시트에서 필요한 컬럼을 찾지 못했습니다: {miss}")
    else:
        dmin, dmax = safe_date_bounds(single_df["_입고일_dt"])
        f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
        with f1:
            start = st.date_input("시작일(목록)", value=max(dmin, dmax - dt.timedelta(days=90)), key=f"list_start_{file_sig}")
        with f2:
            end = st.date_input("종료일(목록)", value=dmax, key=f"list_end_{file_sig}")
        with f3:
            cg_opts = sorted([x for x in single_df["색상군"].dropna().unique().tolist()])
            cg = st.multiselect("색상군(목록)", cg_opts, key=f"list_cg_{file_sig}")
        with f4:
            pc_opts = sorted([x for x in single_df["제품코드"].dropna().unique().tolist()])
            pc = st.multiselect("제품코드(목록)", pc_opts, key=f"list_pc_{file_sig}")

        if start > end:
            start, end = end, start

        df_list = single_df.copy()
        df_list = df_list[(df_list["_입고일_dt"].dt.date >= start) & (df_list["_입고일_dt"].dt.date <= end)]
        if cg:
            df_list = df_list[df_list["색상군"].isin(cg)]
        if pc:
            df_list = df_list[df_list["제품코드"].isin(pc)]

        view = pd.DataFrame({
            "제조일자": df_list["_입고일_dt"].dt.date,
            "색상군": df_list["색상군"],
            "제품코드": df_list["제품코드"],
            "사용된바인더": df_list.get("사용된 바인더 Lot", ""),
            "BinderType": df_list["_BinderType_fix"],
            "단일색Lot": df_list["_Lot_fix"],
            "점도(cP)": df_list["_점도"],
            "점도판정": df_list["_점도판정_fix"],
            "색차(ΔE76)": df_list["_ΔE76_fix"],
        }).sort_values(by="제조일자", ascending=False)

        st.dataframe(view, use_container_width=True, height=340)

        st.divider()

        # ---- 1-1) 평균 점도(점 + 값)
        st.subheader("1-1) 색상군별 평균 점도 (점 + 값 표시)")
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
            points = base.mark_circle(size=260)
            labels = base.mark_text(dx=12, dy=-10).encode(text="표시:N")
            st.altair_chart((points + labels).interactive(), use_container_width=True)

    st.divider()

    # ---- 2) 추이 그래프(Lot별)
    st.subheader("2) 단일색 점도 변화 추이 (Lot별)")
    st.caption("Lot 값이 엑셀 수식/캐시 문제로 None으로 읽혀도, 앱이 자동 복원해서 그래프가 정상 표시됩니다.")

    df_tr = single_df.copy()
    df_tr = df_tr.dropna(subset=["_입고일_dt", "_점도"])
    df_tr["Lot"] = df_tr["_Lot_fix"].astype(str).replace(["nan", "None"], "").str.strip()
    df_tr = df_tr[df_tr["Lot"] != ""]

    if len(df_tr) == 0:
        st.info("입고일/점도/Lot 값이 비어있어 추이 그래프를 표시할 수 없습니다.")
    else:
        dmin, dmax = safe_date_bounds(df_tr["_입고일_dt"])
        f1, f2, f3, f4, f5 = st.columns([1.2, 1.2, 1.6, 2.0, 1.0])
        with f1:
            start = st.date_input("시작일(추이)", value=max(dmin, dmax - dt.timedelta(days=90)), key=f"trend_start_{file_sig}")
        with f2:
            end = st.date_input("종료일(추이)", value=dmax, key=f"trend_end_{file_sig}")
        with f3:
            cg_opts = sorted([x for x in df_tr.get("색상군", pd.Series(dtype=object)).dropna().unique().tolist()]) if "색상군" in df_tr.columns else []
            cg = st.multiselect("색상군(추이)", cg_opts, key=f"trend_cg_{file_sig}")
        with f4:
            pc_opts = sorted([x for x in df_tr.get("제품코드", pd.Series(dtype=object)).dropna().unique().tolist()]) if "제품코드" in df_tr.columns else []
            pc = st.multiselect("제품코드(추이)", pc_opts, key=f"trend_pc_{file_sig}")
        with f5:
            show_labels = st.checkbox("라벨 표시", value=True, key=f"trend_labels_{file_sig}")

        if start > end:
            start, end = end, start

        df_tr = df_tr[(df_tr["_입고일_dt"].dt.date >= start) & (df_tr["_입고일_dt"].dt.date <= end)]
        if cg and "색상군" in df_tr.columns:
            df_tr = df_tr[df_tr["색상군"].isin(cg)]
        if pc and "제품코드" in df_tr.columns:
            df_tr = df_tr[df_tr["제품코드"].isin(pc)]

        lot_list = sorted(df_tr["Lot"].dropna().unique().tolist())
        default_pick = lot_list[-5:] if len(lot_list) > 5 else lot_list
        pick = st.multiselect("표시할 단일색 Lot(복수 선택)", lot_list, default=default_pick, key=f"trend_lots_{file_sig}")
        if pick:
            df_tr = df_tr[df_tr["Lot"].isin(pick)]

        if len(df_tr) == 0:
            st.info("선택한 조건에 해당하는 데이터가 없습니다. (기간/색상군/제품코드/로트 필터 확인)")
        else:
            df_tr = df_tr.sort_values("_입고일_dt")
            df_tr["점도표시"] = df_tr["_점도"].round(0).astype("Int64").astype(str)

            tooltip_cols = ["_입고일_dt:T", "Lot:N", "_점도:Q"]
            if "제품코드" in df_tr.columns:
                tooltip_cols.insert(2, "제품코드:N")
            if "색상군" in df_tr.columns:
                tooltip_cols.insert(3, "색상군:N")
            if "사용된 바인더 Lot" in df_tr.columns:
                tooltip_cols.append("사용된 바인더 Lot:N")

            base = alt.Chart(df_tr).encode(
                x=alt.X("_입고일_dt:T", title="입고일"),
                y=alt.Y("_점도:Q", title="점도(cP)"),
                tooltip=tooltip_cols
            )
            line = base.mark_line()
            points = base.mark_point(size=260).encode(color=alt.Color("Lot:N", title="Lot"))
            if show_labels:
                labels = base.mark_text(dy=-12).encode(
                    color=alt.Color("Lot:N", legend=None),
                    text="점도표시:N"
                )
                chart = (line + points + labels).interactive()
            else:
                chart = (line + points).interactive()

            st.altair_chart(chart, use_container_width=True)

    st.divider()

    st.subheader("최근 20건 (단일색) — Lot/판정 보정값 포함")
    show = single_df.copy()
    show = show.sort_values(by="_입고일_dt", ascending=False)
    show_view = pd.DataFrame({
        "입고일": show["_입고일_dt"].dt.date,
        "잉크타입": show.get("잉크타입\n(HEMA/Silicone)", ""),
        "색상군": show.get("색상군", ""),
        "제품코드": show.get("제품코드", ""),
        "단일색Lot(보정)": show["_Lot_fix"],
        "사용바인더Lot": show.get("사용된 바인더 Lot", ""),
        "BinderType(보정)": show["_BinderType_fix"],
        "점도(cP)": show["_점도"],
        "점도하한(보정)": show["_점도하한_fix"],
        "점도상한(보정)": show["_점도상한_fix"],
        "점도판정(보정)": show["_점도판정_fix"],
    })
    st.dataframe(show_view.head(20), use_container_width=True)

# =========================
# 잉크 입고 (단일색 입력만)
# =========================
with tab_ink_in:
    st.subheader("단일색 잉크 입력(입고)")
    st.info("이 탭은 **단일색_수입검사** 시트에 행을 Append하여 누적합니다. (동시 사용 시 충돌 가능)")

    ink_types = ["HEMA", "Silicone"]
    color_groups = sorted(spec_single.get("색상군", pd.Series(dtype=object)).dropna().unique().tolist())
    product_codes = sorted(spec_single.get("제품코드", pd.Series(dtype=object)).dropna().unique().tolist())

    binder_lots = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).tolist()
    binder_lots = sorted(set([x.strip() for x in binder_lots if x.strip()]), reverse=True)

    with st.form("single_form", clear_on_submit=True):
        col1, col2, col3, col4 = st.columns([1.2, 1.3, 1.5, 2.0])
        with col1:
            in_date = st.date_input("입고일", value=dt.date.today(), key=f"single_in_date_{file_sig}")
            ink_type = st.selectbox("잉크타입", ink_types, key=f"single_ink_type_{file_sig}")
            color_group = st.selectbox("색상군", color_groups, key=f"single_cg_{file_sig}")
        with col2:
            product_code = st.selectbox("제품코드", product_codes, key=f"single_pc_{file_sig}")
            binder_lot = st.selectbox("사용된 바인더 Lot", binder_lots, key=f"single_blot_{file_sig}")
        with col3:
            visc_meas = st.number_input("점도측정값(cP)", min_value=0.0, step=1.0, format="%.1f", key=f"single_visc_{file_sig}")
            supplier = st.selectbox("바인더제조처", ["내부", "외주"], index=0, key=f"single_supplier_{file_sig}")
        with col4:
            st.caption("선택: 착색력(L*a*b*) 입력 시, 기준LAB이 있으면 ΔE(76)을 계산해 '비고'에 기록합니다.")
            L = st.number_input("착색력_L*", value=0.0, step=0.1, format="%.2f", key=f"single_L_{file_sig}")
            a = st.number_input("착색력_a*", value=0.0, step=0.1, format="%.2f", key=f"single_a_{file_sig}")
            b = st.number_input("착색력_b*", value=0.0, step=0.1, format="%.2f", key=f"single_b_{file_sig}")
            lab_enabled = st.checkbox("L*a*b* 입력함", value=False, key=f"single_lab_en_{file_sig}")

        note = st.text_input("비고", value="", key=f"single_note_{file_sig}")
        submit_s = st.form_submit_button("저장(단일색)")

    if submit_s:
        # 기존 데이터(Lot 보정 포함)를 기반으로 다음 Lot 생성
        existing_lots = single_df["_Lot_fix"].astype(str).replace(["nan", "None"], "").str.strip()
        existing_lots = [x for x in existing_lots.tolist() if x]

        # 새 lot 생성
        pfx = lot_prefix_from_single(product_code, color_group, in_date)
        if not pfx:
            st.error("Lot 자동 생성 실패: 제품코드/색상군/입고일을 확인해주세요.")
        else:
            # pfx + seq 최댓값 찾기
            seqs = []
            for lot in existing_lots:
                if lot.startswith(pfx):
                    m = re.match(rf"^{re.escape(pfx)}(\d{{2,}})$", lot)
                    if m:
                        try:
                            seqs.append(int(m.group(1)))
                        except Exception:
                            pass
            seq = (max(seqs) + 1) if seqs else 1
            new_lot = f"{pfx}{seq:02d}"

            binder_type = infer_binder_type_from_lot(spec_binder, binder_lot)

            # 점도 기준 lookup
            hit = spec_single[
                (spec_single["색상군"].astype(str).str.strip() == str(color_group).strip()) &
                (spec_single["제품코드"].astype(str).str.strip() == str(product_code).strip())
            ].copy()
            if binder_type and "BinderType" in hit.columns and len(hit) > 0:
                hit2 = hit[hit["BinderType"].astype(str).str.strip() == str(binder_type).strip()]
                if len(hit2) > 0:
                    hit = hit2

            if len(hit) == 0:
                lo, hi, visc_judge = None, None, None
                st.warning("점도 기준을 Spec_Single_H&S에서 찾지 못했습니다. (색상군/제품코드/바인더타입 조합 확인)")
            else:
                lo = safe_to_float(hit.iloc[0].get("하한", None))
                hi = safe_to_float(hit.iloc[0].get("상한", None))
                visc_judge = judge_range(visc_meas, lo, hi)

            # ΔE 기록(비고에 남김)
            note2 = note
            if lab_enabled:
                base_hit = base_lab[base_lab.get("제품코드", pd.Series(dtype=str)).astype(str).str.strip() == str(product_code).strip()]
                if len(base_hit) == 1 and all(c in base_hit.columns for c in ["기준_L*", "기준_a*", "기준_b*"]):
                    ref = (
                        safe_to_float(base_hit.iloc[0].get("기준_L*", None)),
                        safe_to_float(base_hit.iloc[0].get("기준_a*", None)),
                        safe_to_float(base_hit.iloc[0].get("기준_b*", None)),
                    )
                    if None not in ref:
                        de = delta_e76((float(L), float(a), float(b)), ref)
                        note2 = (note2 + " " if note2 else "") + f"[ΔE76={de:.2f}]"
                    else:
                        note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"
                else:
                    note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"

            # ✅ 저장은 "헤더 norm_key 기준"으로만 매핑(기존 데이터 건드림 없음)
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

# =========================
# 바인더 입출고 (반품(kg) → 바인더입고 → 구글시트 최신순)
# =========================
with tab_binder:
    st.subheader("업체반환(반품) 입력 (kg 단위)")
    st.caption("※ 20kg(1통) 기준이더라도 실제 반환량은 kg 단위로 입력합니다.")

    binder_names = sorted(spec_binder.get("바인더명", pd.Series(dtype=object)).dropna().unique().tolist())
    binder_lots = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).tolist()
    binder_lots = sorted(set([x.strip() for x in binder_lots if x.strip()]), reverse=True)

    with st.form("binder_return_form", clear_on_submit=True):
        c1, c2, c3 = st.columns([1.2, 1.2, 2.6])
        with c1:
            r_date = st.date_input("반환일자", value=dt.date.today(), key=f"ret_date_{file_sig}")
        with c2:
            r_type = st.selectbox("바인더타입", ["HEMA", "Silicone"], key=f"ret_type_{file_sig}")
        with c3:
            r_name = st.selectbox("바인더명", binder_names, key=f"ret_name_{file_sig}")

        c4, c5, c6 = st.columns([2.0, 1.2, 2.8])
        with c4:
            r_lot = st.selectbox("바인더 Lot(선택)", ["(직접입력)"] + binder_lots, key=f"ret_lot_sel_{file_sig}")
            r_lot_text = st.text_input("바인더 Lot 직접입력", value="", key=f"ret_lot_text_{file_sig}") if r_lot == "(직접입력)" else ""
            final_lot = r_lot_text.strip() if r_lot == "(직접입력)" else r_lot
        with c5:
            r_kg = st.number_input("반환량(kg)", min_value=0.0, step=0.5, format="%.1f", key=f"ret_kg_{file_sig}")
        with c6:
            r_note = st.text_input("비고", value="", key=f"ret_note_{file_sig}")

        submit_ret = st.form_submit_button("반품 저장")

    if submit_ret:
        if r_kg <= 0:
            st.error("반환량(kg)은 0보다 커야 합니다.")
        else:
            row = {
                norm_key("일자"): r_date,
                norm_key("바인더타입"): r_type,
                norm_key("바인더명"): r_name,
                norm_key("바인더 Lot"): final_lot,
                norm_key("반환량(kg)"): float(r_kg),
                norm_key("비고"): r_note,
            }
            try:
                append_row_to_sheet(xlsx_path, SHEET_BINDER_RETURN, row)
                st.success("반품 저장 완료!")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"반품 저장 실패: {e}")

    st.divider()

    st.subheader("바인더 입력 (제조/입고) — 여러 Lot/날짜 묶음 입력 지원")
    st.caption("※ 여러 날짜의 Lot가 한 번에 입고되는 상황을 고려해, 날짜별/수량별 묶음 입력을 지원합니다.")

    input_mode = st.radio("입력 방식", ["개별 입력", "묶음 입력(여러 날짜/수량)"], horizontal=True, key=f"binder_input_mode_{file_sig}")

    existing_binder_lots = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).tolist()
    existing_binder_lots = [x.strip() for x in existing_binder_lots if x.strip()]

    if input_mode == "개별 입력":
        with st.form("binder_form_single", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                mfg_date = st.date_input("제조/입고일", value=dt.date.today(), key=f"b_single_date_{file_sig}")
                b_name = st.selectbox("바인더명", binder_names, key=f"b_single_name_{file_sig}")
            with col2:
                visc = st.number_input("점도(cP)", min_value=0.0, step=1.0, format="%.1f", key=f"b_single_visc_{file_sig}")
                uv = st.number_input("UV흡광도(선택)", min_value=0.0, step=0.01, format="%.3f", key=f"b_single_uv_{file_sig}")
                uv_enabled = st.checkbox("UV 값 입력함", value=False, key=f"b_single_uv_en_{file_sig}")
            with col3:
                note = st.text_input("비고", value="", key=f"b_single_note_{file_sig}")
                submit_b = st.form_submit_button("저장(바인더)")

        if submit_b:
            visc_lo, visc_hi, uv_hi, _ = get_binder_limits(spec_binder, b_name)
            lot = generate_binder_lot(spec_binder, b_name, mfg_date, existing_binder_lots)

            judge_v = judge_range(visc, visc_lo, visc_hi)
            judge_u = judge_range(uv if uv_enabled else None, None, uv_hi)
            judge = "부적합" if (judge_v == "부적합" or judge_u == "부적합") else "적합"

            row = {
                norm_key("제조/입고일"): mfg_date,
                norm_key("바인더명"): b_name,
                norm_key("Lot(자동)"): lot,
                norm_key("점도(cP)"): float(visc),
                norm_key("UV흡광도(선택)"): float(uv) if uv_enabled else None,
                norm_key("판정"): judge,
                norm_key("비고"): note,
            }
            try:
                append_row_to_sheet(xlsx_path, SHEET_BINDER, row)
                st.success(f"저장 완료! 바인더 Lot = {lot}")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    else:
        st.caption("표에 날짜/바인더명/수량(통)/점도/UV/비고를 입력하고 한 번에 저장하세요. (매일 제조 X 상황 대응)")

        # 기본 3줄 템플릿
        base_rows = st.session_state.get(f"binder_batch_rows_{file_sig}")
        if base_rows is None:
            base_rows = [
                {"제조/입고일": dt.date.today(), "바인더명": (binder_names[0] if binder_names else ""), "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
                {"제조/입고일": dt.date.today() - dt.timedelta(days=1), "바인더명": (binder_names[0] if binder_names else ""), "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
                {"제조/입고일": dt.date.today() - dt.timedelta(days=2), "바인더명": (binder_names[0] if binder_names else ""), "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
            ]
        edit_df = st.data_editor(pd.DataFrame(base_rows), use_container_width=True, num_rows="dynamic", key=f"binder_batch_editor_{file_sig}")
        submit_batch = st.button("묶음 저장(바인더)", type="primary", key=f"binder_batch_submit_{file_sig}")

        if submit_batch:
            tmp = edit_df.copy()
            tmp["제조/입고일"] = tmp["제조/입고일"].apply(normalize_date)
            tmp["수량(통)"] = pd.to_numeric(tmp["수량(통)"], errors="coerce").fillna(0).astype(int)
            tmp["점도(cP)"] = pd.to_numeric(tmp["점도(cP)"].astype(str).str.replace(",", "", regex=False), errors="coerce")

            tmp = tmp.dropna(subset=["제조/입고일", "바인더명", "점도(cP)"])
            tmp = tmp[tmp["수량(통)"] > 0]

            if len(tmp) == 0:
                st.error("저장할 행이 없습니다. (날짜/바인더명/수량/점도 입력 확인)")
            else:
                rows_out = []
                existing_list = existing_binder_lots[:]
                seq_counters = {}
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
                        st.error(f"[{b_name}] 순번(-##)이 없어 수량 {qty}를 자동 Lot로 생성할 수 없습니다.")
                        st.stop()

                    key = (prefix, date_str)
                    if key not in seq_counters:
                        seq_counters[key] = next_seq_for_pattern(existing_list, prefix, date_str, sep="-")

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

                        row = {
                            norm_key("제조/입고일"): mfg_date,
                            norm_key("바인더명"): b_name,
                            norm_key("Lot(자동)"): lot,
                            norm_key("점도(cP)"): float(visc),
                            norm_key("UV흡광도(선택)"): float(uv_val) if uv_enabled and uv_val is not None else None,
                            norm_key("판정"): judge,
                            norm_key("비고"): note,
                        }
                        rows_out.append(row)
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

    if st.button("지금 최신값으로 다시 불러오기", key=f"binder_refresh_{file_sig}"):
        st.cache_data.clear()
        st.rerun()

# =========================
# Search
# =========================
with tab_search:
    st.subheader("빠른 검색")
    c1, c2, c3 = st.columns([2, 2, 3])
    with c1:
        mode = st.selectbox("검색 종류", ["바인더 Lot", "단일색 잉크 Lot", "제품코드", "색상군", "기간(입고일)"], key=f"search_mode_{file_sig}")
    with c2:
        q = st.text_input("검색어", placeholder="예: PCB20250112-01 / PLB25041501 ...", key=f"search_q_{file_sig}")
    with c3:
        st.caption("💡 단일색 Lot/제품코드/색상군/기간으로 빠르게 필터링합니다.")

    if mode == "기간(입고일)":
        d1, d2 = st.columns(2)
        with d1:
            start = st.date_input("시작일", value=dt.date.today() - dt.timedelta(days=30), key=f"search_start_{file_sig}")
        with d2:
            end = st.date_input("종료일", value=dt.date.today(), key=f"search_end_{file_sig}")

        df = single_df.copy()
        df = df.dropna(subset=["_입고일_dt"])
        df = df[(df["_입고일_dt"].dt.date >= start) & (df["_입고일_dt"].dt.date <= end)]

        out = pd.DataFrame({
            "입고일": df["_입고일_dt"].dt.date,
            "색상군": df.get("색상군", ""),
            "제품코드": df.get("제품코드", ""),
            "단일색Lot(보정)": df["_Lot_fix"],
            "사용바인더Lot": df.get("사용된 바인더 Lot", ""),
            "BinderType(보정)": df["_BinderType_fix"],
            "점도(cP)": df["_점도"],
            "점도판정(보정)": df["_점도판정_fix"],
        }).sort_values(by="입고일", ascending=False)
        st.dataframe(out, use_container_width=True)

    elif mode == "바인더 Lot":
        b = binder_df.copy()
        if q:
            b = b[b.astype(str).apply(lambda r: r.str.contains(str(q).strip(), case=False, na=False)).any(axis=1)]
        st.subheader("바인더_제조_입고")
        st.dataframe(b.sort_values(by="제조/입고일", ascending=False) if "제조/입고일" in b.columns else b, use_container_width=True)

        if q and "사용된 바인더 Lot" in single_df.columns:
            s_hit = single_df[single_df["사용된 바인더 Lot"].astype(str).str.contains(str(q).strip(), case=False, na=False)]
            st.subheader("연결된 단일색 (사용된 바인더 Lot)")
            out = pd.DataFrame({
                "입고일": s_hit["_입고일_dt"].dt.date,
                "색상군": s_hit.get("색상군", ""),
                "제품코드": s_hit.get("제품코드", ""),
                "단일색Lot(보정)": s_hit["_Lot_fix"],
                "점도(cP)": s_hit["_점도"],
                "점도판정(보정)": s_hit["_점도판정_fix"],
            }).sort_values(by="입고일", ascending=False)
            st.dataframe(out, use_container_width=True)

    elif mode == "단일색 잉크 Lot":
        s = single_df.copy()
        s["Lot검색"] = s["_Lot_fix"].astype(str)
        if q:
            s = s[s["Lot검색"].str.contains(str(q).strip(), case=False, na=False)]
        out = pd.DataFrame({
            "입고일": s["_입고일_dt"].dt.date,
            "색상군": s.get("색상군", ""),
            "제품코드": s.get("제품코드", ""),
            "단일색Lot(보정)": s["_Lot_fix"],
            "사용바인더Lot": s.get("사용된 바인더 Lot", ""),
            "BinderType(보정)": s["_BinderType_fix"],
            "점도(cP)": s["_점도"],
            "점도판정(보정)": s["_점도판정_fix"],
        }).sort_values(by="입고일", ascending=False)
        st.dataframe(out, use_container_width=True)

    elif mode == "제품코드":
        s = single_df.copy()
        if q:
            s = s[s.get("제품코드", "").astype(str).str.contains(str(q).strip(), case=False, na=False)]
        out = pd.DataFrame({
            "입고일": s["_입고일_dt"].dt.date,
            "색상군": s.get("색상군", ""),
            "제품코드": s.get("제품코드", ""),
            "단일색Lot(보정)": s["_Lot_fix"],
            "점도(cP)": s["_점도"],
            "점도판정(보정)": s["_점도판정_fix"],
        }).sort_values(by="입고일", ascending=False)
        st.dataframe(out, use_container_width=True)

    elif mode == "색상군":
        s = single_df.copy()
        if q:
            s = s[s.get("색상군", "").astype(str).str.contains(str(q).strip(), case=False, na=False)]
        out = pd.DataFrame({
            "입고일": s["_입고일_dt"].dt.date,
            "색상군": s.get("색상군", ""),
            "제품코드": s.get("제품코드", ""),
            "단일색Lot(보정)": s["_Lot_fix"],
            "점도(cP)": s["_점도"],
            "점도판정(보정)": s["_점도판정_fix"],
        }).sort_values(by="입고일", ascending=False)
        st.dataframe(out, use_container_width=True)
