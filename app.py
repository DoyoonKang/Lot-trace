import altair as alt
import streamlit as st
import pandas as pd
import datetime as dt
import re
from pathlib import Path
from io import StringIO
import requests
from openpyxl import load_workbook

st.set_page_config(
    page_title="액상 잉크 Lot 추적 관리",
    page_icon="🧪",
    layout="wide",
)

# =========================================================
# Config
# =========================================================
DEFAULT_XLSX = "액상잉크_Lot추적관리_FINAL.xlsx"

SHEET_BINDER = "바인더_제조_입고"
SHEET_SINGLE = "단일색_수입검사"
SHEET_SPEC_BINDER = "Spec_Binder"
SHEET_SPEC_SINGLE = "Spec_Single_H&S"
SHEET_BASE_LAB = "기준LAB"
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

# =========================================================
# Utils (text / columns)
# =========================================================
def norm_key(x) -> str:
    """컬럼/헤더 비교를 위해: 줄바꿈 제거 + 공백 정리 + 양끝 공백 제거"""
    if x is None:
        return ""
    s = str(x)
    s = s.replace("\n", " ").replace("\r", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def normalize_df_columns(df: pd.DataFrame) -> pd.DataFrame:
    """pandas DataFrame 컬럼명을 정규화(줄바꿈/공백)해서 내부 처리 일관성 확보"""
    df = df.copy()
    cols = [norm_key(c) for c in df.columns]
    # 중복 컬럼명 방지
    seen = {}
    new_cols = []
    for c in cols:
        if c not in seen:
            seen[c] = 0
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}__{seen[c]}")
    df.columns = new_cols
    return df

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """정규화된 컬럼명 기준으로: (1) 정확 일치 -> (2) 포함/유사 매칭"""
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    norm_map = {c: norm_key(c) for c in cols}
    # 1) exact
    cand_norms = [norm_key(c) for c in candidates]
    for c in cols:
        if norm_map[c] in cand_norms:
            return c
    # 2) contains (most strict first)
    for cn in cand_norms:
        for c in cols:
            if cn and (cn in norm_map[c] or norm_map[c].startswith(cn) or cn.startswith(norm_map[c])):
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
    if x is None or (isinstance(x, float) and pd.isna(x)) or (isinstance(x, str) and x.strip() == ""):
        return None
    if isinstance(x, (dt.date, dt.datetime)):
        return x.date() if isinstance(x, dt.datetime) else x
    try:
        d = pd.to_datetime(x, errors="coerce")
        if pd.isna(d):
            return None
        return d.date()
    except Exception:
        return None

def safe_date_bounds(s: pd.Series):
    s = pd.to_datetime(s, errors="coerce").dropna()
    if len(s) == 0:
        today = dt.date.today()
        return today, today
    return s.min().date(), s.max().date()

def judge_range(value, lo, hi):
    v = safe_to_float(value)
    if v is None:
        return None
    if lo is not None and v < float(lo):
        return "부적합"
    if hi is not None and v > float(hi):
        return "부적합"
    return "적합"

def delta_e76(lab1, lab2):
    return float(((lab1[0]-lab2[0])**2 + (lab1[1]-lab2[1])**2 + (lab1[2]-lab2[2])**2) ** 0.5)

# =========================================================
# Google Sheets reader
# =========================================================
@st.cache_data(ttl=60, show_spinner=False)
def read_gsheet_csv(sheet_id: str, sheet_name: str) -> pd.DataFrame:
    base = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
    r = requests.get(base, params={"tqx": "out:csv", "sheet": sheet_name}, timeout=20)
    r.raise_for_status()
    r.encoding = "utf-8"
    return pd.read_csv(StringIO(r.text))

# =========================================================
# Excel IO
# =========================================================
def ensure_sheet_exists(xlsx_path: str, sheet_name: str, headers: list[str]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        wb.save(xlsx_path)

@st.cache_data(show_spinner=False)
def load_excel(xlsx_path: str) -> dict:
    """캐시 로딩 (쓰기 후에는 st.cache_data.clear() 필요)"""
    def read(name: str) -> pd.DataFrame:
        return pd.read_excel(xlsx_path, sheet_name=name)
    return {
        "binder": read(SHEET_BINDER),
        "single": read(SHEET_SINGLE),
        "spec_binder": read(SHEET_SPEC_BINDER),
        "spec_single": read(SHEET_SPEC_SINGLE),
        "base_lab": read(SHEET_BASE_LAB),
    }

def append_row_to_sheet_by_norm(xlsx_path: str, sheet_name: str, row_by_norm: dict):
    """엑셀 1행 헤더(원본) 기준으로 append. row_by_norm 키는 norm_key(헤더)로 준다."""
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    out = []
    for h in headers:
        if h is None:
            out.append(None)
            continue
        out.append(row_by_norm.get(norm_key(h), None))
    ws.append(out)
    wb.save(xlsx_path)

def append_rows_to_sheet_by_norm(xlsx_path: str, sheet_name: str, rows_by_norm: list[dict]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    for row_by_norm in rows_by_norm:
        out = []
        for h in headers:
            if h is None:
                out.append(None)
                continue
            out.append(row_by_norm.get(norm_key(h), None))
        ws.append(out)
    wb.save(xlsx_path)

# =========================================================
# Spec helpers
# =========================================================
def get_binder_limits(spec_binder: pd.DataFrame, binder_name: str):
    df = spec_binder[spec_binder["바인더명"].astype(str) == str(binder_name)].copy()
    visc = df[df["시험항목"].astype(str).str.contains("점도", na=False)]
    uv = df[df["시험항목"].astype(str).str.contains("UV", na=False)]

    visc_lo = safe_to_float(visc["하한"].dropna().iloc[0]) if len(visc["하한"].dropna()) else None
    visc_hi = safe_to_float(visc["상한"].dropna().iloc[0]) if len(visc["상한"].dropna()) else None
    uv_hi = safe_to_float(uv["상한"].dropna().iloc[0]) if len(uv["상한"].dropna()) else None
    rule = df["Lot부여규칙"].dropna().iloc[0] if "Lot부여규칙" in df.columns and len(df["Lot부여규칙"].dropna()) else None
    return visc_lo, visc_hi, uv_hi, rule

def infer_binder_type_from_lot(spec_binder: pd.DataFrame, binder_lot: str):
    """Spec_Binder의 Lot부여규칙 prefix로 바인더명을 역추정(=BinderType(자동) 값으로 사용)"""
    if not binder_lot:
        return None
    lot = str(binder_lot).strip()
    rules = (
        spec_binder[["바인더명", "Lot부여규칙"]]
        .dropna()
        .drop_duplicates()
        .to_dict("records")
    )
    for r in rules:
        rule = str(r["Lot부여규칙"]).strip()
        m = re.match(r"^([A-Za-z0-9]+)\+", rule)
        if m and lot.startswith(m.group(1)):
            return str(r["바인더명"])
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
    m = re.match(r"^([A-Za-z0-9]+)\+YYYYMMDD(-##)?$", str(rule).strip()) if rule else None
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

def generate_single_lot_prefix(product_code: str, color_group: str, in_date: dt.date):
    code = (product_code or "").strip()
    color_code = COLOR_CODE.get(color_group)
    if not color_code or not in_date:
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
    return f"{prefix}{color_code}{date_str}"

def compute_single_lots(df: pd.DataFrame, col_in_date: str, col_pc: str, col_cg: str, col_lot_existing: str | None):
    """기존 lot 결과가 비어있거나(수식 캐시 유실) 중복이 많아도, '표시용' lot를 안정적으로 재생성"""
    out = pd.Series([None] * len(df), index=df.index, dtype="object")

    exist = None
    if col_lot_existing and col_lot_existing in df.columns:
        exist = df[col_lot_existing].astype(str)
        exist = exist.where(~exist.isna(), "")
        exist = exist.replace(["nan", "None"], "", regex=False).astype(str)

    # 1) 가능한 기존 값 먼저 사용
    if exist is not None:
        out = exist.where(exist.str.strip() != "", None)

    # 2) 비어있는 행은 규칙 기반 생성 (prefix+date) + seq
    #    seq는 동일 prefix+date 안에서 기존 lot의 최대 seq 이후부터 이어서 부여
    df2 = df.copy()
    df2["_d"] = pd.to_datetime(df2[col_in_date], errors="coerce").dt.date
    df2["_prefix"] = df2.apply(lambda r: generate_single_lot_prefix(str(r.get(col_pc, "")).strip(), str(r.get(col_cg, "")).strip(), r.get("_d")), axis=1)

    # 기존 seq 파싱
    max_seq = {}
    if exist is not None:
        for v, p in zip(exist.tolist(), df2["_prefix"].tolist()):
            if not p:
                continue
            if not v or str(v).strip() == "":
                continue
            sv = str(v).strip()
            if not sv.startswith(p):
                continue
            rest = sv[len(p):]
            m = re.match(r"^(\d{2,})$", rest)
            if m:
                try:
                    max_seq[p] = max(max_seq.get(p, 0), int(m.group(1)))
                except Exception:
                    pass

    # 생성
    # 날짜/행순서로 안정적 재현
    order = df2.sort_values(by=["_d"]).index.tolist()
    counters = {}
    for idx in order:
        if pd.notna(out.loc[idx]) and str(out.loc[idx]).strip() != "":
            continue
        p = df2.loc[idx, "_prefix"]
        if not p:
            continue
        if p not in counters:
            counters[p] = max_seq.get(p, 0) + 1
        seq = counters[p]
        counters[p] += 1
        out.loc[idx] = f"{p}{seq:02d}"

    return out

def compute_binder_lots(df: pd.DataFrame, col_date: str, col_name: str, col_lot_existing: str | None, spec_binder: pd.DataFrame):
    """바인더 lot(표시용) 재생성 (수식 캐시 유실 대응)"""
    out = pd.Series([None] * len(df), index=df.index, dtype="object")

    exist = None
    if col_lot_existing and col_lot_existing in df.columns:
        exist = df[col_lot_existing].astype(str)
        exist = exist.where(~exist.isna(), "")
        exist = exist.replace(["nan", "None"], "", regex=False).astype(str)
        out = exist.where(exist.str.strip() != "", None)

    # rule prefix 기반 생성
    df2 = df.copy()
    df2["_d"] = pd.to_datetime(df2[col_date], errors="coerce").dt.date
    df2["_name"] = df2[col_name].astype(str).str.strip()

    # rule prefix/seq 여부
    rule_map = {}
    for _, r in spec_binder.dropna(subset=["바인더명", "Lot부여규칙"]).drop_duplicates(subset=["바인더명"]).iterrows():
        m = re.match(r"^([A-Za-z0-9]+)\+YYYYMMDD(-##)?$", str(r["Lot부여규칙"]).strip())
        if m:
            rule_map[str(r["바인더명"])] = (m.group(1), bool(m.group(2)))

    # 기존 seq 파싱
    max_seq = {}
    if exist is not None:
        for lot, name, d in zip(exist.tolist(), df2["_name"].tolist(), df2["_d"].tolist()):
            if not lot or str(lot).strip() == "":
                continue
            if not name or name not in rule_map or not d:
                continue
            prefix, has_seq = rule_map[name]
            ds = d.strftime("%Y%m%d")
            base = f"{prefix}{ds}"
            if not str(lot).startswith(base):
                continue
            if has_seq:
                rest = str(lot)[len(base):]
                if rest.startswith("-"):
                    rest = rest[1:]
                m = re.match(r"^(\d+)", rest)
                if m:
                    try:
                        key = (prefix, ds)
                        max_seq[key] = max(max_seq.get(key, 0), int(m.group(1)))
                    except Exception:
                        pass

    counters = {}
    order = df2.sort_values(by=["_d"]).index.tolist()
    for idx in order:
        if pd.notna(out.loc[idx]) and str(out.loc[idx]).strip() != "":
            continue
        name = df2.loc[idx, "_name"]
        d = df2.loc[idx, "_d"]
        if not name or not d or name not in rule_map:
            continue
        prefix, has_seq = rule_map[name]
        ds = d.strftime("%Y%m%d")
        if has_seq:
            key = (prefix, ds)
            if key not in counters:
                counters[key] = max_seq.get(key, 0) + 1
            seq = counters[key]
            counters[key] += 1
            out.loc[idx] = f"{prefix}{ds}-{seq:02d}"
        else:
            out.loc[idx] = f"{prefix}{ds}"

    return out

def compute_single_spec_row(spec_single: pd.DataFrame, color_group: str, product_code: str, binder_type: str | None):
    """Spec_Single_H&S에서 점도 하한/상한 찾기"""
    df = spec_single.copy()
    if "색상군" in df.columns:
        df = df[df["색상군"].astype(str) == str(color_group)]
    if "제품코드" in df.columns:
        df = df[df["제품코드"].astype(str) == str(product_code)]
    # BinderType 컬럼이 있으면 필터
    bt_col = find_col(df, ["BinderType", "BinderType(자동)", "바인더타입", "바인더 타입", "Binder Type"])
    if bt_col and binder_type:
        df2 = df[df[bt_col].astype(str) == str(binder_type)]
        if len(df2) > 0:
            df = df2
    if len(df) == 0:
        return None, None
    lo = safe_to_float(df["하한"].iloc[0]) if "하한" in df.columns else None
    hi = safe_to_float(df["상한"].iloc[0]) if "상한" in df.columns else None
    return lo, hi

def extract_or_compute_de76(single_view: pd.DataFrame, base_lab: pd.DataFrame) -> pd.Series:
    base = base_lab.copy()
    if "제품코드" in base.columns:
        base["제품코드"] = base["제품코드"].astype(str).str.strip()

    out = pd.Series([None] * len(single_view), index=single_view.index, dtype="float")

    if "비고" in single_view.columns:
        pat = re.compile(r"\[\s*ΔE76\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*\]")
        for idx, val in single_view["비고"].items():
            if pd.isna(val):
                continue
            m = pat.search(str(val))
            if m:
                try:
                    out.loc[idx] = float(m.group(1))
                except Exception:
                    pass

    need_cols = ["제품코드", "착색력_L*", "착색력_a*", "착색력_b*"]
    if all(c in single_view.columns for c in need_cols) and all(c in base.columns for c in ["기준_L*", "기준_a*", "기준_b*", "제품코드"]):
        base_map = base.set_index("제품코드")[["기준_L*", "기준_a*", "기준_b*"]].to_dict("index")
        for idx, row in single_view.iterrows():
            if pd.notna(out.loc[idx]):
                continue
            pc = row.get("제품코드", None)
            if pc is None or pd.isna(pc):
                continue
            pc = str(pc).strip()
            if pc not in base_map:
                continue
            L = safe_to_float(row.get("착색력_L*", None))
            a = safe_to_float(row.get("착색력_a*", None))
            b = safe_to_float(row.get("착색력_b*", None))
            if None in (L, a, b):
                continue
            ref = base_map[pc]
            ref_lab = (safe_to_float(ref["기준_L*"]), safe_to_float(ref["기준_a*"]), safe_to_float(ref["기준_b*"]))
            if None in ref_lab:
                continue
            out.loc[idx] = delta_e76((L, a, b), ref_lab)
    return out

# =========================================================
# Derived views (핵심: 엑셀 수식 캐시 유실에도 값이 안 사라지게 "앱에서 재계산")
# =========================================================
def build_views(binder_raw: pd.DataFrame, single_raw: pd.DataFrame, spec_binder_raw: pd.DataFrame, spec_single_raw: pd.DataFrame, base_lab_raw: pd.DataFrame):
    binder = normalize_df_columns(binder_raw)
    single = normalize_df_columns(single_raw)
    spec_binder = normalize_df_columns(spec_binder_raw)
    spec_single = normalize_df_columns(spec_single_raw)
    base_lab = normalize_df_columns(base_lab_raw)

    # ---- binder view
    b_date = find_col(binder, ["제조/입고일", "제조입고일", "입고일", "일자"])
    b_name = find_col(binder, ["바인더명", "Binder", "바인더"])
    b_lot = find_col(binder, ["Lot(자동)", "Lot", "LOT"])
    b_visc = find_col(binder, ["점도(cP)", "점도", "Viscosity"])
    b_uv = find_col(binder, ["UV흡광도(선택)", "UV흡광도", "UV"])

    binder_view = binder.copy()
    if b_date:
        binder_view["_date"] = pd.to_datetime(binder_view[b_date], errors="coerce").dt.date
    else:
        binder_view["_date"] = None

    # lot (표시용)
    if b_date and b_name:
        binder_view["_lot_calc"] = compute_binder_lots(binder_view, b_date, b_name, b_lot, spec_binder)
    else:
        binder_view["_lot_calc"] = binder_view[b_lot] if b_lot else None

    # 판정(표시용) - 수식 캐시 유실 대응
    if b_name and b_visc:
        lo_hi = {}
        for bn in spec_binder.get("바인더명", pd.Series(dtype=object)).dropna().unique().tolist():
            lo, hi, uv_hi, _ = get_binder_limits(spec_binder, bn)
            lo_hi[str(bn)] = (lo, hi, uv_hi)
        def _bj(r):
            name = str(r.get(b_name, "")).strip()
            v = safe_to_float(r.get(b_visc, None))
            u = safe_to_float(r.get(b_uv, None)) if (b_uv and pd.notna(r.get(b_uv, None))) else None
            if name not in lo_hi:
                return None
            lo, hi, uv_hi = lo_hi[name]
            jv = judge_range(v, lo, hi)
            ju = judge_range(u, None, uv_hi) if u is not None else None
            if jv == "부적합" or ju == "부적합":
                return "부적합"
            if jv == "적합" or ju == "적합":
                return "적합"
            return None
        binder_view["_judge_calc"] = binder_view.apply(_bj, axis=1)
    else:
        binder_view["_judge_calc"] = None

    # ---- single view
    s_date = find_col(single, ["입고일", "제조일자", "제조/입고일", "날짜"])
    s_type = find_col(single, ["잉크타입 (HEMA/Silicone)", "잉크타입", "InkType"])
    s_cg = find_col(single, ["색상군", "ColorGroup"])
    s_pc = find_col(single, ["제품코드", "ProductCode"])
    s_lot = find_col(single, ["단일색잉크 Lot", "단일색 잉크 Lot", "단일색잉크Lot", "단일색Lot"])
    s_blot = find_col(single, ["사용된 바인더 Lot", "사용 바인더 Lot", "바인더 Lot", "BinderLot"])
    s_visc = find_col(single, ["점도측정값(cP)", "점도측정값 (cP)", "점도(cP)", "점도측정값"])

    single_view = single.copy()
    single_view["_date"] = pd.to_datetime(single_view[s_date], errors="coerce").dt.date if s_date else None
    # lot calc (수식 캐시 유실 대응)
    if s_date and s_pc and s_cg:
        single_view["_lot_calc"] = compute_single_lots(single_view, s_date, s_pc, s_cg, s_lot)
    else:
        single_view["_lot_calc"] = single_view[s_lot] if s_lot else None

    # binder type calc
    if s_blot:
        single_view["_binder_type_calc"] = single_view[s_blot].apply(lambda x: infer_binder_type_from_lot(spec_binder, x))
    else:
        single_view["_binder_type_calc"] = None

    # spec/judge calc
    def _spec_lo(r):
        if not (s_cg and s_pc):
            return None
        return compute_single_spec_row(spec_single, str(r.get(s_cg, "")).strip(), str(r.get(s_pc, "")).strip(), r.get("_binder_type_calc"))[0]
    def _spec_hi(r):
        if not (s_cg and s_pc):
            return None
        return compute_single_spec_row(spec_single, str(r.get(s_cg, "")).strip(), str(r.get(s_pc, "")).strip(), r.get("_binder_type_calc"))[1]

    single_view["_spec_lo"] = single_view.apply(_spec_lo, axis=1) if (s_cg and s_pc) else None
    single_view["_spec_hi"] = single_view.apply(_spec_hi, axis=1) if (s_cg and s_pc) else None

    if s_visc:
        single_view["_visc"] = single_view[s_visc].apply(safe_to_float)
        single_view["_judge"] = single_view.apply(lambda r: judge_range(r.get("_visc"), r.get("_spec_lo"), r.get("_spec_hi")), axis=1)
    else:
        single_view["_visc"] = None
        single_view["_judge"] = None

    # ΔE76
    # NOTE: 원본 컬럼명(착색력_*)는 정규화 후에도 동일하다고 가정
    single_view["_ΔE76"] = extract_or_compute_de76(single_view, base_lab)

    # display columns (정규화된 원본 유지 + 파생)
    return binder_view, single_view, spec_binder, spec_single, base_lab

# =========================================================
# Spec editor (대시보드에서 하한/상한 수정)
# =========================================================
def update_spec_single_bounds(xlsx_path: str, edited_df: pd.DataFrame):
    """Spec_Single_H&S 시트에서 (색상군, 제품코드, BinderType) 키로 하한/상한만 업데이트"""
    wb = load_workbook(xlsx_path)
    if SHEET_SPEC_SINGLE not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {SHEET_SPEC_SINGLE}")
    ws = wb[SHEET_SPEC_SINGLE]

    headers = [c.value for c in ws[1]]
    hmap = {norm_key(h): i+1 for i, h in enumerate(headers) if h is not None}

    # required
    ck = [k for k in ["색상군", "제품코드"] if k not in hmap]
    if ck:
        raise ValueError(f"Spec_Single_H&S 헤더 누락: {ck}")

    col_cg = hmap["색상군"]
    col_pc = hmap["제품코드"]
    col_bt = hmap.get("bindertype") or hmap.get("bindertype(자동)") or hmap.get("바인더타입") or hmap.get("binder type")
    col_lo = hmap.get("하한")
    col_hi = hmap.get("상한")
    if col_lo is None or col_hi is None:
        raise ValueError("Spec_Single_H&S에서 '하한' 또는 '상한' 컬럼을 찾지 못했습니다.")

    # build key -> row index
    key_to_row = {}
    for r in range(2, ws.max_row + 1):
        cg = ws.cell(row=r, column=col_cg).value
        pc = ws.cell(row=r, column=col_pc).value
        bt = ws.cell(row=r, column=col_bt).value if col_bt else None
        key = (str(cg).strip() if cg is not None else "", str(pc).strip() if pc is not None else "", str(bt).strip() if bt is not None else "")
        key_to_row[key] = r

    updated = 0
    for _, row in edited_df.iterrows():
        cg = str(row.get("색상군", "")).strip()
        pc = str(row.get("제품코드", "")).strip()
        bt = str(row.get("BinderType", "")).strip() if "BinderType" in edited_df.columns else ""
        key = (cg, pc, bt)

        if key not in key_to_row:
            # BinderType 없는 키로도 시도
            key2 = (cg, pc, "")
            if key2 not in key_to_row:
                continue
            r = key_to_row[key2]
        else:
            r = key_to_row[key]

        ws.cell(row=r, column=col_lo).value = safe_to_float(row.get("하한", None))
        ws.cell(row=r, column=col_hi).value = safe_to_float(row.get("상한", None))
        updated += 1

    wb.save(xlsx_path)
    return updated

# =========================================================
# UI - File selection
# =========================================================
st.title("액상 잉크 Lot 추적 관리 대시보드")
st.caption("✅ 대시보드(목록/평균/추이)  |  ✅ 잉크 입고(엑셀 누적)  |  ✅ 바인더 입출고(구글시트 최신순)  |  ✅ 반품(kg) 기록")

with st.sidebar:
    st.header("데이터 파일")
    xlsx_path = st.text_input(
        "엑셀 파일 경로",
        value=DEFAULT_XLSX,
        help="로컬 실행 시, app.py와 같은 폴더에 엑셀을 두면 기본값 그대로 사용 가능합니다."
    )
    uploaded = st.file_uploader("또는 엑셀 업로드(업로드 모드: 서버 저장 보장 X)", type=["xlsx"])

# 업로드 파일은 '처음 1회만' tmp로 복사 (rerun 때 덮어쓰기 방지)
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

ensure_sheet_exists(
    xlsx_path,
    SHEET_BINDER_RETURN,
    headers=["일자", "바인더타입", "바인더명", "바인더 Lot", "반환량(kg)", "비고"]
)

# Load
raw = load_excel(xlsx_path)
binder_view, single_view, spec_binder, spec_single, base_lab = build_views(
    raw["binder"], raw["single"], raw["spec_binder"], raw["spec_single"], raw["base_lab"]
)

single_ver = str(pd.to_datetime(single_view.get("_date", pd.Series(dtype=object)), errors="coerce").max())

# Tabs
tab_dash, tab_ink_in, tab_binder, tab_search = st.tabs(
    ["📊 대시보드", "✍️ 잉크 입고", "📦 바인더 입출고", "🔎 빠른검색"]
)

# =========================================================
# Dashboard (그래프/표는 여기만)
# =========================================================
with tab_dash:
    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("바인더 기록", f"{len(binder_view):,}")
    c2.metric("바인더 부적합(표시용)", f"{int((binder_view.get('_judge_calc')=='부적합').sum()):,}")
    c3.metric("단일색 기록", f"{len(single_view):,}")
    c4.metric("단일색(점도) 부적합(표시용)", f"{int((single_view.get('_judge')=='부적합').sum()):,}")

    st.divider()

    # ---- Spec Editor
    with st.expander("🛠️ (관리자) 단일색 점도 스펙(하한/상한) 수정", expanded=False):
        st.caption("대시보드에서 **Spec_Single_H&S**의 점도 하한/상한을 직접 수정할 수 있습니다. 저장하면 즉시 그래프에 반영됩니다.")
        spec_show_cols = []
        for c in ["색상군", "제품코드"]:
            if c in spec_single.columns:
                spec_show_cols.append(c)
        bt_col = find_col(spec_single, ["BinderType", "BinderType(자동)", "바인더타입", "binder type"])
        if bt_col:
            spec_show_cols.append(bt_col)
        for c in ["하한", "상한"]:
            if c in spec_single.columns:
                spec_show_cols.append(c)

        spec_df = spec_single[spec_show_cols].copy() if spec_show_cols else spec_single.copy()
        # 표준 컬럼명으로 표시(편집용)
        if bt_col and bt_col in spec_df.columns:
            spec_df = spec_df.rename(columns={bt_col: "BinderType"})

        edited = st.data_editor(
            spec_df,
            use_container_width=True,
            num_rows="dynamic",
            key="spec_editor",
            hide_index=True
        )

        col_save1, col_save2 = st.columns([1, 5])
        with col_save1:
            if st.button("스펙 저장", type="primary"):
                try:
                    updated = update_spec_single_bounds(xlsx_path, edited)
                    st.success(f"스펙 저장 완료: {updated}행 업데이트")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"스펙 저장 실패: {e}")

        with col_save2:
            st.info("※ 스펙 저장은 **하한/상한만** 업데이트합니다. (색상군/제품코드/BinderType 기준)")

    st.divider()

    # ---- 1) List
    st.subheader("1) 단일색 데이터 목록 (엑셀형 보기)")
    need_cols = ["_date", "_lot_calc", "_visc"]
    if any(c not in single_view.columns for c in need_cols):
        st.warning("단일색 데이터에서 표시용 파생 컬럼 생성에 실패했습니다. (시트 구조 확인 필요)")
        st.caption("현재 컬럼: " + ", ".join(list(single_view.columns)[:60]))
    else:
        df_list = single_view.copy()
        df_list["_date"] = pd.to_datetime(df_list["_date"], errors="coerce")
        dmin, dmax = safe_date_bounds(df_list["_date"])

        f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
        with f1:
            start = st.date_input("시작일(목록)", value=max(dmin, dmax - dt.timedelta(days=90)), key=f"list_start_{single_ver}")
        with f2:
            end = st.date_input("종료일(목록)", value=dmax, key=f"list_end_{single_ver}")
        with f3:
            cg_col = find_col(df_list, ["색상군"])
            cg_opts = sorted([x for x in df_list[cg_col].dropna().unique().tolist()]) if cg_col else []
            cg = st.multiselect("색상군(목록)", cg_opts, key=f"list_cg_{single_ver}")
        with f4:
            pc_col = find_col(df_list, ["제품코드"])
            pc_opts = sorted([x for x in df_list[pc_col].dropna().unique().tolist()]) if pc_col else []
            pc = st.multiselect("제품코드(목록)", pc_opts, key=f"list_pc_{single_ver}")

        if start > end:
            start, end = end, start

        df_list = df_list[(df_list["_date"].dt.date >= start) & (df_list["_date"].dt.date <= end)]
        if cg and cg_col:
            df_list = df_list[df_list[cg_col].isin(cg)]
        if pc and pc_col:
            df_list = df_list[df_list[pc_col].isin(pc)]

        blot_col = find_col(df_list, ["사용된 바인더 Lot", "바인더 Lot"])
        view = pd.DataFrame({
            "제조일자": df_list["_date"].dt.date,
            "색상군": df_list[cg_col] if cg_col else None,
            "제품코드": df_list[pc_col] if pc_col else None,
            "단일색Lot(표시용)": df_list["_lot_calc"],
            "사용된바인더": df_list[blot_col] if blot_col else None,
            "BinderType(표시용)": df_list["_binder_type_calc"],
            "점도(cP)": pd.to_numeric(df_list["_visc"], errors="coerce"),
            "점도하한(표시용)": df_list["_spec_lo"],
            "점도상한(표시용)": df_list["_spec_hi"],
            "점도판정(표시용)": df_list["_judge"],
            "색차(ΔE76)": df_list["_ΔE76"],
        }).sort_values(by="제조일자", ascending=False)

        st.dataframe(view, use_container_width=True, height=320)

        st.divider()
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
            mean_df["평균점도표시"] = mean_df["평균점도(cP)"].round(0).astype("Int64").astype(str)
            base = alt.Chart(mean_df).encode(
                x=alt.X("색상군:N", sort=sorted(mean_df["색상군"].unique().tolist()), title="색상군"),
                y=alt.Y("평균점도(cP):Q", title="평균 점도(cP)"),
                tooltip=["색상군:N", "평균점도(cP):Q"]
            )
            points = base.mark_circle(size=260)
            labels = base.mark_text(dx=10, dy=-10).encode(text="평균점도표시:N")
            st.altair_chart((points + labels).interactive(), use_container_width=True)

    st.divider()

    # ---- 2) Trend
    st.subheader("2) 단일색 점도 변화 추이 (Lot별)")
    st.caption("Lot별 입고일 기준 점도 변화를 확인합니다. (점 크게 + 라벨 + 스펙 빨간선)")

    # required
    cg_col = find_col(single_view, ["색상군"])
    pc_col = find_col(single_view, ["제품코드"])
    blot_col = find_col(single_view, ["사용된 바인더 Lot", "바인더 Lot"])

    df = single_view.copy()
    df = df.dropna(subset=["_date", "_visc"])
    df = df[df["_lot_calc"].astype(str).str.strip() != ""]

    if len(df) == 0:
        st.info("입고일/점도/Lot 값이 비어있어 추이 그래프를 표시할 수 없습니다. (엑셀 수식 결과가 비어도 앱이 재계산해야 정상인데, 현재는 원천 데이터가 부족합니다.)")
        with st.expander("🔎 진단(컬럼 확인)"):
            st.write("단일색 컬럼:", list(single_view.columns))
    else:
        dmin, dmax = safe_date_bounds(pd.to_datetime(df["_date"], errors="coerce"))

        f1, f2, f3, f4, f5 = st.columns([1.2, 1.2, 1.6, 2.0, 1.0])
        with f1:
            start = st.date_input("시작일(추이)", value=max(dmin, dmax - dt.timedelta(days=90)), key=f"trend_start_{single_ver}")
        with f2:
            end = st.date_input("종료일(추이)", value=dmax, key=f"trend_end_{single_ver}")
        with f3:
            cg_opts = sorted([x for x in df[cg_col].dropna().unique().tolist()]) if cg_col else []
            cg = st.multiselect("색상군(추이)", cg_opts, key=f"trend_cg_{single_ver}")
        with f4:
            pc_opts = sorted([x for x in df[pc_col].dropna().unique().tolist()]) if pc_col else []
            pc = st.multiselect("제품코드(추이)", pc_opts, key=f"trend_pc_{single_ver}")
        with f5:
            show_labels = st.checkbox("라벨 표시", value=True, key=f"trend_labels_{single_ver}")

        if start > end:
            start, end = end, start

        df = df[(pd.to_datetime(df["_date"], errors="coerce").dt.date >= start) & (pd.to_datetime(df["_date"], errors="coerce").dt.date <= end)]
        if cg and cg_col:
            df = df[df[cg_col].isin(cg)]
        if pc and pc_col:
            df = df[df[pc_col].isin(pc)]

        lot_list = sorted(df["_lot_calc"].dropna().unique().tolist())
        default_pick = lot_list[-5:] if len(lot_list) > 5 else lot_list
        pick = st.multiselect("표시할 단일색 Lot(복수 선택)", lot_list, default=default_pick, key=f"trend_lots_{single_ver}")
        if pick:
            df = df[df["_lot_calc"].isin(pick)]

        if len(df) == 0:
            st.info("선택한 조건에 해당하는 데이터가 없습니다. (기간/색상군/제품코드/로트 필터 확인)")
        else:
            df = df.copy()
            df["_date_ts"] = pd.to_datetime(df["_date"], errors="coerce")
            df = df.sort_values("_date_ts")
            df["점도표시"] = pd.to_numeric(df["_visc"], errors="coerce").round(0).astype("Int64").astype(str)

            # 스펙 선(빨간) — 필터된 데이터에서 대표값 결정
            lo_vals = pd.to_numeric(df["_spec_lo"], errors="coerce").dropna().unique().tolist()
            hi_vals = pd.to_numeric(df["_spec_hi"], errors="coerce").dropna().unique().tolist()
            spec_mode = None
            spec_lo = None
            spec_hi = None
            if len(lo_vals) == 1 and len(hi_vals) == 1:
                spec_mode = "unique"
                spec_lo = float(lo_vals[0])
                spec_hi = float(hi_vals[0])
            elif len(lo_vals) + len(hi_vals) > 0:
                # 여러 스펙이 섞이면 대표값을 선택
                spec_mode = st.radio(
                    "스펙 빨간선 표시 방식",
                    ["표시 안함", "대표값(최소하한/최대상한)"],
                    horizontal=True,
                    index=1,
                    key=f"spec_mode_{single_ver}"
                )
                if spec_mode == "대표값(최소하한/최대상한)":
                    spec_lo = float(min(lo_vals)) if len(lo_vals) else None
                    spec_hi = float(max(hi_vals)) if len(hi_vals) else None

            tooltip_cols = ["_date_ts:T", "_lot_calc:N", "_visc:Q"]
            if pc_col:
                tooltip_cols.insert(2, f"{pc_col}:N")
            if cg_col:
                tooltip_cols.insert(3, f"{cg_col}:N")
            if blot_col:
                tooltip_cols.append(f"{blot_col}:N")
            tooltip_cols += ["_judge:N"]

            base = alt.Chart(df).encode(
                x=alt.X("_date_ts:T", title="입고일"),
                y=alt.Y("_visc:Q", title="점도(cP)"),
                tooltip=tooltip_cols
            )

            line = base.mark_line()
            points = base.mark_point(size=260).encode(color=alt.Color("_lot_calc:N", title="Lot"))

            layers = [line, points]

            if show_labels:
                labels = base.mark_text(dy=-12).encode(
                    color=alt.Color("_lot_calc:N", legend=None),
                    text="점도표시:N"
                )
                layers.append(labels)

            # spec red lines
            if spec_lo is not None:
                lo_df = pd.DataFrame({"y": [spec_lo], "label": [f"Spec Lower: {spec_lo:,.0f}"]})
                layers.append(
                    alt.Chart(lo_df).mark_rule(color="red").encode(y="y:Q")
                )
                layers.append(
                    alt.Chart(lo_df).mark_text(color="red", align="left", dx=6, dy=-6).encode(
                        y="y:Q",
                        text="label:N"
                    )
                )
            if spec_hi is not None:
                hi_df = pd.DataFrame({"y": [spec_hi], "label": [f"Spec Upper: {spec_hi:,.0f}"]})
                layers.append(
                    alt.Chart(hi_df).mark_rule(color="red").encode(y="y:Q")
                )
                layers.append(
                    alt.Chart(hi_df).mark_text(color="red", align="left", dx=6, dy=-6).encode(
                        y="y:Q",
                        text="label:N"
                    )
                )

            st.altair_chart(alt.layer(*layers).interactive(), use_container_width=True)

    st.divider()
    st.subheader("최근 20건 (단일색, 표시용 Lot/스펙 포함)")
    show = single_view.copy()
    show["_date_ts"] = pd.to_datetime(show.get("_date", None), errors="coerce")
    show = show.sort_values(by="_date_ts", ascending=False)
    show_cols = []
    # 원본 주요 컬럼
    for c in ["입고일", "잉크타입 (HEMA/Silicone)", "색상군", "제품코드", "사용된 바인더 Lot", "바인더제조처 (내부/외주)", "점도측정값(cP)"]:
        cc = find_col(show, [c])
        if cc:
            show_cols.append(cc)
    # 표시용 파생
    show_cols += ["_lot_calc", "_binder_type_calc", "_spec_lo", "_spec_hi", "_judge", "_ΔE76"]
    st.dataframe(show[show_cols].head(20), use_container_width=True)

# =========================================================
# 잉크 입고 (단일색 입력)
# =========================================================
with tab_ink_in:
    st.subheader("단일색 잉크 입고 입력")
    st.info("이 탭은 **단일색_수입검사** 시트에 행을 추가(Append)하여 누적합니다. (엑셀 수식 기반 컬럼은 앱에서 재계산하므로, 기존 데이터가 Streamlit에서 '사라지지' 않습니다.)")

    # options
    ink_types = ["HEMA", "Silicone"]
    cg_col_spec = find_col(spec_single, ["색상군"])
    pc_col_spec = find_col(spec_single, ["제품코드"])
    color_groups = sorted(spec_single[cg_col_spec].dropna().unique().tolist()) if cg_col_spec else sorted(COLOR_CODE.keys())
    product_codes = sorted(spec_single[pc_col_spec].dropna().unique().tolist()) if pc_col_spec else []

    # binder lot options: binder_view의 표시용 lot 사용
    binder_lots = binder_view.get("_lot_calc", pd.Series(dtype=object)).dropna().astype(str).tolist()
    binder_lots = sorted(set([x.strip() for x in binder_lots if x.strip()]), reverse=True)

    with st.form("single_form", clear_on_submit=True):
        col1, col2, col3, col4 = st.columns([1.2, 1.3, 1.5, 2.0])
        with col1:
            in_date = st.date_input("입고일", value=dt.date.today(), key="single_in_date")
            ink_type = st.selectbox("잉크타입", ink_types, key="single_ink_type")
            color_group = st.selectbox("색상군", color_groups, key="single_cg")
        with col2:
            product_code = st.selectbox("제품코드", product_codes, key="single_pc") if product_codes else st.text_input("제품코드(직접입력)", value="", key="single_pc_text")
            binder_lot = st.selectbox("사용된 바인더 Lot", binder_lots, key="single_blot") if binder_lots else st.text_input("사용된 바인더 Lot(직접입력)", value="", key="single_blot_text")
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
        # normalize direct inputs
        product_code = product_code.strip() if isinstance(product_code, str) else str(product_code).strip()
        binder_lot = binder_lot.strip() if isinstance(binder_lot, str) else str(binder_lot).strip()

        binder_type = infer_binder_type_from_lot(spec_binder, binder_lot)

        lo, hi = compute_single_spec_row(spec_single, color_group, product_code, binder_type)
        visc_judge = judge_range(visc_meas, lo, hi)

        # lot generation (앱 기준)
        # - 기존 df(표시용 lot 포함)에서 패턴별 seq 이어붙임
        prefix = generate_single_lot_prefix(product_code, color_group, in_date)
        if not prefix:
            st.error("단일색 Lot 자동 생성에 실패했습니다. (색상군/제품코드/날짜 확인)")
        else:
            exist_lots = single_view.get("_lot_calc", pd.Series(dtype=object)).dropna().astype(str)
            # prefix 기반 최대 seq
            seqs = []
            for v in exist_lots.tolist():
                sv = str(v).strip()
                if sv.startswith(prefix):
                    rest = sv[len(prefix):]
                    m = re.match(r"^(\d{2,})$", rest)
                    if m:
                        try:
                            seqs.append(int(m.group(1)))
                        except Exception:
                            pass
            seq = (max(seqs) + 1) if seqs else 1
            new_lot = f"{prefix}{seq:02d}"

            # ΔE -> note
            note2 = note
            if lab_enabled:
                base_hit = base_lab[base_lab.get("제품코드", pd.Series(dtype=str)).astype(str).str.strip() == str(product_code).strip()]
                if len(base_hit) == 1 and all(c in base_hit.columns for c in ["기준_L*", "기준_a*", "기준_b*"]):
                    base = (
                        safe_to_float(base_hit.iloc[0]["기준_L*"]),
                        safe_to_float(base_hit.iloc[0]["기준_a*"]),
                        safe_to_float(base_hit.iloc[0]["기준_b*"]),
                    )
                    if None not in base:
                        de = delta_e76((float(L), float(a), float(b)), base)
                        note2 = (note2 + " " if note2 else "") + f"[ΔE76={de:.2f}]"
                    else:
                        note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"
                else:
                    note2 = (note2 + " " if note2 else "") + f"[Lab=({L:.2f},{a:.2f},{b:.2f})]"

            # 엑셀 append (수식 캐시 유실 영향 최소화를 위해, 수식열도 '값'으로 채움)
            row_by_norm = {
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
                append_row_to_sheet_by_norm(xlsx_path, SHEET_SINGLE, row_by_norm)
                st.success(f"저장 완료! 단일색 Lot = {new_lot} / 점도판정 = {visc_judge}")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

# =========================================================
# 바인더 입출고 (입력/반품/구글시트 보기)
# =========================================================
with tab_binder:
    st.subheader("업체반환(반품) 입력 (kg 단위)")
    st.caption("※ 20kg(1통) 기준이더라도, 실제 반환량은 kg 단위로 입력합니다. (재고요약은 제거됨)")

    binder_names = sorted(spec_binder.get("바인더명", pd.Series(dtype=object)).dropna().unique().tolist())
    binder_lots = binder_view.get("_lot_calc", pd.Series(dtype=object)).dropna().astype(str).tolist()
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
            row_by_norm = {
                "일자": r_date,
                "바인더타입": r_type,
                "바인더명": r_name,
                "바인더 Lot": final_lot,
                "반환량(kg)": float(r_kg),
                "비고": r_note,
            }
            try:
                append_row_to_sheet_by_norm(xlsx_path, SHEET_BINDER_RETURN, row_by_norm)
                st.success("반품 저장 완료!")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"반품 저장 실패: {e}")

    st.divider()

    st.subheader("바인더 입력 (제조/입고) — 여러 Lot/날짜 일괄 입력 지원")
    st.caption("※ 바인더는 여러 날짜의 Lot가 한 번에 입고될 수 있어, 날짜별/수량별로 묶음 입력을 지원합니다.")

    input_mode = st.radio("입력 방식", ["개별 입력", "묶음 입력(여러 날짜/수량)"], horizontal=True, key="binder_input_mode")

    # 기존 바인더 lot 계산용 existing
    existing_binder_lots = binder_view.get("_lot_calc", pd.Series(dtype=str)).dropna().astype(str)

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
            lot = generate_binder_lot(spec_binder, b_name, mfg_date, existing_binder_lots)
            judge_v = judge_range(visc, visc_lo, visc_hi)
            judge_u = judge_range(uv if uv_enabled else None, None, uv_hi)
            judge = "부적합" if (judge_v == "부적합" or judge_u == "부적합") else "적합"

            row_by_norm = {
                "제조/입고일": mfg_date,
                "바인더명": b_name,
                "Lot(자동)": lot,
                "점도(cP)": float(visc),
                "UV흡광도(선택)": float(uv) if uv_enabled else None,
                "판정": judge,
                "비고": note,
            }
            try:
                append_row_to_sheet_by_norm(xlsx_path, SHEET_BINDER, row_by_norm)
                st.success(f"저장 완료! 바인더 Lot = {lot}")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    else:
        st.caption("아래 표에 날짜/바인더명/수량(통)/점도/UV/비고를 입력하고, 한 번에 저장하세요.")
        base_rows = st.session_state.get("binder_batch_rows")
        if base_rows is None:
            base_rows = [
                {"제조/입고일": dt.date.today(), "바인더명": binder_names[0] if binder_names else "", "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
                {"제조/입고일": dt.date.today() - dt.timedelta(days=1), "바인더명": binder_names[0] if binder_names else "", "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
                {"제조/입고일": dt.date.today() - dt.timedelta(days=2), "바인더명": binder_names[0] if binder_names else "", "수량(통)": 8, "점도(cP)": 0.0, "UV입력": False, "UV흡광도(선택)": None, "비고": ""},
            ]
        edit_df = pd.DataFrame(base_rows)
        edit_df = st.data_editor(edit_df, use_container_width=True, num_rows="dynamic", key="binder_batch_editor")
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

            existing_list = existing_binder_lots.dropna().astype(str).tolist()
            seq_counters = {}
            rows_out, preview_out = [], []
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

                    row_by_norm = {
                        "제조/입고일": mfg_date,
                        "바인더명": b_name,
                        "Lot(자동)": lot,
                        "점도(cP)": float(visc) if visc is not None else None,
                        "UV흡광도(선택)": float(uv_val) if uv_enabled and uv_val is not None else None,
                        "판정": judge,
                        "비고": note,
                    }
                    rows_out.append(row_by_norm)
                    preview_out.append(row_by_norm)
                    existing_list.append(lot)

            st.write("저장 미리보기(상위 50건)")
            st.dataframe(pd.DataFrame(preview_out).tail(50), use_container_width=True)

            try:
                append_rows_to_sheet_by_norm(xlsx_path, SHEET_BINDER, rows_out)
                st.success(f"묶음 저장 완료! 총 {len(rows_out)}건 입력했습니다.")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"저장 실패: {e}")

    st.divider()
    st.subheader("바인더 입출고 (Google Sheets 자동 반영, 최신순)")
    st.caption("구글 시트를 수정하면 이 화면은 새로고침 시 자동 반영됩니다. (캐시 60초)")

    def detect_date_col(df: pd.DataFrame):
        candidates = []
        for c in df.columns:
            ck = norm_key(c)
            if any(k in ck for k in ["일자", "날짜", "date", "입고일", "출고일"]):
                candidates.append(c)
        return candidates[0] if candidates else None

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

# =========================================================
# Search (placeholder)
# =========================================================
with tab_search:
    st.info("빠른검색은 필요하시면 조건(기간/제품/색상군/바인더Lot/판정)까지 확장해드릴게요.")
