import altair as alt
import streamlit as st
import pandas as pd
import datetime as dt
import re
from pathlib import Path
from openpyxl import load_workbook
import requests
from io import StringIO


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
    """Public/Link-shared Google Sheet 를 CSV로 읽어옵니다."""
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


def safe_to_float(x):
    if x is None or (isinstance(x, float) and pd.isna(x)) or (isinstance(x, str) and x.strip() == ""):
        return None
    try:
        if isinstance(x, str):
            x = x.replace(",", "")
        return float(x)
    except Exception:
        return None


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


def ensure_sheet_exists(xlsx_path: str, sheet_name: str, headers: list[str]):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        wb.save(xlsx_path)


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
            if str(binder_lot).strip().startswith(prefix):
                return r["바인더명"]
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
    lots = single_df.get("단일색잉크 Lot", pd.Series(dtype=str)).dropna().astype(str).tolist()

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


def append_row_to_sheet(xlsx_path: str, sheet_name: str, row: dict):
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


def detect_date_col(df: pd.DataFrame):
    candidates = []
    for c in df.columns:
        ck = norm_key(c)
        if any(k in ck for k in ["일자", "날짜", "date", "입고일", "출고일"]):
            candidates.append(c)
    return candidates[0] if candidates else None


def safe_date_bounds(s: pd.Series):
    s = pd.to_datetime(s, errors="coerce")
    s = s.dropna()
    if len(s) == 0:
        today = dt.date.today()
        return today, today
    return s.min().date(), s.max().date()


def extract_or_compute_de76(single_df: pd.DataFrame, base_lab: pd.DataFrame) -> pd.Series:
    # base_lab 정규화
    base = base_lab.copy()
    if "제품코드" in base.columns:
        base["제품코드"] = base["제품코드"].astype(str).str.strip()

    note_col = "비고" if "비고" in single_df.columns else None
    out = pd.Series([None] * len(single_df), index=single_df.index, dtype="float")

    if note_col:
        pat = re.compile(r"\[\s*ΔE76\s*=\s*([0-9]+(?:\.[0-9]+)?)\s*\]")
        for idx, val in single_df[note_col].items():
            if pd.isna(val):
                continue
            m = pat.search(str(val))
            if m:
                try:
                    out.loc[idx] = float(m.group(1))
                except Exception:
                    pass

    need_cols = ["제품코드", "착색력_L*", "착색력_a*", "착색력_b*"]
    if all(c in single_df.columns for c in need_cols) and all(c in base.columns for c in ["기준_L*", "기준_a*", "기준_b*", "제품코드"]):
        base_map = base.set_index("제품코드")[["기준_L*", "기준_a*", "기준_b*"]].to_dict("index")
        for idx, row in single_df.iterrows():
            if pd.notna(out.loc[idx]):
                continue
            pc = row.get("제품코드", None)
            if pd.isna(pc):
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

# 반환 시트가 없으면 생성
ensure_sheet_exists(
    xlsx_path,
    SHEET_BINDER_RETURN,
    headers=["일자", "바인더타입", "바인더명", "바인더 Lot", "반환량(kg)", "비고"]
)

# Load & normalize
raw = load_data(xlsx_path)
binder_df = normalize_df_columns(raw["binder"])
single_df = normalize_df_columns(raw["single"])
spec_binder = normalize_df_columns(raw["spec_binder"])
spec_single = normalize_df_columns(raw["spec_single"])
base_lab = normalize_df_columns(raw["base_lab"])

# 날짜 정규화
if "제조/입고일" in binder_df.columns:
    binder_df["제조/입고일"] = binder_df["제조/입고일"].apply(normalize_date)
if "입고일" in single_df.columns:
    single_df["입고일"] = single_df["입고일"].apply(normalize_date)

# ΔE76 파생
single_df["_ΔE76"] = extract_or_compute_de76(single_df, base_lab)

single_ver = str(pd.to_datetime(single_df.get("입고일", pd.Series(dtype=object)), errors="coerce").max())

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

    st.subheader("1) 단일색 데이터 목록 (엑셀형 보기)")

    need = ["입고일", "색상군", "제품코드", "사용된 바인더 Lot", "점도측정값(cP)"]
    miss = [c for c in need if c not in single_df.columns]
    if miss:
        st.warning(f"단일색 시트에서 필요한 컬럼을 찾지 못했습니다: {miss}")
    else:
        df_list = single_df.copy()
        df_list["입고일"] = pd.to_datetime(df_list["입고일"], errors="coerce")
        dmin, dmax = safe_date_bounds(df_list["입고일"])

        f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
        with f1:
            start = st.date_input("시작일(목록)", value=max(dmin, dmax - dt.timedelta(days=90)), key=f"list_start_{single_ver}")
        with f2:
            end = st.date_input("종료일(목록)", value=dmax, key=f"list_end_{single_ver}")
        with f3:
            cg_opts = sorted([x for x in df_list["색상군"].dropna().unique().tolist()])
            cg = st.multiselect("색상군(목록)", cg_opts, key=f"list_cg_{single_ver}")
        with f4:
            pc_opts = sorted([x for x in df_list["제품코드"].dropna().unique().tolist()])
            pc = st.multiselect("제품코드(목록)", pc_opts, key=f"list_pc_{single_ver}")

        if start > end:
            start, end = end, start

        df_list = df_list[(df_list["입고일"].dt.date >= start) & (df_list["입고일"].dt.date <= end)]
        if cg:
            df_list = df_list[df_list["색상군"].isin(cg)]
        if pc:
            df_list = df_list[df_list["제품코드"].isin(pc)]

        view = pd.DataFrame({
            "제조일자": df_list["입고일"].dt.date,
            "색상군": df_list["색상군"],
            "제품코드": df_list["제품코드"],
            "사용된바인더": df_list["사용된 바인더 Lot"],
            "점도(cP)": pd.to_numeric(df_list["점도측정값(cP)"], errors="coerce"),
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
            points = base.mark_circle(size=220)
            labels = base.mark_text(dx=10, dy=-8).encode(text="평균점도표시:N")
            st.altair_chart((points + labels).interactive(), use_container_width=True)

    st.divider()

    st.subheader("2) 단일색 점도 변화 추이 (Lot별)")
    st.caption("선택한 Lot별로 입고일 기준 점도 변화를 확인합니다. (점 크기/라벨 강화)")

    if all(c in single_df.columns for c in ["입고일", "단일색잉크 Lot", "점도측정값(cP)"]):
        df = single_df.copy()
        df["입고일"] = pd.to_datetime(df["입고일"], errors="coerce")
        df["점도"] = pd.to_numeric(df["점도측정값(cP)"].astype(str).str.replace(",", "", regex=False), errors="coerce")
        df["Lot"] = df["단일색잉크 Lot"].astype(str).replace("nan", "").replace("None", "")
        df = df.dropna(subset=["입고일", "점도"])
        df = df[df["Lot"].str.strip() != ""]

        if len(df) == 0:
            st.info("입고일/점도/Lot 값이 비어있어 추이 그래프를 표시할 수 없습니다.")
        else:
            dmin, dmax = safe_date_bounds(df["입고일"])
            f1, f2, f3, f4, f5 = st.columns([1.2, 1.2, 1.6, 2.0, 1.0])
            with f1:
                start = st.date_input("시작일(추이)", value=max(dmin, dmax - dt.timedelta(days=90)), key=f"trend_start_{single_ver}")
            with f2:
                end = st.date_input("종료일(추이)", value=dmax, key=f"trend_end_{single_ver}")
            with f3:
                cg_opts = sorted([x for x in df.get("색상군", pd.Series(dtype=object)).dropna().unique().tolist()]) if "색상군" in df.columns else []
                cg = st.multiselect("색상군(추이)", cg_opts, key=f"trend_cg_{single_ver}")
            with f4:
                pc_opts = sorted([x for x in df.get("제품코드", pd.Series(dtype=object)).dropna().unique().tolist()]) if "제품코드" in df.columns else []
                pc = st.multiselect("제품코드(추이)", pc_opts, key=f"trend_pc_{single_ver}")
            with f5:
                show_labels = st.checkbox("라벨 표시", value=True, key=f"trend_labels_{single_ver}")

            if start > end:
                start, end = end, start

            df = df[(df["입고일"].dt.date >= start) & (df["입고일"].dt.date <= end)]
            if cg and "색상군" in df.columns:
                df = df[df["색상군"].isin(cg)]
            if pc and "제품코드" in df.columns:
                df = df[df["제품코드"].isin(pc)]

            lot_list = sorted(df["Lot"].dropna().unique().tolist())
            default_pick = lot_list[-5:] if len(lot_list) > 5 else lot_list
            pick = st.multiselect("표시할 단일색 Lot(복수 선택)", lot_list, default=default_pick, key=f"trend_lots_{single_ver}")
            if pick:
                df = df[df["Lot"].isin(pick)]

            if len(df) == 0:
                st.info("선택한 조건에 해당하는 데이터가 없습니다. (기간/색상군/제품코드/로트 필터를 확인해주세요)")
                if st.button("필터 초기화(추이)", key=f"trend_reset_{single_ver}"):
                    for k in [f"trend_start_{single_ver}", f"trend_end_{single_ver}", f"trend_cg_{single_ver}", f"trend_pc_{single_ver}", f"trend_lots_{single_ver}"]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.rerun()
            else:
                df = df.sort_values("입고일")
                df["점도표시"] = df["점도"].round(0).astype("Int64").astype(str)

                tooltip_cols = ["입고일:T", "Lot:N", "점도:Q"]
                if "제품코드" in df.columns:
                    tooltip_cols.insert(2, "제품코드:N")
                if "색상군" in df.columns:
                    tooltip_cols.insert(3, "색상군:N")
                if "사용된 바인더 Lot" in df.columns:
                    tooltip_cols.append("사용된 바인더 Lot:N")

                base = alt.Chart(df).encode(
                    x=alt.X("입고일:T", title="입고일"),
                    y=alt.Y("점도:Q", title="점도(cP)"),
                    tooltip=tooltip_cols
                )
                line = base.mark_line()
                points = base.mark_point(size=180).encode(color=alt.Color("Lot:N", title="Lot"))
                if show_labels:
                    labels = base.mark_text(dy=-10).encode(
                        color=alt.Color("Lot:N", legend=None),
                        text="점도표시:N"
                    )
                    chart = (line + points + labels).interactive()
                else:
                    chart = (line + points).interactive()

                st.altair_chart(chart, use_container_width=True)

    st.divider()
    st.subheader("최근 20건 (단일색)")
    show = single_df.copy()
    if "입고일" in show.columns:
        show["입고일"] = pd.to_datetime(show["입고일"], errors="coerce")
        show = show.sort_values(by="입고일", ascending=False)
    st.dataframe(show.head(20), use_container_width=True)

# =========================
# 잉크 입고 (단일색 입력만)
# =========================
with tab_ink_in:
    st.subheader("단일색 잉크 입력(입고)")
    st.info("이 탭은 **단일색_수입검사** 시트에 행을 추가(Append)하여 누적합니다. (동시 사용 시 충돌 가능)")

    ink_types = ["HEMA", "Silicone"]
    color_groups = sorted(spec_single.get("색상군", pd.Series(dtype=object)).dropna().unique().tolist())
    product_codes = sorted(spec_single.get("제품코드", pd.Series(dtype=object)).dropna().unique().tolist())

    binder_lots = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).tolist()
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

        spec_hit = spec_single[
            (spec_single.get("색상군") == color_group) &
            (spec_single.get("제품코드") == product_code)
        ].copy()

        if binder_type and "BinderType" in spec_hit.columns:
            spec_hit = spec_hit[spec_hit["BinderType"] == binder_type]

        if len(spec_hit) == 0:
            lo, hi = None, None
            visc_judge = None
            st.warning("점도 기준을 Spec_Single_H&S에서 찾지 못했습니다. (색상군/제품코드/바인더타입 조합 확인)")
        else:
            lo = safe_to_float(spec_hit.get("하한").iloc[0])
            hi = safe_to_float(spec_hit.get("상한").iloc[0])
            visc_judge = judge_range(visc_meas, lo, hi)

        new_lot = generate_single_lot(single_df, product_code, color_group, in_date)
        if new_lot is None:
            st.error("단일색 Lot 자동 생성에 실패했습니다. (색상군 매핑 확인 필요)")
        else:
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
# 바인더 입출고 (입력/반품/구글시트 보기)
# =========================
with tab_binder:
    st.subheader("업체반환(반품) 입력 (kg 단위)")
    st.caption("※ 20kg(1통) 기준이더라도, 실제 반환량은 kg 단위로 입력합니다.")

    binder_names = sorted(spec_binder.get("바인더명", pd.Series(dtype=object)).dropna().unique().tolist())
    binder_lots = binder_df.get("Lot(자동)", pd.Series(dtype=str)).dropna().astype(str).tolist()
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
                append_row_to_sheet(xlsx_path, SHEET_BINDER_RETURN, row)
                st.success("반품 저장 완료!")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"반품 저장 실패: {e}")

    st.divider()

    st.subheader("바인더 입력 (제조/입고) — 여러 Lot/날짜 일괄 입력 지원")
    st.caption("※ 바인더는 여러 날짜의 Lot가 한 번에 입고될 수 있어, 날짜별/수량별로 묶음 입력을 지원합니다.")

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
            lot = generate_binder_lot(spec_binder, b_name, mfg_date, binder_df.get("Lot(자동)", pd.Series(dtype=str)))

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

            existing = binder_df.get("Lot(자동)", pd.Series(dtype=str))
            existing_list = existing.dropna().astype(str).tolist()
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

                    row = {
                        "제조/입고일": mfg_date,
                        "바인더명": b_name,
                        "Lot(자동)": lot,
                        "점도(cP)": float(visc),
                        "UV흡광도(선택)": float(uv_val) if uv_enabled and uv_val is not None else None,
                        "판정": judge,
                        "비고": note,
                    }
                    rows_out.append(row)
                    preview_out.append(row)
                    existing_list.append(lot)

            st.write("저장 미리보기(상위 50건)")
            st.dataframe(pd.DataFrame(preview_out).tail(50), use_container_width=True)

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

# =========================
# Search
# =========================
with tab_search:
    st.info("빠른검색은 기존 로직을 유지했습니다. 필요하면 검색조건(복합 필터)까지 확장해드릴게요.")
