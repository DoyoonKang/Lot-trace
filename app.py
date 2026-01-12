
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
@st.cache_data(ttl=60, show_spinner=False)  # 60초마다 최신값으로 갱신
def read_gsheet_csv(sheet_id: str, sheet_name: str) -> pd.DataFrame:
    """
    Public/Link-shared Google Sheet 를 CSV로 읽어옵니다.
    (sheet_name이 한글이어도 requests params가 자동 인코딩)
    """
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
def _read_excel_from_path(xlsx_path: str) -> dict[str, pd.DataFrame]:
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


def append_row_to_sheet(xlsx_path: str, sheet_name: str, row: dict):
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
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


# =========================
# UI Header
# =========================
st.title("액상 잉크 Lot 추적 관리 대시보드")
st.caption("✅ 빠른 검색 + ✅ 신규 입력(엑셀에 누적) + ✅ 기본 대시보드 + ✅ 바인더 입출고(구글시트 자동 반영)")


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

data = load_data(xlsx_path)
binder_df = data["binder"].copy()
single_df = data["single"].copy()
spec_binder = data["spec_binder"].copy()
spec_single = data["spec_single"].copy()
base_lab = data["base_lab"].copy()

if "제조/입고일" in binder_df.columns:
    binder_df["제조/입고일"] = binder_df["제조/입고일"].apply(normalize_date)
if "입고일" in single_df.columns:
    single_df["입고일"] = single_df["입고일"].apply(normalize_date)


# =========================
# Tabs (요청 순서)
# =========================
tab_dash, tab_input, tab_binder, tab_search = st.tabs(
    ["📊 대시보드", "✍️ 신규입력", "📦 바인더 입출고", "🔎 빠른검색"]
)


# =========================
# Dashboard
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

    left, right = st.columns([2, 1])

st.subheader("점도 변화 추이 (로트별)")
st.caption("데이터를 입력(저장)할 때마다 최신 데이터로 자동 반영됩니다.")

mode2 = st.radio("데이터 선택", ["단일색(수입검사) 점도", "바인더(제조/입고) 점도"], horizontal=True)

if mode2 == "단일색(수입검사) 점도":
    df = single_df.copy()

    # 필수 컬럼 체크
    need_cols = ["입고일", "단일색잉크 Lot", "점도측정값(cP)"]
    miss = [c for c in need_cols if c not in df.columns]
    if miss:
        st.warning(f"단일색 데이터에 필요한 컬럼이 없습니다: {miss}")
    else:
        df = df.dropna(subset=["입고일", "단일색잉크 Lot", "점도측정값(cP)"])
        df["입고일"] = pd.to_datetime(df["입고일"])

        # ---- 필터 UI
        f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.6, 2.0])
        with f1:
            dmin = df["입고일"].min().date()
            dmax = df["입고일"].max().date()
            start = st.date_input("시작일", value=max(dmin, dmax - dt.timedelta(days=90)))
        with f2:
            end = st.date_input("종료일", value=dmax)
        with f3:
            if "색상군" in df.columns:
                cg = st.multiselect("색상군", sorted(df["색상군"].dropna().unique().tolist()))
            else:
                cg = []
        with f4:
            if "제품코드" in df.columns:
                pc = st.multiselect("제품코드", sorted(df["제품코드"].dropna().unique().tolist()))
            else:
                pc = []

        df = df[(df["입고일"].dt.date >= start) & (df["입고일"].dt.date <= end)]
        if cg and "색상군" in df.columns:
            df = df[df["색상군"].isin(cg)]
        if pc and "제품코드" in df.columns:
            df = df[df["제품코드"].isin(pc)]

        # 로트 선택(너무 많으면 보기 힘드니까 선택형)
        lot_list = sorted(df["단일색잉크 Lot"].astype(str).unique().tolist())
        pick = st.multiselect("표시할 단일색 Lot(복수 선택)", lot_list, default=lot_list[-5:] if len(lot_list) > 5 else lot_list)
        if pick:
            df = df[df["단일색잉크 Lot"].astype(str).isin(pick)]

        # ---- 차트
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

        st.caption("※ 로트가 많으면 화면이 복잡해집니다. 상단에서 로트를 몇 개만 선택해서 보는 걸 추천합니다.")

else:
    df = binder_df.copy()

    need_cols = ["제조/입고일", "Lot(자동)", "점도(cP)"]
    miss = [c for c in need_cols if c not in df.columns]
    if miss:
        st.warning(f"바인더 데이터에 필요한 컬럼이 없습니다: {miss}")
    else:
        df = df.dropna(subset=["제조/입고일", "Lot(자동)", "점도(cP)"])
        df["제조/입고일"] = pd.to_datetime(df["제조/입고일"])

        # 필터
        f1, f2, f3 = st.columns([1.2, 1.2, 2.6])
        with f1:
            dmin = df["제조/입고일"].min().date()
            dmax = df["제조/입고일"].max().date()
            start = st.date_input("시작일(바인더)", value=max(dmin, dmax - dt.timedelta(days=180)))
        with f2:
            end = st.date_input("종료일(바인더)", value=dmax)
        with f3:
            lots = sorted(df["Lot(자동)"].astype(str).unique().tolist())
            pick = st.multiselect("표시할 바인더 Lot(복수 선택)", lots, default=lots[-10:] if len(lots) > 10 else lots)

        df = df[(df["제조/입고일"].dt.date >= start) & (df["제조/입고일"].dt.date <= end)]
        if pick:
            df = df[df["Lot(자동)"].astype(str).isin(pick)]

        chart = (
            alt.Chart(df)
            .mark_line(point=True)
            .encode(
                x=alt.X("제조/입고일:T", title="제조/입고일"),
                y=alt.Y("점도(cP):Q", title="점도(cP)"),
                color=alt.Color("Lot(자동):N", title="Binder Lot"),
                tooltip=["제조/입고일:T", "바인더명:N", "Lot(자동):N", "점도(cP):Q", "판정:N"],
            )
            .interactive()
        )
        st.altair_chart(chart, use_container_width=True)


    
    with left:
        st.subheader("단일색 점도 평균 (색상군별)")
        if "색상군" in single_df.columns and "점도측정값(cP)" in single_df.columns:
            chart_df = single_df[["색상군", "점도측정값(cP)"]].dropna()
            st.bar_chart(chart_df.groupby("색상군")["점도측정값(cP)"].mean())
        else:
            st.info("단일색 데이터에 '색상군' 또는 '점도측정값(cP)' 컬럼이 없습니다.")

    with right:
        st.subheader("최근 20건")
        show = single_df.sort_values(by="입고일", ascending=False).head(20) if "입고일" in single_df.columns else single_df.head(20)
        st.dataframe(show, use_container_width=True)


# =========================
# Input
# =========================
with tab_input:
    st.info("이 탭은 **엑셀 파일에 행을 추가(Append)** 해서 데이터가 누적되도록 만들었습니다. (여러 사람이 동시에 쓰면 충돌 가능)")
    sub_b, sub_s = st.tabs(["바인더 입력", "단일색 잉크 입력"])

    # ---- Binder form
    with sub_b:
        binder_names = sorted(spec_binder["바인더명"].dropna().unique().tolist())
        with st.form("binder_form", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                mfg_date = st.date_input("제조/입고일", value=dt.date.today())
                binder_name = st.selectbox("바인더명", binder_names)
            with col2:
                visc = st.number_input("점도(cP)", min_value=0.0, step=1.0, format="%.1f")
                uv = st.number_input("UV흡광도(선택)", min_value=0.0, step=0.01, format="%.3f")
                uv_enabled = st.checkbox("UV 값 입력함", value=False)
            with col3:
                note = st.text_input("비고", value="")
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

    # ---- Single form
    with sub_s:
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
                st.caption("선택: 착색력(L*a*b*) 입력하면, 기준LAB이 있을 경우 ΔE(76)을 자동 계산해서 '비고'에 기록합니다.")
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
# Binder IO (Google Sheets)
# =========================
with tab_binder:
    st.subheader("바인더 입출고 (Google Sheets 자동 반영)")
    st.caption("구글 시트를 수정하면, 이 화면은 새로고침 시 자동으로 최신 값이 반영됩니다. (캐시 60초)")

    try:
        df_hema = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_HEMA)
        df_sil = read_gsheet_csv(BINDER_SHEET_ID, BINDER_SHEET_SIL)
    except Exception as e:
        st.error("구글시트에서 데이터를 못 불러왔어요. 시트 공유/웹게시/시트명/ID를 확인하세요.")
        st.exception(e)
        st.stop()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### HEMA")
        st.dataframe(df_hema, use_container_width=True)
    with c2:
        st.markdown("### Silicon")
        st.dataframe(df_sil, use_container_width=True)

    if st.button("지금 최신값으로 다시 불러오기"):
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
            start = st.date_input("시작일", value=dt.date.today() - dt.timedelta(days=30))
        with d2:
            end = st.date_input("종료일", value=dt.date.today())
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

