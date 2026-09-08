from datetime import datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# 기본 설정
# =========================================================
st.set_page_config(
    page_title="실시간 물량집계 시스템",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

C_FILE_PATH = DATA_DIR / "chuko_source.xlsx"
B_FILE_PATH = DATA_DIR / "batcha_source.xlsx"


def get_secret(name: str, fallback: str) -> str:
    try:
        return str(st.secrets[name])
    except Exception:
        return fallback


ADMIN_PASSWORD = get_secret("ADMIN_PASSWORD", "bstk2026!")
GENERAL_PASSWORD = get_secret("GENERAL_PASSWORD", "1010")


# =========================================================
# 세션 상태 초기화
# =========================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "role" not in st.session_state:
    st.session_state.role = None


# =========================================================
# 로그인 화면
# =========================================================
def render_login():
    st.markdown("### 📦 실시간 물량집계 시스템")
    st.caption("출고 및 배차 데이터를 업로드하여 통합 집계합니다.")
    st.markdown("")

    mode = st.radio("접속 모드", ["일반모드", "관리자모드"], horizontal=True)
    password = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")

    if st.button("로그인", use_container_width=True, type="primary"):
        expected = GENERAL_PASSWORD if mode == "일반모드" else ADMIN_PASSWORD
        if password == expected:
            st.session_state.authenticated = True
            st.session_state.role = "general" if mode == "일반모드" else "admin"
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")


if not st.session_state.authenticated:
    render_login()
    st.stop()


# =========================================================
# 핵심 집계 로직 (제미나이 원본 반영)[cite: 1]
# =========================================================
def process_logistics_data(df_c, df_b):
    if "차랑번호" in df_b.columns and "차량번호" not in df_b.columns:
        df_b.rename(columns={"차랑번호": "차량번호"}, inplace=True)

    # 3. 배차 ↔ 출고 피킹지출력 매칭
    df_b["prefix12"] = df_b["일련번호"].astype(str).str[:12]
    df_c["prefix12"] = df_c["출고번호"].astype(str).str[:12]

    df_b["key12"] = (
        df_b["출고처명(거래처)"].astype(str).str.strip() + "_" + df_b["prefix12"]
    )
    df_c["key12"] = (
        df_c["출고처명(거래처)"].astype(str).str.strip() + "_" + df_c["prefix12"]
    )

    df_b["prefix10"] = df_b["일련번호"].astype(str).str[:10]
    df_c["prefix10"] = df_c["출고번호"].astype(str).str[:10]

    df_b["key10"] = (
        df_b["출고처명(거래처)"].astype(str).str.strip() + "_" + df_b["prefix10"]
    )
    df_c["key10"] = (
        df_c["출고처명(거래처)"].astype(str).str.strip() + "_" + df_c["prefix10"]
    )

    map12 = df_c.set_index("key12")["피킹지출력"].to_dict()
    map10 = df_c.set_index("key10")["피킹지출력"].to_dict()

    df_b["피킹지출력"] = df_b["key12"].map(map12)
    df_b["피킹지출력"] = df_b["피킹지출력"].fillna(df_b["key10"].map(map10))

    df_b["key_addr"] = (
        df_b["출고처명(거래처)"].astype(str).str.strip()
        + "_"
        + df_b["배송처주소"].astype(str).str.strip()
    )
    df_c["key_addr"] = (
        df_c["출고처명(거래처)"].astype(str).str.strip()
        + "_"
        + df_c["배송처주소"].astype(str).str.strip()
    )

    map_addr = df_c.groupby("key_addr")["피킹지출력"].first().to_dict()
    df_b["피킹지출력"] = df_b["피킹지출력"].fillna(df_b["key_addr"].map(map_addr))

    temp_cols = ["prefix12", "prefix10", "key12", "key10", "key_addr"]
    df = df_b.drop(columns=[c for c in temp_cols if c in df_b.columns])

    def find_col(keywords):
        for keyword in keywords:
            for col in df.columns:
                if keyword in str(col).strip():
                    return col
        return None

    addr_col = find_col(["배송처주소", "주소", "배송지", "수령지"])
    type_col = find_col(["타이어구분", "구분", "타이어", "분류"])
    qty_col = find_col(["배차수량", "수량", "총수량", "작업수량"])
    car_col = find_col(["차량번호", "차랑번호", "운송사", "차량", "배차"])
    driver_col = find_col(["차량기사", "기사"])
    courier_col = find_col(["택배구분", "택배분류", "택배"])
    code_col = find_col(["품목코드", "코드", "자재코드", "SKU"])
    size_col = find_col(["SIZE", "규격", "사이즈"])
    pttn_col = find_col(["PTTN", "패턴"])
    note_col = find_col(["비고", "특이사항", "메모", "품목비고", "요청사항"])
    print_col = find_col(["피킹지출력", "출력", "피킹지"])
    time_col = find_col(["출고일시", "출고시간", "일시"])

    if not addr_col or not type_col or not qty_col:
        raise RuntimeError("필수 열(주소, 구분, 수량)을 찾을 수 없습니다.")

    df_filtered = df.copy()
    df_filtered[qty_col] = pd.to_numeric(
        df_filtered[qty_col], errors="coerce"
    ).fillna(0)
    df_filtered[type_col] = (
        df_filtered[type_col].fillna("").astype(str).str.strip()
    )

    ITEM_NAME_MAP = {
        "장우산(25개입)": "장우산",
        "개폐식 포스터 액자": "개폐식액자",
        "골프공옐로우": "골프공옐로우",
        "공기압 체크기": "공기압체크기",
        "뎁스게이지": "뎁스게이지",
        "물티슈(10EA)": "물티슈",
        "물티슈(30EA)": "물티슈",
        "미니간판": "미니간판",
        "반팔티셔츠 2XL": "반팔티",
        "반팔티셔츠 L": "반팔티",
        "반팔티셔츠 M": "반팔티",
        "반팔티셔츠 XL": "반팔티",
        "볼펜": "볼펜",
        "부채": "부채",
        "아크릴 미니 B 간판": "아크릴 미니간판",
        "일반 알미늄 액자": "알미늄액자",
        "6.5 OZ 종이컵": "종이컵",
        "장우산(30개입)": "장우산",
        "주차 알림판 피규어": "주차피규어",
        "GOLF BALL": "골프공",
    }

    def map_marketing_category(row):
        t_val = str(row[type_col]).upper()
        if "마케팅" in t_val or "골프공" in t_val:
            size_val = str(row[size_col]).strip() if size_col else ""
            code_val = str(row[code_col]).strip() if code_col else ""
            for key, mapped_name in ITEM_NAME_MAP.items():
                if key in size_val or key in code_val:
                    return mapped_name
            if size_val and size_val.lower() not in ["nan", "none", ""]:
                return size_val
            if code_val and code_val.lower() not in ["nan", "none", ""]:
                return code_val
        return row[type_col]

    df_filtered[type_col] = df_filtered.apply(map_marketing_category, axis=1)

    s_car = (
        df_filtered[car_col].fillna("").astype(str).str.strip()
        if car_col
        else pd.Series("", index=df_filtered.index)
    )
    s_driver = (
        df_filtered[driver_col].fillna("").astype(str).str.strip()
        if driver_col
        else pd.Series("", index=df_filtered.index)
    )
    s_courier_type = (
        df_filtered[courier_col].fillna("").astype(str).str.strip()
        if courier_col
        else pd.Series("", index=df_filtered.index)
    )

    df_filtered["_차량기사_str"] = s_car + " / " + s_driver

    if time_col:
        df_filtered["_출고일시_parsed"] = pd.to_datetime(
            df_filtered[time_col], errors="coerce"
        )
        df_filtered["_출고일시_str"] = (
            df_filtered[time_col].fillna("").astype(str).str.strip()
        )
    else:
        df_filtered["_출고일시_parsed"] = pd.NaT
        df_filtered["_출고일시_str"] = ""

    if print_col:
        df_filtered["출력"] = np.where(
            df_filtered[print_col].astype(str).str.upper().str.contains("Y"),
            "○",
            "",
        )
    else:
        df_filtered["출력"] = ""

    s_addr = df_filtered[addr_col].fillna("").astype(str).str.strip()
    s_note = (
        df_filtered[note_col].fillna("").astype(str).str.strip()
        if note_col
        else pd.Series("", index=df_filtered.index)
    )
    s_code = (
        df_filtered[code_col].fillna("").astype(str)
        if code_col
        else pd.Series("", index=df_filtered.index)
    )
    s_pttn = (
        df_filtered[pttn_col].fillna("").astype(str).str.upper()
        if pttn_col
        else pd.Series("", index=df_filtered.index)
    )
    s_size = (
        df_filtered[size_col].fillna("").astype(str).str.upper()
        if size_col
        else pd.Series("", index=df_filtered.index)
    )

    gyeongsang_kw = [
        "부산",
        "대구",
        "울산",
        "경남",
        "경북",
        "경상",
        "밀양",
        "양산",
        "창원",
        "김해",
        "포항",
    ]
    is_gyeongsang = s_addr.str.contains("|".join(gyeongsang_kw), na=False)
    is_express_note = s_note.str.contains("당착|오후", na=False)
    is_dang = is_gyeongsang | is_express_note

    is_jeju = (
        s_addr.str.contains("제주", na=False)
        | s_note.str.contains("제주", na=False)
    )
    is_courier = s_driver.eq("경동택배") & s_courier_type.eq("택배")

    is_blank_time = (
        df_filtered["_출고일시_str"].eq("")
        | df_filtered["_출고일시_str"].str.lower().isin(["nan", "none", "nat"])
        | df_filtered["_출고일시_parsed"].isna()
    )

    df_filtered["당/익/제주/택배"] = np.select(
        [is_blank_time, is_jeju, is_courier, is_dang],
        ["미배차", "제주", "택배", "당"],
        default="익",
    )

    df_filtered["윈터"] = np.where(
        s_code.str.startswith("PXR")
        | s_pttn.str.contains("WINTER|BLIZZAK", na=False)
        | s_size.str.contains("WINTER", na=False),
        "o",
        "",
    )

    barcode_kw = (
        "TURANZA EV|POTENZA SPORT AS|DESTINATION LE3|"
        "DRIVEGUARD PLUS|BLIZZAK 6|DRGU+"
    )
    df_filtered["바코드"] = np.where(
        s_pttn.str.contains(barcode_kw, na=False), "Y", ""
    )

    agg_dict = {
        "출력": ("출력", "max"),
        "윈터": ("윈터", "max"),
        "바코드": ("바코드", "max"),
        "수량": (qty_col, "sum"),
        "_출고일시_parsed": ("_출고일시_parsed", "min"),
        "_출고일시_str": ("_출고일시_str", "first"),
        "_차량기사_str": ("_차량기사_str", "first"),
    }
    if code_col:
        agg_dict["제품가짓수"] = (code_col, "nunique")

    grouped = (
        df_filtered.groupby([type_col, "당/익/제주/택배", addr_col])
        .agg(**agg_dict)
        .reset_index()
    )

    if "제품가짓수" not in grouped.columns:
        grouped["제품가짓수"] = 0

    grouped.rename(
        columns={type_col: "구분", addr_col: "배송처주소"}, inplace=True
    )

    g_cat = grouped["구분"].astype(str).str.upper()
    g_qty = grouped["수량"]
    g_day = grouped["당/익/제주/택배"]
    g_addr = grouped["배송처주소"].astype(str)

    is_large = np.where(
        g_cat.str.contains("화물|TBR", na=False), g_qty >= 70, g_qty >= 100
    )
    is_g_gyeongsang = g_addr.str.contains("|".join(gyeongsang_kw), na=False)

    calculated_note = np.select(
        [(~is_g_gyeongsang) & (g_day == "당") & is_large, is_large, g_day == "미배차"],
        ["장거리당착", "대량", "미배차"],
        default="",
    )
    grouped["비고"] = calculated_note

    return df_filtered, grouped


# =========================================================
# 메인 웹 UI 화면 구성
# =========================================================
st.title("📦 통합 물류 실시간 물량집계 시스템")
st.markdown("출고 파일과 배차 파일을 업로드하면 원하시는 기준대로 자동 분류 및 집계됩니다.")

col1, col2 = st.columns(2)
with col1:
    chuko_file = st.file_uploader("출고 엑셀 파일 업로드", type=["xlsx", "xls"])
with col2:
    batcha_file = st.file_uploader("배차 엑셀 파일 업로드", type=["xlsx", "xls"])

if chuko_file and batcha_file:
    try:
        df_c = pd.read_excel(chuko_file)
        df_b = pd.read_excel(batcha_file)

        df_filtered, grouped = process_logistics_data(df_c, df_b)
        st.success("✨ 데이터 집계가 성공적으로 완료되었습니다!")

        tab1, tab2, tab3 = st.tabs(
            ["📊 통합 집계 결과", "🚚 미배차 목록", "📋 전체 상세 데이터"]
        )

        with tab1:
            st.subheader("구분별 / 당익착별 집계 현황")
            st.dataframe(grouped, use_container_width=True)

        with tab2:
            st.subheader("미배차 목록")
            unassigned = grouped[grouped["당/익/제주/택배"] == "미배차"]
            st.dataframe(unassigned, use_container_width=True)

        with tab3:
            st.subheader("원본 필터링 상세 데이터")
            st.dataframe(df_filtered, use_container_width=True)

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
else:
    st.info("💡 위 양쪽에 출고 파일과 배차 파일을 모두 업로드해주세요.")

if st.button("로그아웃"):
    st.session_state.authenticated = False
    st.rerun()