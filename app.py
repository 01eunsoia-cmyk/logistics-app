from datetime import datetime, timezone, timedelta
import glob
import re
import numpy as np
import pandas as pd
import streamlit as st

# 페이지 기본 설정
st.set_page_config(page_title="배차 및 재고 관리 시스템", layout="wide")


# --- KST (한국 표준시) 기준 날짜 가져오기 ---
def get_kst_today():
    return datetime.now(timezone(timedelta(hours=9))).strftime("%Y-%m-%d")


# --- 마케팅 품목 매핑 테이블 ---
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


def map_mkt(row):
    p_val = str(row.get("PTTN", "")).upper().strip()
    s_val = str(row.get("SIZE", "")).upper().strip()
    c_val = str(row.get("품목코드", "")).upper().strip()
    t_val = str(row.get("타이어구분", "")).upper().strip()

    if "TUBE" in p_val or "TUBE" in s_val:
        return "TBR"

    if "마케팅" in t_val or "골프공" in t_val:
        for k, v in ITEM_NAME_MAP.items():
            if k in s_val or k in c_val:
                return v
        if s_val and s_val.lower() not in ["nan", "none", ""]:
            return s_val
        if c_val and c_val.lower() not in ["nan", "none", ""]:
            return c_val

    return str(row.get("타이어구분", "")).strip()


# --- 배차 데이터 로드 함수 ---
@st.cache_data(ttl=60)
def load_dispatch_data():
    files = glob.glob("*.xls*")
    b_files = [f for f in files if "배차" in f]

    if not b_files:
        return pd.DataFrame(), ""

    df_list = []
    latest_time = ""

    for f in b_files:
        try:
            xls = pd.ExcelFile(f)
            for sheet in xls.sheet_names:
                df_temp = pd.read_excel(xls, sheet_name=sheet)
                df_list.append(df_temp)

            match = re.search(r"\d{14}", f)
            if match:
                t_str = match.group(0)
                formatted_t = f"{t_str[:4]}-{t_str[4:6]}-{t_str[6:8]} {t_str[8:10]}:{t_str[10:12]}:{t_str[12:14]}"
                if formatted_t > latest_time:
                    latest_time = formatted_t
        except Exception:
            continue

    if not df_list:
        return pd.DataFrame(), ""

    df_all = pd.concat(df_list, ignore_index=True)

    df_all["수량_표시"] = pd.to_numeric(
        df_all.get("배차수량", 0), errors="coerce"
    ).fillna(0)

    kst_today = get_kst_today()
    if "예정일자" in df_all.columns:
        parsed_dates = pd.to_datetime(
            df_all["예정일자"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")
        df_all["물량구분"] = np.where(parsed_dates > kst_today, "익일", "당일")
    else:
        df_all["물량구분"] = "당일"

    df_all["최종구분"] = df_all.apply(map_mkt, axis=1)

    def set_dispatch_type(row):
        t = str(row.get("택배구분", "")).strip()
        car = str(row.get("차랑번호", "")).strip()
        if "택배" in t:
            return "택배"
        elif car in ["미배차", "", "nan", "None"]:
            return "미배차"
        elif "당" in t:
            return "당"
        else:
            return "익"

    df_all["배송유형"] = df_all.apply(set_dispatch_type, axis=1)

    return df_all, latest_time


# --- 재고 데이터 로드 함수 ---
@st.cache_data(ttl=60)
def load_stock_data():
    files = glob.glob("*.xls*")
    s_files = [f for f in files if "재고" in f]

    if not s_files:
        return pd.DataFrame(), ""

    latest_file = max(s_files)
    try:
        xls = pd.ExcelFile(latest_file)
        df_stock = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

        # 요청한 11개 컬럼 순서 지정
        target_cols = [
            "타이어구분",
            "품목코드",
            "SIZE",
            "PTTN",
            "품목구분",
            "입고로케이션",
            "DOT",
            "전체재고수량",
            "가용재고수량",
            "컨테이너번호",
            "입고일자",
        ]

        existing_cols = [c for c in target_cols if c in df_stock.columns]
        df_stock = df_stock[existing_cols]

        # 입고일자 날짜 포맷 정리
        if "입고일자" in df_stock.columns:
            df_stock["입고일자"] = pd.to_datetime(
                df_stock["입고일자"], errors="coerce"
            ).dt.strftime("%Y-%m-%d")

        return df_stock, latest_file
    except Exception:
        return pd.DataFrame(), ""


# --- 세션 상태 초기화 ---
if "main_mode" not in st.session_state:
    st.session_state["main_mode"] = "배차관리"  # '배차관리' or '재고조회'
if "page_stack" not in st.session_state:
    st.session_state["page_stack"] = ["home_category"]
if "selected_category" not in st.session_state:
    st.session_state["selected_category"] = None
if "selected_sub" not in st.session_state:
    st.session_state["selected_sub"] = None

df, update_time = load_dispatch_data()
df_stock, stock_file_name = load_stock_data()

# --- CSS 스타일 (이모티콘 제거, 글씨 크기 대폭 확대) ---
st.markdown(
    """
<style>
    .stApp {
        background-color: #F8F9FA;
    }

    /* 네비게이션 버튼 텍스트 확대 */
    .stButton > button {
        width: 100%;
        border: none !important;
        border-radius: 16px !important;
        padding: 22px 16px !important;
        color: white !important;
        text-align: left !important;
        box-shadow: 0 6px 16px rgba(0,0,0,0.06);
        transition: all 0.2s ease-in-out !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 24px rgba(0,0,0,0.12);
    }

    /* 카드 제목 및 수량 크기 확대 */
    .card-title-inline {
        font-size: 1.55rem !important;
        font-weight: 800 !important;
        letter-spacing: -0.5px;
        line-height: 1.2;
    }
    .card-title-sub {
        font-size: 1.35rem !important;
        font-weight: 700 !important;
        margin-bottom: 8px;
    }
    .card-value-large {
        font-size: 2.8rem !important;
        font-weight: 900 !important;
        letter-spacing: -1px;
    }

    /* 메인 모드 탭 버튼 스타일 */
    div[data-testid="stHorizontalBlock"] .mode-btn button {
        background-color: #FFFFFF !important;
        color: #1E293B !important;
        border: 2px solid #CBD5E1 !important;
        font-size: 1.2rem !important;
        font-weight: bold !important;
        text-align: center !important;
        padding: 12px !important;
    }

    /* --- 1단계 컬러 그라데이션 --- */
    .cat-psr > button {
        background: linear-gradient(135deg, #1E40AF 0%, #3B82F6 100%) !important;
    }
    .cat-tbr > button {
        background: linear-gradient(135deg, #C2410C 0%, #EA580C 100%) !important;
    }
    .cat-mkt > button {
        background: linear-gradient(135deg, #047857 0%, #10B981 100%) !important;
    }

    /* --- 2단계 컬러 그라데이션 --- */
    .sub-dang > button {
        background: linear-gradient(135deg, #D97706 0%, #F59E0B 100%) !important;
    }
    .sub-ik > button {
        background: linear-gradient(135deg, #2563EB 0%, #60A5FA 100%) !important;
    }
    .sub-courier > button {
        background: linear-gradient(135deg, #059669 0%, #34D399 100%) !important;
    }
    .sub-unassigned > button {
        background: linear-gradient(135deg, #374151 0%, #6B7280 100%) !important;
    }
    .sub-tomorrow > button {
        background: linear-gradient(135deg, #7C3AED 0%, #A78BFA 100%) !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# --- 상단 헤더 & 메뉴 ---
header_col1, header_col2 = st.columns([6, 4])
with header_col1:
    st.markdown(
        f"<h3 style='margin:0; font-weight:800;'>최근 업데이트 (KST): <span style='color:#2563EB;'>{update_time}</span></h3>",
        unsafe_allow_html=True,
    )
with header_col2:
    m_col1, m_col2 = st.columns(2)
    with m_col1:
        if st.button("배차 관리", use_container_width=True, key="mode_dispatch"):
            st.session_state["main_mode"] = "배차관리"
            st.rerun()
    with m_col2:
        if st.button("재고 조회", use_container_width=True, key="mode_stock"):
            st.session_state["main_mode"] = "재고조회"
            st.rerun()

st.divider()

# ==========================================
# 모드 1: 배차 관리
# ==========================================
if st.session_state["main_mode"] == "배차관리":
    # 상단 네비게이션
    nav_col1, nav_col2, nav_col3, _ = st.columns([1.5, 1.5, 1.5, 5.5])
    with nav_col1:
        if st.button("홈", use_container_width=True, key="nav_home"):
            st.session_state["page_stack"] = ["home_category"]
            st.session_state["selected_category"] = None
            st.session_state["selected_sub"] = None
            st.rerun()

    with nav_col2:
        if len(st.session_state["page_stack"]) > 1:
            if st.button("뒤로가기", use_container_width=True, key="nav_back"):
                st.session_state["page_stack"].pop()
                st.rerun()

    with nav_col3:
        if st.button("로그아웃", use_container_width=True, key="nav_logout"):
            st.info("로그아웃 되었습니다.")

    st.write("")

    if df.empty:
        st.warning("분석할 배차 데이터 파일이 없습니다.")
        st.stop()

    current_page = st.session_state["page_stack"][-1]

    # --- 1단계: 구분 선택 ---
    if current_page == "home_category":
        st.markdown(
            "<h2 style='text-align: left; font-weight:800; font-size:2.2rem; margin-bottom: 20px;'>구분 선택</h2>",
            unsafe_allow_html=True,
        )

        df_today = df[df["물량구분"] != "익일"]

        psr_df = df_today[
            df_today["최종구분"].str.upper().str.contains("승용|PSR", na=False)
        ]
        tbr_df = df_today[
            df_today["최종구분"].str.upper().str.contains("화물|TBR", na=False)
        ]
        mkt_df = df_today[
            ~df_today["최종구분"]
            .str.upper()
            .str.contains("승용|화물|PSR|TBR", na=False)
        ]

        qty_psr = int(psr_df["수량_표시"].sum())
        qty_tbr = int(tbr_df["수량_표시"].sum())
        qty_mkt = int(mkt_df["수량_표시"].sum())

        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown('<div class="cat-psr">', unsafe_allow_html=True)
            label_html = f"""<div class="card-title-inline">승용 ({qty_psr:,} 개)</div>"""
            if st.button(
                label_html,
                use_container_width=True,
                key="cat_psr",
                type="secondary",
            ):
                st.session_state["selected_category"] = "승용"
                st.session_state["page_stack"].append("sub_category")
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="cat-tbr">', unsafe_allow_html=True)
            label_html = f"""<div class="card-title-inline">화물 ({qty_tbr:,} 개)</div>"""
            if st.button(
                label_html,
                use_container_width=True,
                key="cat_tbr",
                type="secondary",
            ):
                st.session_state["selected_category"] = "화물"
                st.session_state["page_stack"].append("sub_category")
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        with col3:
            st.markdown('<div class="cat-mkt">', unsafe_allow_html=True)
            label_html = f"""<div class="card-title-inline">마케팅 ({qty_mkt:,} 개)</div>"""
            if st.button(
                label_html,
                use_container_width=True,
                key="cat_mkt",
                type="secondary",
            ):
                st.session_state["selected_category"] = "마케팅"
                st.session_state["page_stack"].append("sub_category")
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    # --- 2단계: 배송 유형 선택 ---
    elif current_page == "sub_category":
        cat = st.session_state["selected_category"]
        st.markdown(
            f"<h2 style='text-align: left; font-weight:800; font-size:2.2rem; margin-bottom: 20px;'>[{cat}] 배송 유형 선택</h2>",
            unsafe_allow_html=True,
        )

        if cat == "승용":
            cat_df = df[
                df["최종구분"].str.upper().str.contains("승용|PSR", na=False)
            ]
        elif cat == "화물":
            cat_df = df[
                df["최종구분"].str.upper().str.contains("화물|TBR", na=False)
            ]
        else:
            cat_df = df[
                ~df["최종구분"]
                .str.upper()
                .str.contains("승용|화물|PSR|TBR", na=False)
            ]

        qty_dang = int(
            cat_df[(cat_df["배송유형"] == "당") & (cat_df["물량구분"] == "당일")][
                "수량_표시"
            ].sum()
        )
        qty_ik = int(
            cat_df[
                (cat_df["배송유형"].isin(["익", "제주"]))
                & (cat_df["물량구분"] == "당일")
            ]["수량_표시"].sum()
        )
        qty_courier = int(
            cat_df[
                (cat_df["배송유형"] == "택배") & (cat_df["물량구분"] == "당일")
            ]["수량_표시"].sum()
        )
        qty_unassigned = int(
            cat_df[
                (cat_df["배송유형"] == "미배차") & (cat_df["물량구분"] == "당일")
            ]["수량_표시"].sum()
        )
        qty_tomorrow = int(
            cat_df[cat_df["물량구분"] == "익일"]["수량_표시"].sum()
        )

        col1, col2, col3, col4, col5 = st.columns(5)

        with col1:
            st.markdown('<div class="sub-dang">', unsafe_allow_html=True)
            if st.button(
                f"""<div class="card-title-sub">당착</div><div class="card-value-large">{qty_dang:,} <span style="font-size:1.4rem;">개</span></div>""",
                use_container_width=True,
                key="btn_dang",
            ):
                st.session_state["selected_sub"] = "당"
                st.session_state["page_stack"].append("list_view")
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="sub-ik">', unsafe_allow_html=True)
            if st.button(
                f"""<div class="card-title-sub">익착</div><div class="card-value-large">{qty_ik:,} <span style="font-size:1.4rem;">개</span></div>""",
                use_container_width=True,
                key="btn_ik",
            ):
                st.session_state["selected_sub"] = "익"
                st.session_state["page_stack"].append("list_view")
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        with col3:
            st.markdown('<div class="sub-courier">', unsafe_allow_html=True)
            if st.button(
                f"""<div class="card-title-sub">택배</div><div class="card-value-large">{qty_courier:,} <span style="font-size:1.4rem;">개</span></div>""",
                use_container_width=True,
                key="btn_courier",
            ):
                st.session_state["selected_sub"] = "택배"
                st.session_state["page_stack"].append("list_view")
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        with col4:
            st.markdown('<div class="sub-unassigned">', unsafe_allow_html=True)
            if st.button(
                f"""<div class="card-title-sub">미배차</div><div class="card-value-large">{qty_unassigned:,} <span style="font-size:1.4rem;">개</span></div>""",
                use_container_width=True,
                key="btn_unassigned",
            ):
                st.session_state["selected_sub"] = "미배차"
                st.session_state["page_stack"].append("list_view")
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        with col5:
            st.markdown('<div class="sub-tomorrow">', unsafe_allow_html=True)
            if st.button(
                f"""<div class="card-title-sub">내일물량</div><div class="card-value-large">{qty_tomorrow:,} <span style="font-size:1.4rem;">개</span></div>""",
                use_container_width=True,
                key="btn_tomorrow",
            ):
                st.session_state["selected_sub"] = "내일물량"
                st.session_state["page_stack"].append("list_view")
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    # --- 3단계: 상세 목록 ---
    elif current_page == "list_view":
        cat = st.session_state["selected_category"]
        sub = st.session_state["selected_sub"]

        st.markdown(
            f"<h3 style='font-weight:800;'>{cat} > {sub} 상세 목록</h3>",
            unsafe_allow_html=True,
        )

        if cat == "승용":
            target_df = df[
                df["최종구분"].str.upper().str.contains("승용|PSR", na=False)
            ]
        elif cat == "화물":
            target_df = df[
                df["최종구분"].str.upper().str.contains("화물|TBR", na=False)
            ]
        else:
            target_df = df[
                ~df["최종구분"]
                .str.upper()
                .str.contains("승용|화물|PSR|TBR", na=False)
            ]

        if sub == "내일물량":
            view_df = target_df[target_df["물량구분"] == "익일"]
        else:
            view_df = target_df[
                (target_df["물량구분"] == "당일")
                & (target_df["배송유형"] == sub)
            ]

        if view_df.empty:
            st.info("해당하는 물량이 없습니다.")
        else:
            st.dataframe(view_df, use_container_width=True)

# ==========================================
# 모드 2: 재고 조회
# ==========================================
elif st.session_state["main_mode"] == "재고조회":
    st.markdown(
        "<h2 style='font-weight:800; font-size:2.2rem; margin-bottom: 20px;'>전체 재고 현황 조회</h2>",
        unsafe_allow_html=True,
    )

    if df_stock.empty:
        st.warning("분석할 재고 데이터 파일(재고*.xls)이 없습니다.")
    else:
        # 검색 및 필터 옵션
        f_col1, f_col2, f_col3 = st.columns([2, 2, 4])

        with f_col1:
            tire_types = ["전체"] + list(
                df_stock["타이어구분"].dropna().unique()
            )
            selected_tire = st.selectbox("타이어구분", tire_types)

        with f_col2:
            item_types = ["전체"] + list(
                df_stock["품목구분"].dropna().unique()
            )
            selected_item = st.selectbox("품목구분", item_types)

        with f_col3:
            search_kw = st.text_input(
                "검색어 (품목코드 / SIZE / PTTN / 입고로케이션)", ""
            )

        # 데이터 필터링
        filtered_stock = df_stock.copy()

        if selected_tire != "전체":
            filtered_stock = filtered_stock[
                filtered_stock["타이어구분"] == selected_tire
            ]

        if selected_item != "전체":
            filtered_stock = filtered_stock[
                filtered_stock["품목구분"] == selected_item
            ]

        if search_kw:
            kw = search_kw.strip().lower()
            filtered_stock = filtered_stock[
                filtered_stock["품목코드"]
                .astype(str)
                .str.lower()
                .str.contains(kw)
                | filtered_stock["SIZE"]
                .astype(str)
                .str.lower()
                .str.contains(kw)
                | filtered_stock["PTTN"]
                .astype(str)
                .str.lower()
                .str.contains(kw)
                | filtered_stock["입고로케이션"]
                .astype(str)
                .str.lower()
                .str.contains(kw)
            ]

        # 요약 정보 표시
        total_qty = int(
            pd.to_numeric(
                filtered_stock["전체재고수량"], errors="coerce"
            ).sum()
        )
        avail_qty = int(
            pd.to_numeric(
                filtered_stock["가용재고수량"], errors="coerce"
            ).sum()
        )

        s_sum1, s_sum2, s_sum3 = st.columns(3)
        s_sum1.metric("총 검색 품목 건수", f"{len(filtered_stock):,} 건")
        s_sum2.metric("총 재고 수량", f"{total_qty:,} 개")
        s_sum3.metric("가용 재고 수량", f"{avail_qty:,} 개")

        st.divider()

        # 요청받은 11개 컬럼 순서대로 출력
        st.dataframe(filtered_stock, use_container_width=True, height=600)