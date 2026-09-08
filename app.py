from datetime import datetime, timezone, timedelta
import glob
import re
import numpy as np
import pandas as pd
import streamlit as st

# 페이지 기본 설정
st.set_page_config(page_title="배차 관리 시스템", layout="wide")


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


# --- 데이터 로드 함수 ---
@st.cache_data(ttl=60)
def load_data():
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

            # 파일명에서 최신 시간 추출 (예: 배차처리20260908170158.xls)
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

    # 데이터 가공
    df_all["수량_표시"] = pd.to_numeric(
        df_all.get("배차수량", 0), errors="coerce"
    ).fillna(0)

    # 당일 / 내일물량 구분
    kst_today = get_kst_today()
    if "예정일자" in df_all.columns:
        parsed_dates = pd.to_datetime(
            df_all["예정일자"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")
        df_all["물량구분"] = np.where(parsed_dates > kst_today, "익일", "당일")
    else:
        df_all["물량구분"] = "당일"

    # 최종구분 (승용/화물/마케팅)
    df_all["최종구분"] = df_all.apply(map_mkt, axis=1)

    # 배송유형 정형화
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


# --- 세션 상태 초기화 ---
if "page_stack" not in st.session_state:
    st.session_state["page_stack"] = ["home_category"]
if "selected_category" not in st.session_state:
    st.session_state["selected_category"] = None
if "selected_sub" not in st.session_state:
    st.session_state["selected_sub"] = None

df, update_time = load_data()

# --- 상단 공통 헤더 ---
st.markdown(
    f"🕒 **최근 업데이트 일자 (KST):** `<span style='color:#2e7d32; font-weight:bold;'>{update_time}</span>`",
    unsafe_allow_html=True,
)

nav_col1, nav_col2, nav_col3, _ = st.columns([1, 1, 1, 5])
with nav_col1:
    if st.button("🏠 홈", use_container_width=True):
        st.session_state["page_stack"] = ["home_category"]
        st.session_state["selected_category"] = None
        st.session_state["selected_sub"] = None
        st.rerun()

with nav_col2:
    if len(st.session_state["page_stack"]) > 1:
        if st.button("↩️ 뒤로가기", use_container_width=True):
            st.session_state["page_stack"].pop()
            st.rerun()

with nav_col3:
    if st.button("🔒 로그아웃", use_container_width=True):
        st.info("로그아웃 되었습니다.")

st.divider()

# 데이터 확인
if df.empty:
    st.warning("분석할 배차 데이터 파일이 없습니다.")
    st.stop()

current_page = st.session_state["page_stack"][-1]

# ==========================================
# 1단계: 승용 / 화물 / 마케팅 선택 (첫 화면)
# ==========================================
if current_page == "home_category":
    st.markdown(
        "<h2 style='text-align: center;'>구분 선택</h2>", unsafe_allow_html=True
    )

    # 내일물량(익일)을 제외한 당일 물량만 수량 집계 대상에 포함
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

    # 첫 화면 전용 커스텀 스타일 (버튼별 색상 부여)
    st.markdown(
        """
    <style>
    div[data-testid="stColumn"]:nth-child(1) button {
        background-color: #E3F2FD !important;
        border: 1.5px solid #2196F3 !important;
        color: #0D47A1 !important;
        font-weight: bold !important;
        font-size: 18px !important;
    }
    div[data-testid="stColumn"]:nth-child(2) button {
        background-color: #FFF3E0 !important;
        border: 1.5px solid #FF9800 !important;
        color: #E65100 !important;
        font-weight: bold !important;
        font-size: 18px !important;
    }
    div[data-testid="stColumn"]:nth-child(3) button {
        background-color: #E8F5E9 !important;
        border: 1.5px solid #4CAF50 !important;
        color: #1B5E20 !important;
        font-weight: bold !important;
        font-size: 18px !important;
    }
    </style>
    """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button(
            f"🚗 승용\n\n({qty_psr:,} 개)",
            use_container_width=True,
            key="cat_psr",
        ):
            st.session_state["selected_category"] = "승용"
            st.session_state["page_stack"].append("sub_category")
            st.rerun()

    with col2:
        if st.button(
            f"🚚 화물\n\n({qty_tbr:,} 개)",
            use_container_width=True,
            key="cat_tbr",
        ):
            st.session_state["selected_category"] = "화물"
            st.session_state["page_stack"].append("sub_category")
            st.rerun()

    with col3:
        if st.button(
            f"🎁 마케팅\n\n({qty_mkt:,} 개)",
            use_container_width=True,
            key="cat_mkt",
        ):
            st.session_state["selected_category"] = "마케팅"
            st.session_state["page_stack"].append("sub_category")
            st.rerun()

# ==========================================
# 2단계: 배송 유형 선택 (당착 / 익착 / 택배 / 미배차 / 내일물량)
# ==========================================
elif current_page == "sub_category":
    cat = st.session_state["selected_category"]
    st.markdown(
        f"<h2 style='text-align: center;'>[{cat}] 배송 유형 선택</h2>",
        unsafe_allow_html=True,
    )

    if cat == "승용":
        cat_df = df[df["최종구분"].str.upper().str.contains("승용|PSR", na=False)]
    elif cat == "화물":
        cat_df = df[df["최종구분"].str.upper().str.contains("화물|TBR", na=False)]
    else:
        cat_df = df[
            ~df["최종구분"].str.upper().str.contains("승용|화물|PSR|TBR", na=False)
        ]

    # 수량 집계 (당착 / 익착 / 택배 / 미배차 / 내일물량)
    # 당착, 익착, 택배, 미배차는 '당일' 물량 중에서만 집계
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
        cat_df[(cat_df["배송유형"] == "택배") & (cat_df["물량구분"] == "당일")][
            "수량_표시"
        ].sum()
    )
    qty_unassigned = int(
        cat_df[(cat_df["배송유형"] == "미배차") & (cat_df["물량구분"] == "당일")][
            "수량_표시"
        ].sum()
    )

    # 5번째 '내일물량' 집계
    qty_tomorrow = int(
        cat_df[cat_df["물량구분"] == "익일"]["수량_표시"].sum()
    )

    # 2단계 5개 버튼 커스텀 CSS (각각 알맞은 색상 부여)
    st.markdown(
        """
    <style>
    div[data-testid="stColumn"] {
        padding: 0 4px;
    }
    /* 1. 당착 (주황/황금) */
    div[data-testid="stColumn"]:nth-child(1) button {
        background-color: #FFF3E0 !important;
        border: 1.5px solid #FF9800 !important;
        color: #E65100 !important;
        font-weight: bold !important;
    }
    /* 2. 익착 (파랑) */
    div[data-testid="stColumn"]:nth-child(2) button {
        background-color: #E3F2FD !important;
        border: 1.5px solid #2196F3 !important;
        color: #0D47A1 !important;
        font-weight: bold !important;
    }
    /* 3. 택배 (초록) */
    div[data-testid="stColumn"]:nth-child(3) button {
        background-color: #E8F5E9 !important;
        border: 1.5px solid #4CAF50 !important;
        color: #1B5E20 !important;
        font-weight: bold !important;
    }
    /* 4. 미배차 (회색/경고) */
    div[data-testid="stColumn"]:nth-child(4) button {
        background-color: #ECEFF1 !important;
        border: 1.5px solid #78909C !important;
        color: #37474F !important;
        font-weight: bold !important;
    }
    /* 5. 내일물량 (보라) */
    div[data-testid="stColumn"]:nth-child(5) button {
        background-color: #F3E5F5 !important;
        border: 1.5px solid #9C27B0 !important;
        color: #4A148C !important;
        font-weight: bold !important;
    }
    </style>
    """,
        unsafe_allow_html=True,
    )

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        if st.button(
            f"⚡ 당착\n\n({qty_dang:,} 개)",
            use_container_width=True,
            key="btn_dang",
        ):
            st.session_state["selected_sub"] = "당"
            st.session_state["page_stack"].append("list_view")
            st.rerun()
    with col2:
        if st.button(
            f"🌙 익착\n\n({qty_ik:,} 개)", use_container_width=True, key="btn_ik"
        ):
            st.session_state["selected_sub"] = "익"
            st.session_state["page_stack"].append("list_view")
            st.rerun()
    with col3:
        if st.button(
            f"📦 택배\n\n({qty_courier:,} 개)",
            use_container_width=True,
            key="btn_courier",
        ):
            st.session_state["selected_sub"] = "택배"
            st.session_state["page_stack"].append("list_view")
            st.rerun()
    with col4:
        if st.button(
            f"⚠️ 미배차\n\n({qty_unassigned:,} 개)",
            use_container_width=True,
            key="btn_unassigned",
        ):
            st.session_state["selected_sub"] = "미배차"
            st.session_state["page_stack"].append("list_view")
            st.rerun()
    with col5:
        if st.button(
            f"📅 내일물량\n\n({qty_tomorrow:,} 개)",
            use_container_width=True,
            key="btn_tomorrow",
        ):
            st.session_state["selected_sub"] = "내일물량"
            st.session_state["page_stack"].append("list_view")
            st.rerun()

# ==========================================
# 3단계: 상세 목록 화면
# ==========================================
elif current_page == "list_view":
    cat = st.session_state["selected_category"]
    sub = st.session_state["selected_sub"]

    st.subheader(f"📋 {cat} > {sub} 상세보기")

    # 선택한 카테고리 필터링
    if cat == "승용":
        target_df = df[df["최종구분"].str.upper().str.contains("승용|PSR", na=False)]
    elif cat == "화물":
        target_df = df[df["최종구분"].str.upper().str.contains("화물|TBR", na=False)]
    else:
        target_df = df[
            ~df["최종구분"].str.upper().str.contains("승용|화물|PSR|TBR", na=False)
        ]

    # 선택한 배송유형 필터링
    if sub == "내일물량":
        view_df = target_df[target_df["물량구분"] == "익일"]
    else:
        view_df = target_df[
            (target_df["물량구분"] == "당일") & (target_df["배송유형"] == sub)
        ]

    if view_df.empty:
        st.info("해당하는 물량이 없습니다.")
    else:
        st.dataframe(view_df, use_container_width=True)