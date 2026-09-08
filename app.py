import os
import pickle
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st

# ==========================================
# 0. 페이지 기본 설정 & 기본 세션 상태 초기화
# ==========================================
st.set_page_config(
    page_title="실시간 물량집계 시스템",
    page_icon="🚚",
    layout="wide"
)

# 데이터 저장 디렉토리 생성
DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "processed_data.pkl")
os.makedirs(DATA_DIR, exist_ok=True)

# 세션 상태 초기화
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "role" not in st.session_state:
    st.session_state["role"] = None  # 'admin' 또는 'user'
if "page_stack" not in st.session_state:
    st.session_state["page_stack"] = ["home_category"]  # 화면 이동 내비게이션 스택
if "selected_category" not in st.session_state:
    st.session_state["selected_category"] = None
if "selected_sub" not in st.session_state:
    st.session_state["selected_sub"] = None
if "selected_detail_row" not in st.session_state:
    st.session_state["selected_detail_row"] = None

# 비밀번호 가져오기 함수 (Secrets 또는 기본값)
def get_password(role_type):
    try:
        return st.secrets["passwords"][role_type]
    except Exception:
        return "bstk2026!" if role_type == "admin" else "1010"

# ==========================================
# 1. 핵심 데이터 처리 로직 (제공된 파이썬 스크립트 기반)
# ==========================================
def process_uploaded_files(files_dict):
    """
    업로드된 엑셀 파일들(출고, 배차, 피킹조회 등)을 전달받아
    통합 데이터프레임과 메타정보를 생성합니다.
    """
    df_c = None
    df_b_list = []

    for name, file_obj in files_dict.items():
        try:
            xls = pd.ExcelFile(file_obj)
            for sheet_name in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet_name)
                file_text = str(name)
                sheet_text = str(sheet_name)

                if df_c is None and ("출고" in file_text or "출고" in sheet_text):
                    df_c = temp_df
                elif "배차" in file_text or "배차" in sheet_text:
                    df_b_list.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({name}): {e}")

    if df_c is None or not df_b_list:
        raise ValueError("출고 파일 또는 배차 파일을 올바르게 인식하지 못했습니다. 파일명을 확인해주세요.")

    df_b = pd.concat(df_b_list, ignore_index=True)

    if "차랑번호" in df_b.columns and "차량번호" not in df_b.columns:
        df_b.rename(columns={"차랑번호": "차량번호"}, inplace=True)

    # 매칭 처리
    df_b["prefix12"] = df_b["일련번호"].astype(str).str[:12]
    df_c["prefix12"] = df_c["출고번호"].astype(str).str[:12]
    df_b["key12"] = df_b["출고처명(거래처)"].astype(str).str.strip() + "_" + df_b["prefix12"]
    df_c["key12"] = df_c["출고처명(거래처)"].astype(str).str.strip() + "_" + df_c["prefix12"]

    df_b["prefix10"] = df_b["일련번호"].astype(str).str[:10]
    df_c["prefix10"] = df_c["출고번호"].astype(str).str[:10]
    df_b["key10"] = df_b["출고처명(거래처)"].astype(str).str.strip() + "_" + df_b["prefix10"]
    df_c["key10"] = df_c["출고처명(거래처)"].astype(str).str.strip() + "_" + df_c["prefix10"]

    map12 = df_c.set_index("key12")["피킹지출력"].to_dict() if "피킹지출력" in df_c.columns else {}
    map10 = df_c.set_index("key10")["피킹지출력"].to_dict() if "피킹지출력" in df_c.columns else {}

    df_b["피킹지출력"] = df_b["key12"].map(map12)
    df_b["피킹지출력"] = df_b["피킹지출력"].fillna(df_b["key10"].map(map10))

    df_b["key_addr"] = df_b["출고처명(거래처)"].astype(str).str.strip() + "_" + df_b["배송처주소"].astype(str).str.strip()
    df_c["key_addr"] = df_c["출고처명(거래처)"].astype(str).str.strip() + "_" + df_c["배송처주소"].astype(str).str.strip()

    if "피킹지출력" in df_c.columns:
        map_addr = df_c.groupby("key_addr")["피킹지출력"].first().to_dict()
        df_b["피킹지출력"] = df_b["피킹지출력"].fillna(df_b["key_addr"].map(map_addr))

    df = df_b.copy()

    # 칼럼 찾기
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
    store_col = find_col(["출고처명(거래처)", "점포명", "거래처"])

    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    df[type_col] = df[type_col].fillna("").astype(str).str.strip()

    # 마케팅 품목 매핑
    ITEM_NAME_MAP = {
        "장우산(25개입)": "장우산", "개폐식 포스터 액자": "개폐식액자", "골프공옐로우": "골프공옐로우",
        "공기압 체크기": "공기압체크기", "뎁스게이지": "뎁스게이지", "물티슈(10EA)": "물티슈",
        "물티슈(30EA)": "물티슈", "미니간판": "미니간판", "반팔티셔츠 2XL": "반팔티",
        "반팔티셔츠 L": "반팔티", "반팔티셔츠 M": "반팔티", "반팔티셔츠 XL": "반팔티",
        "볼펜": "볼펜", "부채": "부채", "아크릴 미니 B 간판": "아크릴 미니간판",
        "일반 알미늄 액자": "알미늄액자", "6.5 OZ 종이컵": "종이컵", "장우산(30개입)": "장우산",
        "주차 알림판 피규어": "주차피규어", "GOLF BALL": "골프공"
    }

    def map_mkt(row):
        t_val = str(row[type_col]).upper()
        if "마케팅" in t_val or "골프공" in t_val:
            s_val = str(row[size_col]).strip() if size_col else ""
            c_val = str(row[code_col]).strip() if code_col else ""
            for k, v in ITEM_NAME_MAP.items():
                if k in s_val or k in c_val: return v
            if s_val and s_val.lower() not in ["nan", "none", ""]: return s_val
            if c_val and c_val.lower() not in ["nan", "none", ""]: return c_val
        return row[type_col]

    df["최종구분"] = df.apply(map_mkt, axis=1)

    s_driver = df[driver_col].fillna("").astype(str).str.strip() if driver_col else pd.Series("", index=df.index)
    s_courier = df[courier_col].fillna("").astype(str).str.strip() if courier_col else pd.Series("", index=df.index)
    s_addr = df[addr_col].fillna("").astype(str).str.strip()
    s_note = df[note_col].fillna("").astype(str).str.strip() if note_col else pd.Series("", index=df.index)

    time_str = df[time_col].fillna("").astype(str).str.strip() if time_col else pd.Series("", index=df.index)
    time_parsed = pd.to_datetime(df[time_col], errors="coerce") if time_col else pd.Series(pd.NaT, index=df.index)

    gyeongsang_kw = ["부산", "대구", "울산", "경남", "경북", "경상", "밀양", "양산", "창원", "김해", "포항"]
    is_gyeongsang = s_addr.str.contains("|".join(gyeongsang_kw), na=False)
    is_express = s_note.str.contains("당착|오후", na=False)
    is_dang = is_gyeongsang | is_express
    is_jeju = s_addr.str.contains("제주", na=False) | s_note.str.contains("제주", na=False)
    is_courier = s_driver.eq("경동택배") & s_courier.eq("택배")
    is_blank = time_str.eq("") | time_str.str.lower().isin(["nan", "none", "nat"]) | time_parsed.isna()

    df["배송유형"] = np.select(
        [is_blank, is_jeju, is_courier, is_dang],
        ["미배차", "제주", "택배", "당"],
        default="익"
    )

    df["출력상태"] = np.where(df[print_col].astype(str).str.upper().str.contains("Y"), "출력", "미출력") if print_col else "미출력"
    df["차량번호_표시"] = df[car_col].fillna("-") if car_col else "-"
    df["점포명_표시"] = df[store_col].fillna("-") if store_col else "-"
    df["수량_표시"] = df[qty_col]
    df["주소_표시"] = df[addr_col]
    df["패턴_표시"] = df[pttn_col].fillna("-") if pttn_col else "-"
    df["사이즈_표시"] = df[size_col].fillna("-") if size_col else "-"
    df["상품코드_표시"] = df[code_col].fillna("-") if code_col else "-"

    processed_data = {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "raw_df": df
    }

    with open(DATA_FILE, "wb") as f:
        pickle.dump(processed_data, f)

    return processed_data

# 저장된 데이터 로드
def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "rb") as f:
                return pickle.load(f)
        except Exception:
            return None
    return None

# ==========================================
# 2. 내비게이션 & 헤더 UI (수정 완료 지점)
# ==========================================
def render_header(data_time_str):
    st.markdown(
        f"""
        <div style="background-color: #F8F9FA; padding: 10px 15px; border-radius: 8px; margin-bottom: 15px; border: 1px solid #E9ECEF;">
            <div style="font-size: 13px; color: #6C757D; font-weight: bold;">
                🕒 최근 업데이트 일자: <span style="color: #0D6EFD;">{data_time_str if data_time_str else "등록된 데이터 없음"}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_nav1, col_nav2, col_nav3 = st.columns([1, 1, 6])
    with col_nav1:
        if st.button("🏠 홈", use_container_width=True):
            st.session_state["page_stack"] = ["home_category"]
            st.rerun()
    with col_nav2:
        if st.button("↩ 뒤로가기", use_container_width=True):
            if len(st.session_state["page_stack"]) > 1:
                st.session_state["page_stack"].pop()
                st.rerun()
    with col_nav3:
        if st.session_state["authenticated"]:
            if st.button("🔒 로그아웃", key="logout_btn"):
                st.session_state["authenticated"] = False
                st.session_state["role"] = None
                st.rerun()

# ==========================================
# 3. 로그인 및 모드 분리 화면
# ==========================================
def render_login():
    st.title("🚚 실시간 물량집계 시스템")
    st.subheader("접속 모드를 선택하고 비밀번호를 입력해주세요.")

    tab1, tab2 = st.tabs(["🔒 관리자 모드", "👤 일반 모드"])

    with tab1:
        st.info("관리자 모드는 파일 업로드 및 데이터 업데이트가 가능합니다.")
        admin_pwd = st.text_input("관리자 비밀번호", type="password", key="admin_pwd_input")
        if st.button("관리자 로그인", use_container_width=True):
            if admin_pwd == get_password("admin"):
                st.session_state["authenticated"] = True
                st.session_state["role"] = "admin"
                st.success("관리자로 로그인되었습니다.")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")

    with tab2:
        st.info("일반 모드는 업데이트된 물량 조회 및 상세 현황을 확인합니다.")
        user_pwd = st.text_input("일반 비밀번호", type="password", key="user_pwd_input")
        if st.button("일반 사용자 로그인", use_container_width=True):
            if user_pwd == get_password("user"):
                st.session_state["authenticated"] = True
                st.session_state["role"] = "user"
                st.success("로그인되었습니다.")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")

# ==========================================
# 4. 화면별 뷰 레이아웃
# ==========================================
def render_main_app():
    data = load_data()
    updated_at = data["updated_at"] if data else None
    df = data["raw_df"] if data else None

    render_header(updated_at)

    # 관리자 전용 파일 업로드 섹션
    if st.session_state["role"] == "admin":
        with st.expander("📂 [관리자] 신규 데이터 파일 업로드 (출고 / 배차 / 피킹조회)", expanded=(df is None)):
            uploaded_files = st.file_uploader(
                "엑셀 파일(.xls, .xlsx)을 여러 개 첨부하세요.",
                type=["xls", "xlsx"],
                accept_multiple_files=True
            )
            if st.button("데이터 분석 및 업데이트 처리", type="primary"):
                if uploaded_files:
                    try:
                        files_dict = {f.name: f for f in uploaded_files}
                        with st.spinner("엑셀 매칭 및 물량 집계 중..."):
                            process_uploaded_files(files_dict)
                        st.success("데이터 업데이트 성공!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"오류 발생: {e}")
                else:
                    st.warning("업로드할 파일을 선택해주세요.")

    if df is None:
        st.warning("현재 등록된 데이터가 없습니다. 관리자 모드에서 파일을 업로드해주세요.")
        return

    current_page = st.session_state["page_stack"][-1]

    # --- 1단계: 승용 / 화물 / 마케팅 선택 ---
    if current_page == "home_category":
        st.markdown("<h2 style='text-align: center;'>구분 선택</h2>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🚗\n\n승용", use_container_width=True, key="cat_psr"):
                st.session_state["selected_category"] = "승용"
                st.session_state["page_stack"].append("sub_category")
                st.rerun()
        with col2:
            if st.button("🚚\n\n화물", use_container_width=True, key="cat_tbr"):
                st.session_state["selected_category"] = "화물"
                st.session_state["page_stack"].append("sub_category")
                st.rerun()
        with col3:
            if st.button("🎁\n\n마케팅", use_container_width=True, key="cat_mkt"):
                st.session_state["selected_category"] = "마케팅"
                st.session_state["page_stack"].append("sub_category")
                st.rerun()

    # --- 2단계: 당착 / 익착 / 택배 / 미배차 선택 ---
    elif current_page == "sub_category":
        cat = st.session_state["selected_category"]
        st.markdown(f"<h2 style='text-align: center;'>[{cat}] 배송 유형 선택</h2>", unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            if st.button("⚡ 당착", use_container_width=True):
                st.session_state["selected_sub"] = "당"
                st.session_state["page_stack"].append("list_view")
                st.rerun()
        with col2:
            if st.button("🌙 익착", use_container_width=True):
                st.session_state["selected_sub"] = "익"
                st.session_state["page_stack"].append("list_view")
                st.rerun()
        with col3:
            if st.button("📦 택배", use_container_width=True):
                st.session_state["selected_sub"] = "택배"
                st.session_state["page_stack"].append("list_view")
                st.rerun()
        with col4:
            if st.button("⚠️ 미배차", use_container_width=True):
                st.session_state["selected_sub"] = "미배차"
                st.session_state["page_stack"].append("list_view")
                st.rerun()

    # --- 3단계: 목록 현황 ---
    elif current_page == "list_view":
        cat = st.session_state["selected_category"]
        sub = st.session_state["selected_sub"]

        st.markdown(f"### 📋 {cat} > {sub} 현황 목록")

        if cat == "승용":
            sub_df = df[df["최종구분"].str.upper().str.contains("승용|PSR", na=False)]
        elif cat == "화물":
            sub_df = df[df["최종구분"].str.upper().str.contains("화물|TBR", na=False)]
        else:
            sub_df = df[~df["최종구분"].str.upper().str.contains("승용|화물|PSR|TBR", na=False)]

        if sub in ["익", "제주"]:
            sub_df = sub_df[sub_df["배송유형"].isin(["익", "제주"])]
        else:
            sub_df = sub_df[sub_df["배송유형"] == sub]

        filter_status = st.radio("출력 상태 선택", ["전체", "출력", "미출력"], horizontal=True)
        if filter_status != "전체":
            sub_df = sub_df[sub_df["출력상태"] == filter_status]

        search_kw = st.text_input("🔍 검색어 입력 (차량번호 / 점포명 / 주소)", "")
        if search_kw:
            sub_df = sub_df[
                sub_df["차량번호_표시"].astype(str).str.contains(search_kw, case=False) |
                sub_df["점포명_표시"].astype(str).str.contains(search_kw, case=False) |
                sub_df["주소_표시"].astype(str).str.contains(search_kw, case=False)
            ]

        display_cols = ["출력상태", "차량번호_표시", "점포명_표시", "수량_표시", "주소_표시"]
        renamed_cols = {"출력상태": "출력여부", "차량번호_표시": "차량번호", "점포명_표시": "점포명", "수량_표시": "수량", "주소_표시": "주소"}

        show_df = sub_df[display_cols].rename(columns=renamed_cols)

        st.caption("👇 항목을 클릭하여 세부사항을 확인하세요.")

        event = st.dataframe(
            show_df,
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row"
        )

        selected_rows = event.selection.get("rows", [])
        if selected_rows:
            idx = selected_rows[0]
            selected_row_data = sub_df.iloc[idx]
            st.session_state["selected_detail_row"] = selected_row_data
            st.session_state["page_stack"].append("detail_view")
            st.rerun()

    # --- 4단계: 세부사항 뷰 ---
    elif current_page == "detail_view":
        row_data = st.session_state.get("selected_detail_row", None)

        st.markdown("### 🔍 리스트 누름: 세부사항 뷰")

        if row_data is not None:
            st.text_input("🔍 세부항목 내 검색", key="detail_search_kw")

            matching_items = df[
                (df["점포명_표시"] == row_data["점포명_표시"]) &
                (df["주소_표시"] == row_data["주소_표시"])
            ]

            detail_display = matching_items[["패턴_표시", "사이즈_표시", "상품코드_표시", "최종구분", "수량_표시"]].rename(
                columns={
                    "패턴_표시": "패턴",
                    "사이즈_표시": "사이즈",
                    "상품코드_표시": "상품코드",
                    "최종구분": "구분",
                    "수량_표시": "수량"
                }
            )

            st.dataframe(detail_display, use_container_width=True)
        else:
            st.warning("선택된 상세 정보가 없습니다.")

# ==========================================
# 5. 메인 실행 흐름
# ==========================================
def main():
    if not st.session_state["authenticated"]:
        render_login()
    else:
        render_main_app()

if __name__ == "__main__":
    main()