import os
import pickle
from datetime import datetime, timezone, timedelta
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

def get_kst_now():
    kst = timezone(timedelta(hours=9))
    return datetime.now(kst).strftime("%Y-%m-%d %H:%M:%S")

DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "processed_data.pkl")
os.makedirs(DATA_DIR, exist_ok=True)

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "role" not in st.session_state:
    st.session_state["role"] = None
if "page_stack" not in st.session_state:
    st.session_state["page_stack"] = ["home_category"]
if "selected_category" not in st.session_state:
    st.session_state["selected_category"] = None
if "selected_sub" not in st.session_state:
    st.session_state["selected_sub"] = None
if "selected_detail_row" not in st.session_state:
    st.session_state["selected_detail_row"] = None

def get_password(role_type):
    try:
        return st.secrets["passwords"][role_type]
    except Exception:
        return "bstk2026!" if role_type == "admin" else "1010"

# ==========================================
# 1. 핵심 데이터 처리 로직
# ==========================================
def process_uploaded_files(files_dict):
    df_c = None
    df_p = None
    df_i = None
    df_b_list = []

    for name, file_obj in files_dict.items():
        try:
            xls = pd.ExcelFile(file_obj)
            for sheet_name in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet_name)
                file_text = str(name)
                sheet_text = str(sheet_name)

                if df_i is None and ("재고" in file_text or "재고" in sheet_text):
                    df_i = temp_df
                elif df_p is None and ("피킹" in file_text or "피킹" in sheet_text):
                    df_p = temp_df
                elif df_c is None and ("출고" in file_text or "출고" in sheet_text):
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

    if df_p is not None and "일련번호" in df_p.columns:
        df_p["serial_key"] = df_p["일련번호"].astype(str).str.strip()
        df_b["serial_key"] = df_b["일련번호"].astype(str).str.strip()

        loc_map = df_p.groupby("serial_key")["로케이션"].first().to_dict() if "로케이션" in df_p.columns else {}
        dot_map = df_p.groupby("serial_key")["DOT"].first().to_dict() if "DOT" in df_p.columns else {}
        li_map = df_p.groupby("serial_key")["LI"].first().to_dict() if "LI" in df_p.columns else {}
        ss_map = df_p.groupby("serial_key")["SS"].first().to_dict() if "SS" in df_p.columns else {}

        df_b["로케이션_표시"] = df_b["serial_key"].map(loc_map)
        df_b["디오티_표시"] = df_b["serial_key"].map(dot_map)
        df_b["LI_표시"] = df_b["serial_key"].map(li_map)
        df_b["SS_표시"] = df_b["serial_key"].map(ss_map)

    df = df_b.copy()

    def find_col(keywords):
        for keyword in keywords:
            for col in df.columns:
                if keyword in str(col).strip():
                    return col
        return None

    addr_col = find_col(["배송처주소", "주소", "배송지", "수령지"])
    store_col = find_col(["출고처명", "점포명", "거래처", "수령처", "고객명"])
    type_col = find_col(["타이어구분", "구분", "타이어", "분류"])
    qty_col = find_col(["배차수량", "수량", "총수량", "작업수량"])
    car_col = find_col(["차량번호", "차랑번호", "운송사", "차량", "배차"])
    driver_col = find_col(["차량기사", "기사명", "기사"])
    phone_col = find_col(["연락처", "전화번호", "기사연락처"])
    courier_col = find_col(["택배구분", "택배분류", "택배"])
    code_col = find_col(["품목코드", "코드", "자재코드", "SKU"])
    size_col = find_col(["SIZE", "규격", "사이즈"])
    pttn_col = find_col(["PTTN", "패턴"])
    note_col = find_col(["비고", "특이사항", "메모", "품목비고", "요청사항"])
    print_col = find_col(["피킹지출력", "출력", "피킹지"])
    time_col = find_col(["출고일시", "출고시간", "일시"])
    plan_date_col = find_col(["출고예정", "예정일자", "예정일", "배차예정", "출고예정일", "배송예정", "배송예정일", "납기예정", "납기일"])

    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    df[type_col] = df[type_col].fillna("").astype(str).str.strip()

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
        p_val = str(row[pttn_col]).upper().strip() if pttn_col else ""
        s_val = str(row[size_col]).upper().strip() if size_col else ""
        c_val = str(row[code_col]).upper().strip() if code_col else ""
        t_val = str(row[type_col]).upper().strip()

        if "TUBE" in p_val or "TUBE" in s_val:
            return "TBR"

        if "마케팅" in t_val or "골프공" in t_val:
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
    
    # 내일물량 최우선 판별 로직 (출고예정 혹은 출고일시가 내일 이후인 경우)
    kst_today = datetime.now(timezone(timedelta(hours=9))).date()
    tomorrow_date = kst_today + timedelta(days=1)

    plan_date_parsed = pd.to_datetime(df[plan_date_col], errors="coerce").dt.date if plan_date_col else pd.Series(pd.NaT, index=df.index)
    time_parsed = pd.to_datetime(df[time_col], errors="coerce").dt.date if time_col else pd.Series(pd.NaT, index=df.index)

    is_plan_tomorrow = plan_date_parsed.apply(lambda d: d >= tomorrow_date if pd.notna(d) else False)
    is_time_tomorrow = time_parsed.apply(lambda d: d >= tomorrow_date if pd.notna(d) else False)
    is_tomorrow = is_plan_tomorrow | is_time_tomorrow

    gyeongsang_kw = ["부산", "대구", "울산", "경남", "경북", "경상", "밀양", "양산", "창원", "김해", "포항"]
    is_gyeongsang = s_addr.str.contains("|".join(gyeongsang_kw), na=False)
    is_express = s_note.str.contains("당착|오후", na=False)
    is_dang = is_gyeongsang | is_express
    is_jeju = s_addr.str.contains("제주", na=False) | s_note.str.contains("제주", na=False)
    is_courier = s_driver.eq("경동택배") & s_courier.eq("택배")
    is_blank = time_str.eq("") | time_str.str.lower().isin(["nan", "none", "nat"])

    # np.select 순서상 is_tomorrow가 최우선 조건으로 반영됨
    df["배송유형"] = np.select(
        [is_tomorrow, is_blank, is_jeju, is_courier, is_dang],
        ["내일물량", "미배차", "제주", "택배", "당"],
        default="익"
    )

    df["출력상태"] = np.where(df[print_col].astype(str).str.upper().str.contains("Y"), "○", "X") if print_col else "X"
    df["차량번호_표시"] = df[car_col].fillna("-") if car_col else "-"
    df["기사명_표시"] = df[driver_col].fillna("-") if driver_col else "-"
    df["연락처_표시"] = df[phone_col].fillna("-") if phone_col else "-"
    df["점포명_표시"] = df[store_col].fillna("-") if store_col else "-"
    df["수량_표시"] = df[qty_col]
    df["주소_표시"] = df[addr_col]
    df["출고일시_표시"] = time_str

    car_str = df["차량번호_표시"].astype(str)
    driver_str = df["기사명_표시"].astype(str)
    
    is_temp_car = car_str.str.contains("9999") | driver_str.str.contains("임의차량")
    df["차량그룹키"] = np.where(
        is_temp_car,
        df["차량번호_표시"].astype(str) + "_" + df["출고일시_표시"].astype(str),
        df["차량번호_표시"].astype(str)
    )

    df["타이어구분_표시"] = df[type_col].fillna("-") if type_col else "-"
    df["패턴_표시"] = df[pttn_col].fillna("-") if pttn_col else "-"
    df["사이즈_표시"] = df[size_col].fillna("-") if size_col else "-"
    df["상품코드_표시"] = df[code_col].fillna("-") if code_col else "-"
    
    for col_k in ["로케이션_표시", "디오티_표시", "LI_표시", "SS_표시"]:
        if col_k not in df.columns:
            df[col_k] = "-"
        else:
            df[col_k] = df[col_k].fillna("-").astype(str).str.replace(r"\.0$", "", regex=True)

    processed_data = {
        "updated_at": get_kst_now(),
        "raw_df": df,
        "raw_inventory_df": df_i
    }

    with open(DATA_FILE, "wb") as f:
        pickle.dump(processed_data, f)

    return processed_data

def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "rb") as f:
                return pickle.load(f)
        except Exception:
            return None
    return None

# ==========================================
# 2. 내비게이션 & 헤더 UI
# ==========================================
def render_header(data_time_str):
    display_time = data_time_str if data_time_str else "등록된 데이터 없음"
    col_info, col_logout = st.columns([7, 2])
    with col_info:
        st.markdown(f"🕒 **최근 업데이트 일자 (KST):** `{display_time}`")
    with col_logout:
        if st.session_state["authenticated"]:
            if st.button("🔒 로그아웃", key="logout_btn", use_container_width=True):
                st.session_state["authenticated"] = False
                st.session_state["role"] = None
                st.rerun()

def render_top_nav():
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🏠 홈으로 이동", key="top_btn_home", use_container_width=True):
            st.session_state["page_stack"] = ["home_category"]
            st.rerun()
    with col2:
        is_disabled = len(st.session_state["page_stack"]) <= 1
        if st.button("↩ 이전 화면 (뒤로가기)", key="top_btn_back", use_container_width=True, disabled=is_disabled):
            if not is_disabled:
                st.session_state["page_stack"].pop()
                st.rerun()

def render_bottom_nav():
    st.write("<br><br><br>", unsafe_allow_html=True)
    st.markdown("""
        <style>
            div[data-testid="stVerticalBlock"] > div.element-container:has(div.bottom-nav-marker) {
                position: fixed;
                bottom: 0;
                left: 0;
                width: 100%;
                background-color: #ffffff;
                padding: 10px 20px;
                box-shadow: 0px -2px 10px rgba(0,0,0,0.1);
                z-index: 99999;
            }
        </style>
        <div class="bottom-nav-marker"></div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🏠 홈으로 이동", key="bot_btn_home", use_container_width=True):
            st.session_state["page_stack"] = ["home_category"]
            st.rerun()
    with col2:
        is_disabled = len(st.session_state["page_stack"]) <= 1
        if st.button("↩ 이전 화면 (뒤로가기)", key="bot_btn_back", use_container_width=True, disabled=is_disabled):
            if not is_disabled:
                st.session_state["page_stack"].pop()
                st.rerun()

# ==========================================
# 3. 로그인 화면
# ==========================================
def render_login():
    st.title("🚚 실시간 물량집계 시스템")
    st.subheader("접속 모드를 선택하고 비밀번호를 입력해주세요.")

    tab1, tab2 = st.tabs(["👤 일반 모드", "🔒 관리자 모드"])

    with tab1:
        user_pwd = st.text_input("일반 비밀번호", type="password", key="user_pwd_input")
        if st.button("일반 사용자 로그인", use_container_width=True):
            if user_pwd == get_password("user"):
                st.session_state["authenticated"] = True
                st.session_state["role"] = "user"
                st.success("로그인되었습니다.")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")

    with tab2:
        admin_pwd = st.text_input("관리자 비밀번호", type="password", key="admin_pwd_input")
        if st.button("관리자 로그인", use_container_width=True):
            if admin_pwd == get_password("admin"):
                st.session_state["authenticated"] = True
                st.session_state["role"] = "admin"
                st.success("관리자로 로그인되었습니다.")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")

# ==========================================
# 4. 메인 어플리케이션
# ==========================================
def render_main_app():
    data = load_data()
    updated_at = data["updated_at"] if data else None
    df = data["raw_df"] if data else None
    df_i = data.get("raw_inventory_df", None) if data else None

    render_header(updated_at)
    render_top_nav()
    st.write("---")

    if st.session_state["role"] == "admin":
        with st.expander("📂 [관리자] 신규 데이터 파일 업로드 (출고 / 배차 / 피킹 / 재고)", expanded=(df is None)):
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

    current_page = st.session_state["page_stack"][-1] if st.session_state["page_stack"] else "home_category"

    # --- 1단계: 승용 / 화물 / 마케팅 / 재고조회 선택 ---
    if current_page == "home_category":
        st.markdown("<h2 style='text-align: center;'>구분 선택</h2>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
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
        with col4:
            if st.button("📦\n\n재고조회", use_container_width=True, key="cat_inventory"):
                st.session_state["selected_category"] = "재고조회"
                st.session_state["page_stack"].append("inventory_view")
                st.rerun()

    # --- 일반 통합 재고조회 뷰 화면 ---
    elif current_page == "inventory_view":
        st.markdown("<h3 style='text-align: center;'>📦 통합 재고조회 현황</h3>", unsafe_allow_html=True)
        
        if df_i is None or df_i.empty:
            st.warning("재고 파일(.xls)이 업로드되지 않았거나 데이터가 없습니다.")
        else:
            search_inv = st.text_input("🔍 재고 검색어 입력 (품목코드 / SIZE / PTTN / 입고로케이션)", "", key="inv_search_box")
            inv_df = df_i.copy()

            def find_inv_col(names):
                for n in names:
                    for col in inv_df.columns:
                        if n.lower() in str(col).lower():
                            return col
                return None

            c_loc = find_inv_col(["입고로케이션"])
            c_code = find_inv_col(["품목코드"])
            c_size = find_inv_col(["SIZE", "사이즈"])
            c_pttn = find_inv_col(["PTTN", "패턴"])
            c_dot = find_inv_col(["DOT"])
            c_avail = find_inv_col(["가용재고수량"])
            c_total = find_inv_col(["전체재고수량"])
            c_item_type = find_inv_col(["품목구분"])
            c_storage_note = find_inv_col(["적치비고"])
            c_date = find_inv_col(["입고일자"])
            c_container = find_inv_col(["컨테이너번호"])

            if c_item_type:
                inv_df["_sort_key"] = np.where(
                    inv_df[c_item_type].fillna("").astype(str).str.contains("불량품", na=False), 0, 1
                )
                inv_df = inv_df.sort_values(by="_sort_key", ascending=True).reset_index(drop=True)

            if search_inv:
                inv_df = inv_df[
                    inv_df["품목코드"].astype(str).str.contains(search_inv, case=False, na=False) |
                    inv_df["SIZE"].astype(str).str.contains(search_inv, case=False, na=False) |
                    inv_df["PTTN"].astype(str).str.contains(search_inv, case=False, na=False) |
                    inv_df["입고로케이션"].astype(str).str.contains(search_inv, case=False, na=False)
                ]

            display_inv = pd.DataFrame()
            if c_loc: display_inv["입고로케이션"] = inv_df[c_loc].fillna("-")
            if c_code: display_inv["품목코드"] = inv_df[c_code].fillna("-")
            if c_size: display_inv["SIZE"] = inv_df[c_size].fillna("-")
            if c_pttn: display_inv["PTTN"] = inv_df[c_pttn].fillna("-")
            if c_dot: display_inv["DOT"] = inv_df[c_dot].fillna("-").astype(str).str.replace(r"\.0$", "", regex=True)
            if c_avail: display_inv["가용재고수량"] = pd.to_numeric(inv_df[c_avail], errors="coerce").fillna(0).astype(int)
            if c_total: display_inv["전체재고수량"] = pd.to_numeric(inv_df[c_total], errors="coerce").fillna(0).astype(int)
            if c_item_type: display_inv["품목구분"] = inv_df[c_item_type].fillna("-")
            if c_storage_note: display_inv["적치비고"] = inv_df[c_storage_note].fillna("-")
            if c_date: display_inv["입고일자"] = inv_df[c_date].fillna("-")
            if c_container: display_inv["컨테이너번호"] = inv_df[c_container].fillna("-")

            def highlight_inventory(data):
                styles = pd.DataFrame("", index=data.index, columns=data.columns)
                if "가용재고수량" in styles.columns:
                    styles["가용재고수량"] = "background-color: rgba(230, 240, 255, 0.6); font-weight: bold;"
                return styles

            styled_inv = display_inv.style.apply(highlight_inventory, axis=None)
            st.dataframe(styled_inv, use_container_width=True, hide_index=True, height=550)

    # --- 마케팅 전용 재고조회 뷰 화면 ---
    elif current_page == "mkt_inventory_view":
        st.markdown("<h3 style='text-align: center;'>🎁 마케팅 / GOLF 재고조회 현황</h3>", unsafe_allow_html=True)
        
        if df_i is None or df_i.empty:
            st.warning("재고 파일(.xls)이 업로드되지 않았거나 데이터가 없습니다.")
        else:
            inv_df = df_i.copy()

            mkt_type_col = None
            for col in inv_df.columns:
                if any(k in str(col) for k in ["품목구분", "타이어구분", "구분", "분류", "타이어"]):
                    mkt_type_col = col
                    break
            
            if mkt_type_col:
                inv_df[mkt_type_col] = inv_df[mkt_type_col].fillna("").astype(str)
                inv_df = inv_df[inv_df[mkt_type_col].str.upper().str.contains("마케팅|GOLF|골프공", na=False)]

            search_mkt_inv = st.text_input("🔍 마케팅 재고 검색어 입력 (품목구분 / 상품코드 / SIZE)", "", key="mkt_inv_search_box")
            if search_mkt_inv:
                def find_c(names):
                    for n in names:
                        for col in inv_df.columns:
                            if n.lower() in str(col).lower():
                                return col
                    return None
                c_code_s = find_c(["품목코드", "상품코드", "코드"])
                c_size_s = find_c(["SIZE", "사이즈"])
                
                cond = pd.Series(False, index=inv_df.index)
                if mkt_type_col:
                    cond = cond | inv_df[mkt_type_col].astype(str).str.contains(search_mkt_inv, case=False, na=False)
                if c_code_s:
                    cond = cond | inv_df[c_code_s].astype(str).str.contains(search_mkt_inv, case=False, na=False)
                if c_size_s:
                    cond = cond | inv_df[c_size_s].astype(str).str.contains(search_mkt_inv, case=False, na=False)
                inv_df = inv_df[cond]

            def find_inv_col(names):
                for n in names:
                    for col in inv_df.columns:
                        if n.lower() in str(col).lower():
                            return col
                    return None
                return None

            c_code = find_inv_col(["품목코드", "상품코드", "코드"])
            c_size = find_inv_col(["SIZE", "사이즈"])
            c_avail = find_inv_col(["가용재고수량", "가용재고"])
            c_total = find_inv_col(["전체재고수량", "총재고", "재고수량"])

            display_mkt_inv = pd.DataFrame()
            display_mkt_inv["품목구분"] = inv_df[mkt_type_col].fillna("-") if mkt_type_col else "-"
            display_mkt_inv["상품코드"] = inv_df[c_code].fillna("-") if c_code else "-"
            display_mkt_inv["SIZE"] = inv_df[c_size].fillna("-") if c_size else "-"
            display_mkt_inv["가용재고수량"] = pd.to_numeric(inv_df[c_avail], errors="coerce").fillna(0).astype(int) if c_avail else 0
            display_mkt_inv["전체재고수량"] = pd.to_numeric(inv_df[c_total], errors="coerce").fillna(0).astype(int) if c_total else 0

            def highlight_mkt_inventory(data):
                styles = pd.DataFrame("", index=data.index, columns=data.columns)
                if "가용재고수량" in styles.columns:
                    styles["가용재고수량"] = "background-color: rgba(230, 240, 255, 0.6); font-weight: bold;"
                return styles

            styled_mkt_inv = display_mkt_inv.style.apply(highlight_mkt_inventory, axis=None)
            st.dataframe(styled_mkt_inv, use_container_width=True, hide_index=True, height=550)

    # --- 2단계: 배송 유형 선택 ---
    elif current_page == "sub_category":
        cat = st.session_state["selected_category"]
        st.markdown(f"<h2 style='text-align: center;'>[{cat}] 배송 유형 선택</h2>", unsafe_allow_html=True)

        if cat == "승용":
            cat_df = df[df["최종구분"].str.upper().str.contains("승용|PSR", na=False)]
        elif cat == "화물":
            cat_df = df[df["최종구분"].str.upper().str.contains("화물|TBR", na=False)]
        else:
            cat_df = df[~df["최종구분"].str.upper().str.contains("승용|화물|PSR|TBR", na=False)]

        qty_dang = int(cat_df[cat_df["배송유형"] == "당"]["수량_표시"].sum())
        qty_ik = int(cat_df[cat_df["배송유형"].isin(["익", "제주"])]["수량_표시"].sum())
        qty_unassigned = int(cat_df[cat_df["배송유형"] == "미배차"]["수량_표시"].sum())
        qty_tomorrow = int(cat_df[cat_df["배송유형"] == "내일물량"]["수량_표시"].sum())
        
        if cat == "승용":
            qty_courier = int(df[df["배송유형"] == "택배"]["수량_표시"].sum())
        elif cat == "화물":
            qty_courier = int(df[(df["배송유형"] == "택배") & (df["최종구분"].str.upper().str.contains("화물|TBR", na=False))]["수량_표시"].sum())
        else:
            qty_courier = int(cat_df[cat_df["배송유형"] == "택배"]["수량_표시"].sum())

        if cat == "마케팅":
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                if st.button(f"⚡ 당착\n\n({qty_dang:,} 개)", use_container_width=True):
                    st.session_state["selected_sub"] = "당"
                    st.session_state["page_stack"].append("list_view")
                    st.rerun()
            with col2:
                if st.button(f"🌙 익착\n\n({qty_ik:,} 개)", use_container_width=True):
                    st.session_state["selected_sub"] = "익"
                    st.session_state["page_stack"].append("list_view")
                    st.rerun()
            with col3:
                if st.button(f"📦 택배\n\n({qty_courier:,} 개)", use_container_width=True):
                    st.session_state["selected_sub"] = "택배"
                    st.session_state["page_stack"].append("list_view")
                    st.rerun()
            with col4:
                if st.button("📦\n\n재고조회", use_container_width=True):
                    st.session_state["page_stack"].append("mkt_inventory_view")
                    st.rerun()
        else:
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                if st.button(f"⚡ 당착\n\n({qty_dang:,} 개)", use_container_width=True):
                    st.session_state["selected_sub"] = "당"
                    st.session_state["page_stack"].append("list_view")
                    st.rerun()
            with col2:
                if st.button(f"🌙 익착\n\n({qty_ik:,} 개)", use_container_width=True):
                    st.session_state["selected_sub"] = "익"
                    st.session_state["page_stack"].append("list_view")
                    st.rerun()
            with col3:
                if st.button(f"⚠️ 미배차\n\n({qty_unassigned:,} 개)", use_container_width=True):
                    st.session_state["selected_sub"] = "미배차"
                    st.session_state["page_stack"].append("list_view")
                    st.rerun()
            with col4:
                if st.button(f"📦 택배\n\n({qty_courier:,} 개)", use_container_width=True):
                    st.session_state["selected_sub"] = "택배"
                    st.session_state["page_stack"].append("list_view")
                    st.rerun()
            with col5:
                if st.button(f"📅 내일물량\n\n({qty_tomorrow:,} 개)", use_container_width=True):
                    st.session_state["selected_sub"] = "내일물량"
                    st.session_state["page_stack"].append("list_view")
                    st.rerun()

    # --- 3단계: 목록 현황 ---
    elif current_page == "list_view":
        cat = st.session_state["selected_category"]
        sub = st.session_state["selected_sub"]

        st.markdown(f"### 📋 {cat} > {sub} 현황 목록")

        if sub == "택배":
            if cat == "승용":
                sub_df = df[df["배송유형"] == "택배"].copy()
            elif cat == "화물":
                sub_df = df[(df["배송유형"] == "택배") & (df["최종구분"].str.upper().str.contains("화물|TBR", na=False))].copy()
            else:
                sub_df = df[(df["배송유형"] == "택배") & (~df["최종구분"].str.upper().str.contains("승용|화물|PSR|TBR", na=False))].copy()
        else:
            if cat == "승용":
                sub_df = df[df["최종구분"].str.upper().str.contains("승용|PSR", na=False)].copy()
            elif cat == "화물":
                sub_df = df[df["최종구분"].str.upper().str.contains("화물|TBR", na=False)].copy()
            else:
                sub_df = df[~df["최종구분"].str.upper().str.contains("승용|화물|PSR|TBR", na=False)].copy()

            if sub in ["익", "제주"]:
                sub_df = sub_df[sub_df["배송유형"].isin(["익", "제주"])]
            else:
                sub_df = sub_df[sub_df["배송유형"] == sub]

        total_qty = int(sub_df["수량_표시"].sum())

        if sub == "택배":
            st.metric("📦 택배 총 수량", f"{total_qty:,} 개")
            st.write("---")

            search_kw = st.text_input("🔍 검색어 입력 (타이어구분 / 패턴 / 사이즈 / 상품코드)", "", key="courier_search")
            if search_kw:
                sub_df = sub_df[
                    sub_df["타이어구분_표시"].astype(str).str.contains(search_kw, case=False) |
                    sub_df["패턴_표시"].astype(str).str.contains(search_kw, case=False) |
                    sub_df["사이즈_표시"].astype(str).str.contains(search_kw, case=False) |
                    sub_df["상품코드_표시"].astype(str).str.contains(search_kw, case=False)
                ]

            grouped_df = sub_df.groupby(
                ["타이어구분_표시", "상품코드_표시", "사이즈_표시", "패턴_표시"], 
                as_index=False
            )["수량_표시"].sum()

            grouped_df = grouped_df.sort_values(
                by=["패턴_표시", "사이즈_표시"],
                ascending=[False, False]
            ).reset_index(drop=True)

            col_order = ["타이어구분_표시", "상품코드_표시", "사이즈_표시", "패턴_표시", "수량_표시"]
            show_df = grouped_df[col_order].rename(columns={
                "타이어구분_표시": "타이어구분",
                "상품코드_표시": "상품코드",
                "사이즈_표시": "사이즈",
                "패턴_표시": "패턴",
                "수량_표시": "수량"
            })

            styled_show_df = show_df.style.map(
                lambda val: "font-weight: bold;", subset=["수량"]
            )

            st.dataframe(styled_show_df, use_container_width=True, hide_index=True, height=550)

        else:
            sub_df["출력상태"] = sub_df["출력상태"].replace({"출력": "○", "미출력": "X"})
            printed_qty = int(sub_df[sub_df["출력상태"] == "○"]["수량_표시"].sum())
            unprinted_qty = int(sub_df[sub_df["출력상태"] == "X"]["수량_표시"].sum())

            m_col1, m_col2, m_col3 = st.columns(3)
            m_col1.metric("📦 전체 총수량", f"{total_qty:,} 개")
            m_col2.metric("✅ 출력 완료 수량 (○)", f"{printed_qty:,} 개")
            m_col3.metric("⏳ 미출력 수량 (X)", f"{unprinted_qty:,} 개")

            st.write("---")

            filter_status = st.radio("출력 상태 선택", ["전체", "○ (출력)", "X (미출력)"], horizontal=True, key="print_status_radio")
            if filter_status == "○ (출력)":
                sub_df = sub_df[sub_df["출력상태"] == "○"]
            elif filter_status == "X (미출력)":
                sub_df = sub_df[sub_df["출력상태"] == "X"]

            search_kw = st.text_input("🔍 검색어 입력 (차량번호 / 점포명 / 주소 / 기사명)", "", key="list_search_box")
            if search_kw:
                sub_df = sub_df[
                    sub_df["차량번호_표시"].astype(str).str.contains(search_kw, case=False) |
                    sub_df["점포명_표시"].astype(str).str.contains(search_kw, case=False) |
                    sub_df["주소_표시"].astype(str).str.contains(search_kw, case=False) |
                    sub_df["기사명_표시"].astype(str).str.contains(search_kw, case=False)
                ]

            grouped_df = sub_df.groupby(
                ["출력상태", "점포명_표시", "주소_표시", "차량번호_표시", "기사명_표시", "연락처_표시", "차량그룹키", "배송유형"], 
                as_index=False
            )["수량_표시"].sum()

            grouped_df = grouped_df.sort_values(by=["차량그룹키", "점포명_표시"]).reset_index(drop=True)

            renamed_cols = {
                "출력상태": "출력여부",
                "점포명_표시": "점포명",
                "주소_표시": "주소",
                "수량_표시": "수량",
                "차량번호_표시": "차량번호",
                "기사명_표시": "기사명",
                "연락처_표시": "연락처"
            }

            col_order = ["출력상태", "점포명_표시", "주소_표시", "수량_표시", "차량번호_표시", "기사명_표시", "연락처_표시", "차량그룹키", "배송유형"]
            show_df = grouped_df[col_order].rename(columns=renamed_cols)

            def highlight_vehicles(data):
                car_keys = grouped_df["차량그룹키"].tolist()
                styles = pd.DataFrame("", index=data.index, columns=data.columns)
                color_flag = False
                prev_key = None
                for i, key in enumerate(car_keys):
                    if prev_key is not None and key != prev_key:
                        color_flag = not color_flag
                    prev_key = key
                    
                    if color_flag:
                        styles.iloc[i] = "background-color: rgba(128, 128, 128, 0.12);"
                    else:
                        styles.iloc[i] = "background-color: transparent;"
                
                if "수량" in styles.columns:
                    styles["수량"] = styles["수량"] + " font-weight: bold;"
                return styles

            display_df = show_df.drop(columns=["차량그룹키", "배송유형"])
            styled_df = display_df.style.apply(highlight_vehicles, axis=None)

            st.dataframe(styled_df, use_container_width=True, hide_index=True, height=550)

            st.write("---")
            st.caption("👇 세부사항을 확인할 점포명을 선택하세요.")
            
            store_options = ["선택하세요"] + show_df["점포명"].unique().tolist()
            selected_store_choice = st.selectbox("🔍 점포 선택", store_options, key="select_store_dropdown")

            if selected_store_choice != "선택하세요":
                matched_rows = show_df[show_df["점포명"] == selected_store_choice]
                if len(matched_rows) == 1:
                    row_data = matched_rows.iloc[0]
                else:
                    sub_opts = [f"{r.점포명} (차량: {r.차량번호} / 주소: {r.주소})" for _, r in matched_rows.iterrows()]
                    sub_choice = st.selectbox("동일한 점포명이 존재합니다. 세부 건을 선택하세요:", sub_opts, key="sub_store_choice")
                    chosen_idx = sub_opts.index(sub_choice)
                    row_data = matched_rows.iloc[chosen_idx]

                if st.button("🚀 선택한 점포 세부 품목 보기", use_container_width=True, type="primary", key="go_detail_btn"):
                    st.session_state["selected_detail_row"] = row_data
                    st.session_state["page_stack"].append("detail_view")
                    st.rerun()

    # --- 4단계: 세부사항 뷰 ---
    elif current_page == "detail_view":
        row_data = st.session_state.get("selected_detail_row", None)

        st.markdown("### 🔍 세부사항 뷰")

        if row_data is not None:
            st.info(f"**점포명:** {row_data['점포명']} | **차량번호:** {row_data['차량번호']} | **기사명:** {row_data['기사명']} | **연락처:** {row_data['연락처']} | **주소:** {row_data['주소']}")

            sub_type = row_data.get("배송유형", None)
            
            match_cond = (
                (df["점포명_표시"] == row_data["점포명"]) &
                (df["주소_표시"] == row_data["주소"]) &
                (df["차량그룹키"] == row_data["차량그룹키"])
            )
            if sub_type:
                match_cond = match_cond & (df["배송유형"] == sub_type)

            matching_items = df[match_cond].copy()

            if "품목코드" in matching_items.columns:
                matching_items["품목코드_표시"] = matching_items["품목코드"].fillna("-")
            elif "품목코드_표시" in matching_items.columns:
                pass
            elif "상품코드_표시" in matching_items.columns:
                matching_items["품목코드_표시"] = matching_items["상품코드_표시"].fillna("-")
            else:
                matching_items["품목코드_표시"] = "-"

            if "품목비고" in matching_items.columns:
                matching_items["품목비고_표시"] = matching_items["품목비고"].fillna("-")
            else:
                matching_items["품목비고_표시"] = "-"

            for c_name, default_val in [
                ("로케이션_표시", "-"), ("디오티_표시", "-"), 
                ("LI_표시", "-"), ("SS_표시", "-"),
                ("패턴_표시", "-"), ("사이즈_표시", "-"), ("수량_표시", 0)
            ]:
                if c_name not in matching_items.columns:
                    matching_items[c_name] = default_val

            detail_display = pd.DataFrame({
                "로케이션": matching_items["로케이션_표시"],
                "수량": matching_items["수량_표시"],
                "디오티": matching_items["디오티_표시"],
                "상품코드": matching_items["품목코드_표시"],
                "사이즈": matching_items["사이즈_표시"],
                "LI": matching_items["LI_표시"],
                "SS": matching_items["SS_표시"],
                "품목비고": matching_items["품목비고_표시"]
            })

            detail_display = detail_display.sort_values(
                by=["로케이션", "사이즈"], 
                ascending=[False, False]
            )

            def highlight_detail_row(row):
                note = str(row["품목비고"]).lower()
                keywords = ["dot", "old", "년식", "연식", "올드", "디오티", "24y", "25y", "26y"]
                is_match = any(kw in note for kw in keywords)
                
                res = []
                for col_name in row.index:
                    s = ""
                    if is_match:
                        s += "color: #0000FF; "
                    if col_name == "수량":
                        s += "font-weight: bold; "
                    res.append(s)
                return res

            styled_detail_display = detail_display.style.apply(highlight_detail_row, axis=1)

            st.dataframe(styled_detail_display, use_container_width=True, hide_index=True, height=550)
        else:
            st.warning("선택된 상세 정보가 없습니다.")

    render_bottom_nav()

def main():
    if not st.session_state["authenticated"]:
        render_login()
    else:
        render_main_app()

if __name__ == "__main__":
    main()