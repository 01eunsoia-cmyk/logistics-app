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

    df_b["key_addr"] = df_b["출고처명(거래처)"].astype(str).str.strip() + "_" + df_b["배송처주소"].astype(str).str.strip()
    df_c["key_addr"] = df_c["출고처명(거래처)"].astype(str).str.strip() + "_" + df_c["배송처주소"].astype(str).str.strip()

    if "피킹지출력" in df_c.columns:
        map_addr = df_c.groupby("key_addr")["피킹지출력"].first().to_dict()
        df_b["피킹지출력"] = df_b["key_addr"].map(map_addr)

    df = df_b.copy()

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
    driver_col = find_col(["차량기사", "기사명", "기사"])
    phone_col = find_col(["연락처", "전화번호", "기사연락처"])
    courier_col = find_col(["택배구분", "택배분류", "택배"])
    code_col = find_col(["품목코드", "코드", "자재코드", "SKU"])
    size_col = find_col(["SIZE", "규격", "사이즈"])
    pttn_col = find_col(["PTTN", "패턴"])
    note_col = find_col(["비고", "특이사항", "메모", "품목비고", "요청사항"])
    print_col = find_col(["피킹지출력", "출력", "피킹지"])
    time_col = find_col(["출고일시", "출고시간", "일시"])
    plan_date_col = find_col(["예정일자", "예정일"])
    store_col = find_col(["출고처명(거래처)", "점포명", "거래처"])

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
    
    kst_today = datetime.now(timezone(timedelta(hours=9))).date()

    plan_date_parsed = pd.to_datetime(df[plan_date_col], errors="coerce").dt.date if plan_date_col else pd.Series(pd.NaT, index=df.index)
    time_parsed = pd.to_datetime(df[time_col], errors="coerce").dt.date if time_col else pd.Series(pd.NaT, index=df.index)

    is_plan_tomorrow = plan_date_parsed.apply(lambda d: d > kst_today if pd.notna(d) else False)
    is_time_tomorrow = time_parsed.apply(lambda d: d > kst_today if pd.notna(d) else False)
    is_tomorrow = is_plan_tomorrow | is_time_tomorrow

    gyeongsang_kw = ["부산", "대구", "울산", "경남", "경북", "경상", "밀양", "양산", "창원", "김해", "포항"]
    is_gyeongsang = s_addr.str.contains("|".join(gyeongsang_kw), na=False)
    is_express = s_note.str.contains("당착|오후", na=False)
    is_dang = is_gyeongsang | is_express
    is_jeju = s_addr.str.contains("제주", na=False) | s_note.str.contains("제주", na=False)
    is_courier = s_driver.eq("경동택배") & s_courier.eq("택배")
    is_blank = time_str.eq("") | time_str.str.lower().isin(["nan", "none", "nat"])

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

    processed_data = {
        "updated_at": get_kst_now(),
        "raw_df": df,
        "raw_inventory_df": df_i,
        "raw_picking_df": df_p
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
    df_p = data.get("raw_picking_df", None) if data else None

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

    # --- 1단계: 구분 선택 홈 ---
    if current_page == "home_category":
        st.markdown("<h2 style='text-align: center;'>구분 선택</h2>", unsafe_allow_html=True)
        col1, col2, col3, col4, col5 = st.columns(5)
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
        with col5:
            if st.button("📋\n\n피킹조회", use_container_width=True, key="cat_picking"):
                st.session_state["selected_category"] = "피킹조회"
                st.session_state["page_stack"].append("picking_view")
                st.rerun()

    # --- 통합 재고조회 뷰 ---
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
            if c_dot: display_inv["DOT"] = inv_df[c_dot].fillna("-").astype(str)
            if c_avail: display_inv["가용재고수량"] = pd.to_numeric(inv_df[c_avail], errors="coerce").fillna(0).astype(int)
            if c_total: display_inv["전체재고수량"] = pd.to_numeric(inv_df[c_total], errors="coerce").fillna(0).astype(int)
            if c_item_type: display_inv["품목구분"] = inv_df[c_item_type].fillna("-")
            if c_storage_note: display_inv["적치비고"] = inv_df[c_storage_note].fillna("-")
            if c_date: display_inv["입고일자"] = inv_df[c_date].fillna("-")
            if c_container: display_inv["컨테이너번호"] = inv_df[c_container].fillna("-")

            st.dataframe(display_inv, use_container_width=True, hide_index=True, height=550)

    # --- 피킹조회 뷰 ---
    elif current_page == "picking_view":
        st.markdown("<h3 style='text-align: center;'>📋 피킹조회 현황</h3>", unsafe_allow_html=True)
        if df_p is None or df_p.empty:
            st.warning("피킹 파일(.xls)이 업로드되지 않았거나 데이터가 없습니다.")
        else:
            pick_df = df_p.copy()
            st.dataframe(pick_df, use_container_width=True, hide_index=True, height=550)

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
        qty_courier = int(cat_df[cat_df["배송유형"] == "택배"]["수량_표시"].sum())

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

        if cat == "승용":
            sub_df = df[df["최종구분"].str.upper().str.contains("승용|PSR", na=False)].copy()
        elif cat == "화물":
            sub_df = df[df["최종구분"].str.upper().str.contains("화물|TBR", na=False)].copy()
        else:
            sub_df = df[~df["최종구분"].str.upper().str.contains("승용|화물|PSR|TBR", na=False)].copy()

        if sub == "택배":
            sub_df = sub_df[sub_df["배송유형"] == "택배"]
        elif sub in ["익", "제주"]:
            sub_df = sub_df[sub_df["배송유형"].isin(["익", "제주"])]
        else:
            sub_df = sub_df[sub_df["배송유형"] == sub]

        total_qty = int(sub_df["수량_표시"].sum())
        st.metric("📦 총 수량", f"{total_qty:,} 개")
        st.write("---")

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

        st.dataframe(show_df.drop(columns=["차량그룹키", "배송유형"]), use_container_width=True, hide_index=True, height=550)

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

    # --- 4단계: 세부사항 뷰 (피킹조회 데이터 완벽 연동 복구) ---
    elif current_page == "detail_view":
        row_data = st.session_state.get("selected_detail_row", None)

        st.markdown("### 🔍 세부사항 뷰 (피킹 품목 리스트)")

        if row_data is not None:
            st.info(f"**점포명:** {row_data['점포명']} | **차량번호:** {row_data['차량번호']} | **기사명:** {row_data['기사명']} | **연락처:** {row_data['연락처']} | **주소:** {row_data['주소']}")

            # 피킹조회 원본 파일(df_p)에서 해당 점포(출고처명/배송처주소)의 품목 리스트를 직접 추출하여 표시
            if df_p is not None and not df_p.empty:
                p_store_col = [c for c in df_p.columns if "출고처" in c or "거래처" in c or "점포" in c]
                p_addr_col = [c for c in df_p.columns if "주소" in c or "배송처" in c]
                
                if p_store_col and p_addr_col:
                    p_match = df_p[
                        (df_p[p_store_col[0]].astype(str).str.strip() == str(row_data["점포명"]).strip()) &
                        (df_p[p_addr_col[0]].astype(str).str.strip() == str(row_data["주소"]).strip())
                    ].copy()
                else:
                    p_match = pd.DataFrame()

                if not p_match.empty:
                    def get_p_col(kw_list):
                        for kw in kw_list:
                            for col in p_match.columns:
                                if kw in str(col):
                                    return col
                        return None

                    c_loc = get_p_col(["로케이션"])
                    c_qty = get_p_col(["피킹수량", "수량"])
                    c_dot = get_p_col(["DOT"])
                    c_code = get_p_col(["품목코드"])
                    c_size = get_p_col(["SIZE", "사이즈"])
                    c_li = get_p_col(["LI"])
                    c_ss = get_p_col(["SS"])
                    c_note = get_p_col(["품목비고", "비고"])

                    detail_display = pd.DataFrame()
                    detail_display["로케이션"] = p_match[c_loc].fillna("-") if c_loc else "-"
                    detail_display["수량"] = pd.to_numeric(p_match[c_qty], errors="coerce").fillna(0) if c_qty else 0
                    detail_display["디오티"] = p_match[c_dot].fillna("-").astype(str) if c_dot else "-"
                    detail_display["상품코드"] = p_match[c_code].fillna("-") if c_code else "-"
                    detail_display["사이즈"] = p_match[c_size].fillna("-") if c_size else "-"
                    detail_display["LI"] = p_match[c_li].fillna("-") if c_li else "-"
                    detail_display["SS"] = p_match[c_ss].fillna("-") if c_ss else "-"
                    detail_display["품목비고"] = p_match[c_note].fillna("-") if c_note else "-"

                    detail_display = detail_display.sort_values(by=["로케이션", "사이즈"], ascending=[False, False]).reset_index(drop=True)
                    st.dataframe(detail_display, use_container_width=True, hide_index=True, height=550)
                else:
                    st.warning("피킹조회 데이터에서 해당 점포의 일치하는 상세 품목을 찾지 못했습니다. 배차 기준 데이터를 표시합니다.")
                    # Fallback to 배차 기준 행 표시
                    match_cond = (
                        (df["점포명_표시"].astype(str).str.strip() == str(row_data["점포명"]).strip()) &
                        (df["주소_표시"].astype(str).str.strip() == str(row_data["주소"]).strip())
                    )
                    matching_items = df[match_cond].copy()
                    detail_display = pd.DataFrame({
                        "로케이션": matching_items.get("로케이션_표시", "-"),
                        "수량": matching_items["수량_표시"],
                        "디오티": matching_items.get("디오티_표시", "-"),
                        "상품코드": matching_items.get("상품코드_표시", "-"),
                        "사이즈": matching_items.get("사이즈_표시", "-"),
                        "LI": matching_items.get("LI_표시", "-"),
                        "SS": matching_items.get("SS_표시", "-"),
                        "품목비고": matching_items.get("품목비고", "-")
                    })
                    st.dataframe(detail_display, use_container_width=True, hide_index=True, height=550)
            else:
                st.warning("피킹조회 파일이 업로드되지 않았습니다.")
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