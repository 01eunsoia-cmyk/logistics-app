import streamlit as st
import pandas as pd
import numpy as np
import os
import pickle
from datetime import datetime
import pytz

# 페이지 설정
st.set_page_config(page_title="물량 관리 시스템", layout="wide")

# ==========================================
# 현장용 커스텀 CSS (글자 및 버튼 대폭 확대, 이모티콘 제거 대응)
# ==========================================
st.markdown("""
<style>
    /* 메인 버튼 글꼴 크기 및 높이 대폭 확대 */
    .element-container div.stButton > button {
        font-size: 26px !important;
        font-weight: bold !important;
        padding: 18px 20px !important;
        border-radius: 12px !important;
        margin-bottom: 12px !important;
    }
    
    /* 헤더 및 타이틀 글꼴 크기 확대 */
    h1, h2, h3 {
        font-weight: 800 !important;
    }
    
    /* 표(테이블) 텍스트 시인성 강화 */
    .stDataFrame, .stTable {
        font-size: 18px !important;
    }
</style>
""", unsafe_allow_html=True)

DATA_FILE = "processed_data.pkl"

def get_kst_now():
    tz = pytz.timezone('Asia/Seoul')
    return datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')

def process_uploaded_files(files_dict):
    df_c = None
    df_p = None
    df_b_list = []

    for name, file_obj in files_dict.items():
        try:
            xls = pd.ExcelFile(file_obj)
            for sheet_name in xls.sheet_names:
                temp_df = pd.read_excel(xls, sheet_name=sheet_name)
                file_text = str(name)
                sheet_text = str(sheet_name)

                if df_p is None and ("피킹" in file_text or "피킹" in sheet_text):
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

    # 피킹조회 파일 매칭 (로케이션, DOT, LI, SS)
    if df_p is not None and "일련번호" in df_p.columns:
        df_p["serial_key"] = df_p["일련번호"].astype(str).str.strip()
        df_b["serial_key"] = df_b["일련번호"].astype(str).str.strip()

        if "로케이션" in df_p.columns:
            loc_map = df_p.groupby("serial_key")["로케이션"].first().to_dict()
            df_b["로케이션_표시"] = df_b["serial_key"].map(loc_map)
        if "DOT" in df_p.columns:
            dot_map = df_p.groupby("serial_key")["DOT"].first().to_dict()
            df_b["디오티_표시"] = df_b["serial_key"].map(dot_map)
        if "LI" in df_p.columns:
            li_map = df_p.groupby("serial_key")["LI"].first().to_dict()
            df_b["LI_표시"] = df_b["serial_key"].map(li_map)
        if "SS" in df_p.columns:
            ss_map = df_p.groupby("serial_key")["SS"].first().to_dict()
            df_b["SS_표시"] = df_b["serial_key"].map(ss_map)

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

    df["출력상태"] = np.where(df[print_col].astype(str).str.upper().str.contains("Y"), "○", "X") if print_col else "X"
    df["차량번호_표시"] = df[car_col].fillna("-") if car_col else "-"
    df["점포명_표시"] = df[store_col].fillna("-") if store_col else "-"
    df["수량_표시"] = df[qty_col]
    df["주소_표시"] = df[addr_col]
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
        "raw_df": df
    }

    with open(DATA_FILE, "wb") as f:
        pickle.dump(processed_data, f)

    return processed_data


# --- UI 메인 세션 관리 ---
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    st.markdown("<h1 style='text-align: center; font-size: 36px;'>물량 관리 시스템 로그인</h1>", unsafe_allow_html=True)
    pw = st.text_input("비밀번호를 입력하세요", type="password")
    if st.button("로그인", use_container_width=True):
        if pw == "1234":  # 필요 시 비밀번호 변경
            st.session_state["logged_in"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    st.stop()

# --- 데이터 로드 ---
if not os.path.exists(DATA_FILE):
    st.warning("데이터 파일이 없습니다. 엑셀 파일을 업로드해 주세요.")
    uploaded_files = st.file_uploader("배차, 출고, 피킹조회 엑셀 파일 업로드", accept_multiple_files=True)
    if st.button("데이터 분석 및 저장", use_container_width=True):
        if uploaded_files:
            files_dict = {f.name: f for f in uploaded_files}
            try:
                process_uploaded_files(files_dict)
                st.success("데이터가 성공적으로 분석 및 저장되었습니다!")
                st.rerun()
            except Exception as e:
                st.error(f"오류 발생: {e}")
    st.stop()

with open(DATA_FILE, "rb") as f:
    data = pickle.load(f)

df = data["raw_df"]
updated_at = data["updated_at"]

# 상단 네비게이션
c_nav1, c_nav2, c_nav3 = st.columns([1, 1, 1])
with c_nav1:
    if st.button("홈", use_container_width=True):
        st.session_state.pop("selected_category", None)
        st.session_state.pop("selected_delivery_type", None)
        st.session_state.pop("selected_car_number", None)
        st.rerun()

with c_nav2:
    if st.button("뒤로가기", use_container_width=True):
        if "selected_car_number" in st.session_state:
            st.session_state.pop("selected_car_number")
        elif "selected_delivery_type" in st.session_state:
            st.session_state.pop("selected_delivery_type")
        elif "selected_category" in st.session_state:
            st.session_state.pop("selected_category")
        st.rerun()

with c_nav3:
    if st.button("로그아웃", use_container_width=True):
        st.session_state["logged_in"] = False
        st.rerun()

st.markdown(f"<p style='font-size:18px; font-weight:bold;'>최근 업데이트 일자 (KST): <span style='color:green;'>{updated_at}</span></p>", unsafe_allow_html=True)
st.markdown("---")

# --- 네비게이션 상태 파악 ---
cat = st.session_state.get("selected_category")
dtype = st.session_state.get("selected_delivery_type")
car_num = st.session_state.get("selected_car_number")

# 수량 합계 계산 (승용 / 화물 / 마케팅)
psr_df = df[df["최종구분"].str.upper().str.contains("승용|PSR", na=False)]
tbr_df = df[df["최종구분"].str.upper().str.contains("화물|TBR", na=False)]
mkt_df = df[~df["최종구분"].str.upper().str.contains("승용|화물|PSR|TBR", na=False)]

psr_total = int(psr_df["수량_표시"].sum())
tbr_total = int(tbr_df["수량_표시"].sum())
mkt_total = int(mkt_df["수량_표시"].sum())

# 1단계: 승용 / 화물 / 마케팅 선택
if not cat:
    st.markdown("<h1 style='text-align: center; font-size: 38px;'>구분 선택</h1>", unsafe_allow_html=True)
    st.write("")
    
    if st.button(f"승용 ({psr_total:,}개)", use_container_width=True, key="cat_psr"):
        st.session_state["selected_category"] = "승용"
        st.rerun()
        
    if st.button(f"화물 ({tbr_total:,}개)", use_container_width=True, key="cat_tbr"):
        st.session_state["selected_category"] = "화물"
        st.rerun()
        
    if st.button(f"마케팅 ({mkt_total:,}개)", use_container_width=True, key="cat_mkt"):
        st.session_state["selected_category"] = "마케팅"
        st.rerun()

# 2단계: 배송 유형 선택 (당/익/택배/미배차)
elif cat and not dtype:
    if cat == "승용":
        cat_df = psr_df
    elif cat == "화물":
        cat_df = tbr_df
    else:
        cat_df = mkt_df

    st.markdown(f"<h1 style='text-align: center; font-size: 36px;'>{cat} - 배송 유형 선택</h1>", unsafe_allow_html=True)
    
    types = ["당", "익", "택배", "미배차"]
    cols = st.columns(2)
    for idx, t in enumerate(types):
        t_qty = int(cat_df[cat_df["배송유형"] == t]["수량_표시"].sum())
        with cols[idx % 2]:
            if st.button(f"{t} ({t_qty:,}개)", use_container_width=True, key=f"dtype_{t}"):
                st.session_state["selected_delivery_type"] = t
                st.rerun()

# 3단계: 차량 번호 목록 또는 세부 리스트
elif cat and dtype and not car_num:
    if cat == "승용":
        sub_df = psr_df[psr_df["배송유형"] == dtype].copy()
    elif cat == "화물":
        sub_df = tbr_df[tbr_df["배송유형"] == dtype].copy()
    else:
        sub_df = mkt_df[mkt_df["배송유형"] == dtype].copy()

    st.markdown(f"<h1 style='text-align: center; font-size: 34px;'>{cat} > {dtype} 차량 선택</h1>", unsafe_allow_html=True)

    grouped = sub_df.groupby("차량번호_표시")
    for car, group in grouped:
        car_sum = int(group["수량_표시"].sum())
        if st.button(f"차량: {car} ({car_sum:,}개)", use_container_width=True, key=f"car_{car}"):
            st.session_state["selected_car_number"] = car
            st.rerun()

# 4단계: 상세 데이터 뷰
elif cat and dtype and car_num:
    if cat == "승용":
        sub_df = psr_df[(psr_df["배송유형"] == dtype) & (psr_df["차량번호_표시"] == car_num)]
    elif cat == "화물":
        sub_df = tbr_df[(tbr_df["배송유형"] == dtype) & (tbr_df["차량번호_표시"] == car_num)]
    else:
        sub_df = mkt_df[(mkt_df["배송유형"] == dtype) & (mkt_df["차량번호_표시"] == car_num)]

    st.markdown(f"<h2 style='font-size: 30px;'>리스트 상세 뷰: {cat} > {dtype} > 차량 {car_num}</h2>", unsafe_allow_html=True)
    
    view_df = sub_df[["로케이션_표시", "수량_표시", "디오티_표시", "패턴_표시", "사이즈_표시", "LI_표시", "SS_표시"]].copy()
    view_df.columns = ["로케이션", "수량", "디오티", "패턴", "사이즈", "LI", "SS"]
    
    st.dataframe(view_df, use_container_width=True)