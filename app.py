from datetime import datetime
from io import BytesIO
import json
import os
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st


# =========================================================
# 기본 설정
# =========================================================
st.set_page_config(
    page_title="물류/출고 조회 시스템",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# 모바일에서 상단 네비게이션 3개 버튼이 세로로 떨어지지 않도록 고정
st.markdown(
    """
    <style>
    .st-key-top_nav [data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 0.45rem !important;
    }
    .st-key-top_nav [data-testid="column"] {
        flex: 1 1 0 !important;
        width: 33.333% !important;
        min-width: 0 !important;
    }
    .st-key-top_nav button {
        width: 100% !important;
        min-height: 3rem !important;
        padding: 0.2rem 0.35rem !important;
        font-size: 1.15rem !important;
    }
    @media (max-width: 768px) {
        .st-key-top_nav [data-testid="stHorizontalBlock"] {
            gap: 0.35rem !important;
        }
        .st-key-top_nav button {
            min-height: 3.2rem !important;
        }
        [data-testid="stDataFrame"] {
            font-size: 0.82rem !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = Path(__file__).resolve().parent
# 내부 서버/Docker에서는 APP_DATA_DIR을 영구 볼륨 경로로 지정 가능
DATA_DIR = Path(os.getenv("APP_DATA_DIR", str(BASE_DIR / "data")))
DATA_DIR.mkdir(parents=True, exist_ok=True)

SOURCE_PATHS = {
    "출고": DATA_DIR / "latest_shipping.xlsx",
    "피킹조회": DATA_DIR / "latest_picking.xlsx",
    "배차처리": DATA_DIR / "latest_dispatch.xlsx",
}
META_FILE = DATA_DIR / "source_meta.json"

SOURCE_REQUIRED_COLUMNS = {
    "출고": {
        "출고번호",
        "피킹지출력",
        "출고일시",
        "출고처명(거래처)",
        "배송처주소",
        "총수량",
        "차량번호",
        "택배구분",
        "예상도착일시",
    },
    "피킹조회": {
        "타이어구분",
        "피킹일시",
        "출고처명(거래처)",
        "배송처주소",
        "품목코드",
        "SIZE",
        "PTTN",
        "피킹수량",
        "로케이션",
        "차량번호",
        "일련번호",
        "예상도착일시",
    },
    "배차처리": {
        "타이어구분",
        "예정일자",
        "출고일시",
        "출고처명(거래처)",
        "배송처주소",
        "배차수량",
        "차량번호",
        "품목코드",
        "SIZE",
        "PTTN",
        "택배구분",
        "일련번호",
    },
}

CATEGORY_MAP = {
    "PSR": "승용",
    "TBR": "화물",
    "마케팅": "마케팅",
}


def get_secret(name: str, fallback: str) -> str:
    try:
        return str(st.secrets[name])
    except Exception:
        return fallback


ADMIN_PASSWORD = get_secret("ADMIN_PASSWORD", "bstk2026!")
GENERAL_PASSWORD = get_secret("GENERAL_PASSWORD", "1010")


# =========================================================
# 세션 상태
# =========================================================
def init_session_state():
    defaults = {
        "authenticated": False,
        "role": None,
        "step": "home",  # home, sub, list, detail
        "main_category": "",
        "sub_category": "",
        "selected_group": None,
        "table_nonce": 0,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_session_state()


# =========================================================
# 파일 읽기 / 저장
# =========================================================
def clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    # Excel에서 문자열 식별자가 123.0 형태로 읽힌 경우 정리
    if text.endswith(".0"):
        head = text[:-2]
        if head.isdigit():
            return head
    return text


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [clean_text(col) for col in df.columns]

    # 배차처리 원본의 실제 헤더 오타(차랑번호)를 내부 표준명으로 통일
    if "차랑번호" in df.columns and "차량번호" not in df.columns:
        df = df.rename(columns={"차랑번호": "차량번호"})

    # 피킹조회 첫 번째 빈 헤더 등 불필요 열 제거
    drop_cols = [c for c in df.columns if not c or c.startswith("Unnamed:")]
    if drop_cols:
        df = df.drop(columns=drop_cols, errors="ignore")

    # 완전히 빈 행 제거
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def read_csv_with_fallback(source) -> pd.DataFrame:
    last_error = None
    for encoding in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            if isinstance(source, (str, Path)):
                return pd.read_csv(source, encoding=encoding)
            return pd.read_csv(BytesIO(source), encoding=encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
    if last_error:
        raise last_error
    raise ValueError("CSV 파일을 읽을 수 없습니다.")


def dataframe_from_bytes(file_bytes: bytes, suffix: str) -> pd.DataFrame:
    if suffix == ".xlsx":
        df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl", dtype=object)
    elif suffix == ".xls":
        try:
            df = pd.read_excel(BytesIO(file_bytes), engine="xlrd", dtype=object)
        except ImportError as exc:
            raise ValueError(
                ".xls 파일을 읽으려면 xlrd가 필요합니다. requirements.txt를 같이 배포해주세요."
            ) from exc
    elif suffix == ".csv":
        df = read_csv_with_fallback(file_bytes)
    else:
        raise ValueError("지원 형식은 xls, xlsx, csv입니다.")
    return normalize_columns(df)


def detect_source_type(df: pd.DataFrame) -> str:
    cols = set(df.columns)

    # 특징 열 우선 판별
    if {"출고번호", "피킹지출력"}.issubset(cols):
        return "출고"
    if {"피킹수량", "로케이션", "일련번호"}.issubset(cols):
        return "피킹조회"
    if {"배차수량", "일련번호"}.issubset(cols):
        return "배차처리"

    raise ValueError(
        "파일 종류를 판별할 수 없습니다. 출고 / 피킹조회 / 배차처리 원본 파일인지 확인해주세요."
    )


def validate_source_df(source_type: str, df: pd.DataFrame):
    missing = sorted(SOURCE_REQUIRED_COLUMNS[source_type] - set(df.columns))
    if missing:
        raise ValueError(
            f"{source_type} 파일에 필요한 열이 없습니다: " + ", ".join(missing)
        )


def load_metadata() -> dict:
    if not META_FILE.exists():
        return {}
    try:
        data = json.loads(META_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_metadata(meta: dict):
    temp = DATA_DIR / ".source_meta.tmp.json"
    temp.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(META_FILE)


def save_upload_batch(uploaded_files):
    """업로드한 종류만 교체. 올리지 않은 종류는 기존 마지막 파일을 그대로 유지."""
    if not uploaded_files:
        raise ValueError("업데이트할 파일을 선택해주세요.")

    parsed = {}

    # 전부 먼저 메모리에서 검증 → 하나라도 이상하면 기존 파일은 건드리지 않음
    for uploaded in uploaded_files:
        suffix = Path(uploaded.name).suffix.lower()
        if suffix not in {".xls", ".xlsx", ".csv"}:
            raise ValueError(f"지원하지 않는 형식입니다: {uploaded.name}")

        df = dataframe_from_bytes(uploaded.getvalue(), suffix)
        source_type = detect_source_type(df)
        validate_source_df(source_type, df)

        if source_type in parsed:
            raise ValueError(
                f"{source_type} 파일이 한 번에 2개 이상 선택되었습니다. 종류별로 1개만 올려주세요."
            )
        parsed[source_type] = (df, uploaded.name, suffix)

    updated_at = datetime.now(ZoneInfo("Asia/Seoul")).strftime("%Y-%m-%d %H:%M:%S")
    meta = load_metadata()

    # 각 원본을 내부적으로 xlsx로 통일 저장
    for source_type, (df, original_name, suffix) in parsed.items():
        target = SOURCE_PATHS[source_type]
        temp = DATA_DIR / f".{target.stem}.tmp.xlsx"
        df.to_excel(temp, index=False, engine="openpyxl")
        temp.replace(target)

        meta[source_type] = {
            "original_filename": original_name,
            "updated_at": updated_at,
            "row_count": int(len(df)),
            "original_format": suffix.lstrip("."),
        }

    save_metadata(meta)
    return list(parsed.keys()), meta


def load_source(source_type: str) -> pd.DataFrame:
    path = SOURCE_PATHS[source_type]
    if not path.exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, engine="openpyxl", dtype=object)
        df = normalize_columns(df)
        validate_source_df(source_type, df)
        return df
    except Exception as exc:
        st.error(f"저장된 {source_type} 데이터를 읽지 못했습니다: {exc}")
        return pd.DataFrame()


# =========================================================
# 3개 원본 → 조회용 데이터 결합
# =========================================================
def normalize_text_columns(df: pd.DataFrame, columns) -> pd.DataFrame:
    df = df.copy()
    for col in columns:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].map(clean_text)
    return df


def first_nonempty(series):
    for value in series:
        text = clean_text(value)
        if text:
            return text
    return ""


def join_unique(series) -> str:
    seen = []
    for value in series:
        text = clean_text(value)
        if text and text not in seen:
            seen.append(text)
    return " / ".join(seen)


def parse_date(value):
    text = clean_text(value)
    if not text:
        return None
    ts = pd.to_datetime(text, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.date()


def build_unified_data(shipping: pd.DataFrame, picking: pd.DataFrame, dispatch: pd.DataFrame):
    # 배차처리가 일일 계획의 기준 데이터. 미배차 항목도 여기 존재함.
    if dispatch.empty:
        return pd.DataFrame()

    dispatch = normalize_text_columns(
        dispatch,
        [
            "타이어구분",
            "예정일자",
            "출고일시",
            "출고처명(거래처)",
            "배송처주소",
            "차량번호",
            "차량기사",
            "연락처",
            "품목코드",
            "SIZE",
            "PTTN",
            "품목비고",
            "택배구분",
            "일련번호",
            "고객사명",
            "배송지역",
            "DOCK번호",
        ],
    )

    dispatch["배차수량"] = pd.to_numeric(dispatch["배차수량"], errors="coerce").fillna(0)
    if "총수량" in dispatch.columns:
        dispatch["총수량"] = pd.to_numeric(dispatch["총수량"], errors="coerce").fillna(0)
    else:
        dispatch["총수량"] = 0

    # -------------------------
    # 피킹 상세: 일련번호 기준 결합
    # 같은 일련번호가 DOT/로케이션 때문에 여러 줄이면 합쳐서 표시
    # -------------------------
    pick_agg = pd.DataFrame()
    if not picking.empty:
        picking = normalize_text_columns(
            picking,
            [
                "일련번호",
                "로케이션",
                "DOT",
                "예상도착일시",
                "도착완료일시",
                "운송장번호",
            ],
        )
        picking["피킹수량"] = pd.to_numeric(picking["피킹수량"], errors="coerce").fillna(0)

        pick_agg = (
            picking.groupby("일련번호", dropna=False, as_index=False)
            .agg(
                로케이션=("로케이션", join_unique),
                DOT=("DOT", join_unique),
                피킹수량=("피킹수량", "sum"),
                피킹_예상도착일시=("예상도착일시", first_nonempty),
                도착완료일시=("도착완료일시", first_nonempty),
                운송장번호=("운송장번호", join_unique),
            )
        )
        dispatch = dispatch.merge(pick_agg, on="일련번호", how="left")

    for col in ["로케이션", "DOT", "피킹_예상도착일시", "도착완료일시", "운송장번호"]:
        if col not in dispatch.columns:
            dispatch[col] = ""
        dispatch[col] = dispatch[col].fillna("").map(clean_text)
    if "피킹수량" not in dispatch.columns:
        dispatch["피킹수량"] = 0
    dispatch["피킹수량"] = pd.to_numeric(dispatch["피킹수량"], errors="coerce").fillna(0)

    # -------------------------
    # 출고 파일: 출력여부 + 예상도착일시 결합
    # 실제 원본 기준으로 출고처/주소/차량/출고일시 조합이 배차·피킹과 일치함
    # -------------------------
    ship_summary = pd.DataFrame()
    ship_key = ["출고처명(거래처)", "배송처주소", "차량번호", "출고일시"]
    if not shipping.empty:
        shipping = normalize_text_columns(
            shipping,
            ship_key + ["피킹지출력", "예상도착일시", "택배구분"],
        )

        def print_state(series):
            values = {clean_text(v).upper() for v in series if clean_text(v)}
            if "N" in values:
                return "미출력"
            if "Y" in values:
                return "출력"
            return "미출력"

        ship_summary = (
            shipping.groupby(ship_key, dropna=False, as_index=False)
            .agg(
                출력여부=("피킹지출력", print_state),
                출고_예상도착일시=("예상도착일시", first_nonempty),
            )
        )
        dispatch = dispatch.merge(ship_summary, on=ship_key, how="left")

    if "출력여부" not in dispatch.columns:
        dispatch["출력여부"] = "미출력"
    dispatch["출력여부"] = dispatch["출력여부"].fillna("미출력").map(clean_text)

    if "출고_예상도착일시" not in dispatch.columns:
        dispatch["출고_예상도착일시"] = ""
    dispatch["출고_예상도착일시"] = dispatch["출고_예상도착일시"].fillna("").map(clean_text)

    # 출고 파일 값을 우선, 없으면 피킹조회 예상도착 사용
    dispatch["예상도착일시"] = dispatch.apply(
        lambda r: clean_text(r.get("출고_예상도착일시"))
        or clean_text(r.get("피킹_예상도착일시")),
        axis=1,
    )

    # 화면용 표준 컬럼
    dispatch["분류"] = dispatch["타이어구분"].map(CATEGORY_MAP).fillna("")
    dispatch["점포명"] = dispatch["출고처명(거래처)"].map(clean_text)
    dispatch["주소"] = dispatch["배송처주소"].map(clean_text)
    dispatch["상품코드"] = dispatch["품목코드"].map(clean_text)
    dispatch["사이즈"] = dispatch["SIZE"].map(clean_text)
    dispatch["패턴"] = dispatch["PTTN"].map(clean_text)
    dispatch["기사명"] = dispatch["차량기사"].map(clean_text)
    dispatch["연락처"] = dispatch["연락처"].map(clean_text)
    dispatch["수량"] = dispatch["배차수량"].where(dispatch["배차수량"] > 0, dispatch["총수량"])

    # 원본에 해당 열이 생기면 자동으로 화면에 연결하고, 없으면 공란으로 둠
    barcode_col = next((c for c in dispatch.columns if "바코드" in clean_text(c)), None)
    winter_col = next(
        (c for c in dispatch.columns if "윈터" in clean_text(c) or "winter" in clean_text(c).lower()),
        None,
    )
    dispatch["바코드"] = dispatch[barcode_col].map(clean_text) if barcode_col else ""
    dispatch["윈터"] = dispatch[winter_col].map(clean_text) if winter_col else ""

    def classify_status(row):
        vehicle = clean_text(row.get("차량번호"))
        driver = clean_text(row.get("차량기사"))
        parcel = clean_text(row.get("택배구분"))
        departure = clean_text(row.get("출고일시"))

        # 업무 기준: 출고일시가 비어 있으면 무조건 미배차
        if not departure:
            return "미배차"

        # 임시 차량은 출고일시가 같은 건을 한 차량으로 취급
        if "임의차량" in driver or "9999" in vehicle or not vehicle:
            return "배차중"

        if parcel == "택배" or "택배" in vehicle:
            return "택배"

        base_date = parse_date(row.get("예정일자")) or parse_date(departure)
        arrival_date = parse_date(row.get("예상도착일시"))
        if base_date and arrival_date:
            return "당착" if arrival_date <= base_date else "익착"
        return "미정"

    dispatch["상태"] = dispatch.apply(classify_status, axis=1)

    def display_vehicle(row):
        if row["상태"] == "미배차":
            return "미배차"
        if row["상태"] == "배차중":
            return "배차중"
        return clean_text(row.get("차량번호")) or "-"

    def vehicle_group_key(row):
        # 배차중은 차량번호 대신 출고일시를 차량 식별키로 사용
        if row["상태"] == "배차중":
            return f"배차중|{clean_text(row.get('출고일시'))}"
        if row["상태"] == "미배차":
            return f"미배차|{clean_text(row.get('점포명'))}|{clean_text(row.get('주소'))}"
        return f"{clean_text(row.get('차량번호'))}|{clean_text(row.get('출고일시'))}"

    dispatch["차량정보"] = dispatch.apply(display_vehicle, axis=1)
    dispatch["_차량그룹키"] = dispatch.apply(vehicle_group_key, axis=1)
    # 현재 원본의 피킹지출력 Y를 피킹내림 표시(○)로 사용
    dispatch["피킹내림"] = dispatch["출력여부"].map(lambda v: "○" if clean_text(v) == "출력" else "")

    # 지원 대상 3개 분류만 화면에 노출
    dispatch = dispatch[dispatch["분류"].isin(["승용", "화물", "마케팅"])].copy()

    # 화면 표시 안정화
    for col in [
        "분류",
        "상태",
        "출력여부",
        "차량번호",
        "점포명",
        "주소",
        "패턴",
        "사이즈",
        "상품코드",
        "로케이션",
        "DOT",
        "택배구분",
        "예정일자",
        "출고일시",
        "예상도착일시",
        "일련번호",
        "품목비고",
        "운송장번호",
        "기사명",
        "연락처",
        "차량정보",
        "피킹내림",
        "바코드",
        "윈터",
    ]:
        if col not in dispatch.columns:
            dispatch[col] = ""
        dispatch[col] = dispatch[col].fillna("").map(clean_text)

    # 정수 수량은 정수로 표시
    if not dispatch.empty and (dispatch["수량"] % 1 == 0).all():
        dispatch["수량"] = dispatch["수량"].astype(int)

    return dispatch.reset_index(drop=True)


# =========================================================
# 로그인 / 공통 UI
# =========================================================
def render_login():
    st.markdown("### 📦 출고/배차 조회 시스템")
    st.caption("사내 업무용 · 중앙관리 방식")
    st.markdown("")

    mode = st.radio("접속 모드", ["일반모드", "관리자모드"], horizontal=True)
    password = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")

    if st.button("로그인", use_container_width=True, type="primary"):
        expected = GENERAL_PASSWORD if mode == "일반모드" else ADMIN_PASSWORD
        if password == expected:
            st.session_state.authenticated = True
            st.session_state.role = "general" if mode == "일반모드" else "admin"
            st.session_state.step = "home"
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")


def logout():
    st.session_state.authenticated = False
    st.session_state.role = None
    st.session_state.step = "home"
    st.session_state.main_category = ""
    st.session_state.sub_category = ""
    st.session_state.selected_group = None
    st.rerun()


def render_top_nav(show_back=True):
    # 모바일에서도 새로고침 / 홈 / 뒤로가기를 반드시 한 줄에 표시
    with st.container(key="top_nav"):
        col1, col2, col3 = st.columns(3, gap="small")
        with col1:
            if st.button("🔄", help="최신 데이터 다시 불러오기", use_container_width=True):
                st.rerun()
        with col2:
            if st.button("🏠", help="홈으로", use_container_width=True):
                st.session_state.step = "home"
                st.session_state.selected_group = None
                st.session_state.table_nonce += 1
                st.rerun()
        with col3:
            if st.button("⬅️", help="뒤로가기", use_container_width=True):
                if st.session_state.step == "detail":
                    st.session_state.step = "list"
                    st.session_state.selected_group = None
                    st.session_state.table_nonce += 1
                    st.rerun()
                elif st.session_state.step == "list":
                    st.session_state.step = "sub"
                    st.rerun()
                elif st.session_state.step == "sub":
                    st.session_state.step = "home"
                    st.rerun()
                else:
                    # 홈에서 뒤로가기는 로그인 화면으로 돌아감
                    logout()
    st.markdown("---")


def render_data_status():
    meta = load_metadata()
    parts = []
    for source_type in ["출고", "피킹조회", "배차처리"]:
        item = meta.get(source_type, {})
        if SOURCE_PATHS[source_type].exists():
            parts.append(
                f"{source_type}: {item.get('original_filename', SOURCE_PATHS[source_type].name)}"
            )
        else:
            parts.append(f"{source_type}: 없음")
    st.caption(" · ".join(parts))


def render_admin_panel():
    if st.session_state.role != "admin":
        return

    meta = load_metadata()
    missing_count = sum(1 for p in SOURCE_PATHS.values() if not p.exists())

    with st.expander("⚙️ 관리자 데이터 관리", expanded=missing_count > 0):
        st.write(
            "출고 / 피킹조회 / 배차처리 파일을 올리면 파일명과 관계없이 열 구조로 자동 구분합니다. "
            "새로 올린 종류만 교체되고, 올리지 않은 종류는 마지막 파일을 계속 사용합니다."
        )

        for source_type in ["출고", "피킹조회", "배차처리"]:
            item = meta.get(source_type, {})
            if SOURCE_PATHS[source_type].exists():
                st.caption(
                    f"✅ {source_type} · {item.get('original_filename', '-')} · "
                    f"{item.get('row_count', 0):,}행 · {item.get('updated_at', '-')}"
                )
            else:
                st.caption(f"⚠️ {source_type} · 아직 업로드되지 않음")

        uploaded_files = st.file_uploader(
            "원본 파일 선택",
            type=["xls", "xlsx", "csv"],
            accept_multiple_files=True,
            help="1개만 교체해도 되고, 3개를 한 번에 올려도 됩니다.",
        )

        if uploaded_files:
            st.caption("선택: " + " / ".join(f.name for f in uploaded_files))
            if st.button("선택한 파일로 업데이트", type="primary", use_container_width=True):
                try:
                    updated_types, _ = save_upload_batch(uploaded_files)
                    st.session_state.selected_group = None
                    st.success("업데이트 완료: " + ", ".join(updated_types))
                    st.rerun()
                except Exception as exc:
                    st.error(f"업로드 실패: {exc}")

        st.info(
            "현재 원본 기준: PSR=승용, TBR=화물, 마케팅=마케팅출고 / "
            "출고일시 공란=미배차 / 차량기사에 임의차량 또는 차량번호에 9999=배차중 / "
            "배차중은 출고일시가 같으면 한 차량으로 묶습니다. "
            "예상도착일이 예정일과 같으면 당착, 이후면 익착입니다."
        )


# =========================================================
# 조회용 보조 함수
# =========================================================
def format_qty(value) -> str:
    try:
        value = float(value)
        if value.is_integer():
            return f"{int(value):,}"
        return f"{value:,.2f}".rstrip("0").rstrip(".")
    except Exception:
        return clean_text(value)


def group_for_list(filtered_df: pd.DataFrame) -> pd.DataFrame:
    if filtered_df.empty:
        return pd.DataFrame()

    group_cols = [
        "_차량그룹키",
        "차량정보",
        "기사명",
        "연락처",
        "점포명",
        "주소",
        "상태",
        "출력여부",
        "출고일시",
        "예정일자",
        "예상도착일시",
    ]

    grouped = (
        filtered_df.groupby(group_cols, dropna=False, as_index=False)
        .agg(
            수량=("수량", "sum"),
            가지수=("상품코드", "nunique"),
            바코드=("바코드", join_unique),
            윈터=("윈터", join_unique),
            비고=("품목비고", join_unique),
        )
        .sort_values(["_차량그룹키", "점포명"], na_position="last")
        .reset_index(drop=True)
    )
    grouped["피킹내림"] = grouped["출력여부"].map(
        lambda v: "○" if clean_text(v) == "출력" else ""
    )
    return grouped


def rows_for_selected_group(df: pd.DataFrame, group: dict) -> pd.DataFrame:
    result = df.copy()
    # 한 배송처 행의 상세를 조회. 배차중은 동일 출고일시를 같은 차량키로 사용함.
    for col in ["_차량그룹키", "점포명", "주소", "상태", "출고일시"]:
        result = result[result[col].map(clean_text) == clean_text(group.get(col))]
    return result.copy()


def natural_location_key(value: str) -> str:
    """A-2-10 같은 로케이션도 숫자 자릿수에 영향 없이 정렬되게 만든 키."""
    import re

    text = clean_text(value).upper()
    return re.sub(r"\d+", lambda m: f"{int(m.group()):010d}", text)


def build_picking_detail(picking: pd.DataFrame, selected_dispatch: pd.DataFrame) -> pd.DataFrame:
    """선택 배송처의 원본 피킹행을 보존해 로케이션/DOT별 실제 피킹수량을 보여줌."""
    if selected_dispatch.empty:
        return pd.DataFrame()

    pick = picking.copy() if not picking.empty else pd.DataFrame()
    if not pick.empty:
        pick = normalize_text_columns(
            pick,
            [
                "일련번호",
                "출고처명(거래처)",
                "배송처주소",
                "품목코드",
                "SIZE",
                "PTTN",
                "로케이션",
                "DOT",
            ],
        )
        if "피킹수량" not in pick.columns:
            pick["피킹수량"] = 0
        pick["피킹수량"] = pd.to_numeric(pick["피킹수량"], errors="coerce").fillna(0)

        serials = {
            clean_text(v)
            for v in selected_dispatch.get("일련번호", pd.Series(dtype=object))
            if clean_text(v)
        }
        if serials:
            matched = pick[pick["일련번호"].map(clean_text).isin(serials)].copy()
        else:
            stores = {clean_text(v) for v in selected_dispatch["점포명"] if clean_text(v)}
            addresses = {clean_text(v) for v in selected_dispatch["주소"] if clean_text(v)}
            matched = pick[
                pick["출고처명(거래처)"].map(clean_text).isin(stores)
                & pick["배송처주소"].map(clean_text).isin(addresses)
            ].copy()

        if not matched.empty:
            matched = matched.rename(
                columns={
                    "품목코드": "상품코드",
                    "피킹수량": "수량",
                    "PTTN": "패턴",
                    "SIZE": "사이즈",
                }
            )
            result = (
                matched.groupby(
                    ["로케이션", "상품코드", "DOT", "패턴", "사이즈"],
                    dropna=False,
                    as_index=False,
                )["수량"]
                .sum()
            )
            result["_loc_sort"] = result["로케이션"].map(natural_location_key)
            result = result.sort_values(
                ["_loc_sort", "상품코드"],
                ascending=[False, True],
                na_position="last",
            ).drop(columns=["_loc_sort"])
            return result[["로케이션", "상품코드", "수량", "DOT", "패턴", "사이즈"]].reset_index(drop=True)

    # 피킹조회 원본이 아직 없거나 매칭되지 않은 경우 배차 데이터로 최소 정보 표시
    fallback = selected_dispatch[["로케이션", "상품코드", "수량", "DOT", "패턴", "사이즈"]].copy()
    fallback = (
        fallback.groupby(
            ["로케이션", "상품코드", "DOT", "패턴", "사이즈"],
            dropna=False,
            as_index=False,
        )["수량"]
        .sum()
    )
    fallback["_loc_sort"] = fallback["로케이션"].map(natural_location_key)
    return (
        fallback.sort_values(["_loc_sort", "상품코드"], ascending=[False, True])
        .drop(columns=["_loc_sort"])[["로케이션", "상품코드", "수량", "DOT", "패턴", "사이즈"]]
        .reset_index(drop=True)
    )


# =========================================================
# 인증 전
# =========================================================
if not st.session_state.authenticated:
    render_login()
    st.stop()


# 매 rerun마다 디스크의 마지막 원본을 읽음
shipping_df = load_source("출고")
picking_df = load_source("피킹조회")
dispatch_df = load_source("배차처리")
df = build_unified_data(shipping_df, picking_df, dispatch_df)


# =========================================================
# 1단계: 홈
# =========================================================
if st.session_state.step == "home":
    render_top_nav(show_back=False)
    st.markdown("### 📦 출고/배차 조회 시스템")
    role_label = "관리자모드" if st.session_state.role == "admin" else "일반모드"
    st.caption(f"{role_label} · 사내 업무용")
    render_data_status()

    render_admin_panel()

    if dispatch_df.empty:
        if st.session_state.role == "admin":
            st.warning("배차처리 파일이 필요합니다. 관리자 영역에서 원본 파일을 업로드해주세요.")
        else:
            st.warning("현재 조회 가능한 배차 데이터가 없습니다. 관리자 업데이트 후 이용해주세요.")
        st.stop()

    if shipping_df.empty or picking_df.empty:
        missing = []
        if shipping_df.empty:
            missing.append("출고")
        if picking_df.empty:
            missing.append("피킹조회")
        st.warning(
            "일부 데이터가 없습니다: " + ", ".join(missing) +
            ". 기본 조회는 가능하지만 출력여부/로케이션 정보가 일부 비어 있을 수 있습니다."
        )

    st.markdown("")

    # 오늘 물량 요약
    if not df.empty:
        total_qty = df["수량"].sum()
        unassigned_qty = df.loc[df["상태"] == "미배차", "수량"].sum()
        c1, c2 = st.columns(2)
        c1.metric("전체 물량", f"{format_qty(total_qty)}개")
        c2.metric("미배차", f"{format_qty(unassigned_qty)}개")
        st.markdown("")

    if st.button("🚗 승용", use_container_width=True, type="primary"):
        st.session_state.main_category = "승용"
        st.session_state.step = "sub"
        st.rerun()

    if st.button("🚚 화물", use_container_width=True, type="primary"):
        st.session_state.main_category = "화물"
        st.session_state.step = "sub"
        st.rerun()

    if st.button("🏷️ 마케팅출고", use_container_width=True, type="secondary"):
        st.session_state.main_category = "마케팅"
        st.session_state.step = "sub"
        st.rerun()


# =========================================================
# 2단계: 중간 메뉴
# =========================================================
elif st.session_state.step == "sub":
    render_top_nav(show_back=True)
    render_data_status()
    category = st.session_state.main_category
    st.markdown(f"### 📂 [{category}] 메뉴 선택")
    st.markdown("")

    category_df = df[df["분류"] == category]

    if category in ["승용", "화물"]:
        sub_menus = ["당착", "익착", "배차중", "택배", "미배차"]
        # 원본에 예상도착일이 비어 있는 배차가 있으면 누락시키지 않음
        if (category_df["상태"] == "미정").any():
            sub_menus.append("미정")

        for menu in sub_menus:
            menu_df = category_df[category_df["상태"] == menu]
            qty = menu_df["수량"].sum() if not menu_df.empty else 0
            label = f"{menu}  ·  {format_qty(qty)}개"
            if st.button(label, use_container_width=True):
                st.session_state.sub_category = menu
                st.session_state.step = "list"
                st.rerun()
    else:
        vehicle_df = category_df[
            (category_df["상태"] != "택배") & (category_df["상태"] != "미배차")
        ]
        parcel_df = category_df[category_df["상태"] == "택배"]

        menu_specs = [
            ("차량출고", vehicle_df["수량"].sum()),
            ("택배출고", parcel_df["수량"].sum()),
            ("상품별 출고 총수량", category_df["수량"].sum()),
        ]
        if (category_df["상태"] == "미배차").any():
            menu_specs.insert(2, ("미배차", category_df.loc[category_df["상태"] == "미배차", "수량"].sum()))

        for menu, qty in menu_specs:
            if st.button(f"{menu}  ·  {format_qty(qty)}개", use_container_width=True):
                st.session_state.sub_category = menu
                st.session_state.step = "list"
                st.rerun()


# =========================================================
# 3단계: 리스트
# =========================================================
elif st.session_state.step == "list":
    render_top_nav(show_back=True)
    render_data_status()
    category = st.session_state.main_category
    menu = st.session_state.sub_category
    st.markdown(f"### 📋 {category} > {menu}")

    filtered_df = df[df["분류"] == category].copy()

    if category in ["승용", "화물"]:
        filtered_df = filtered_df[filtered_df["상태"] == menu]
    else:
        if menu == "차량출고":
            filtered_df = filtered_df[
                (filtered_df["상태"] != "택배") & (filtered_df["상태"] != "미배차")
            ]
        elif menu == "택배출고":
            filtered_df = filtered_df[filtered_df["상태"] == "택배"]
        elif menu == "미배차":
            filtered_df = filtered_df[filtered_df["상태"] == "미배차"]

    # 마케팅 상품별 집계
    if category == "마케팅" and menu == "상품별 출고 총수량":
        search_query = st.text_input(
            "🔍 검색 (상품코드, 패턴, 사이즈, 로케이션)",
            placeholder="검색어를 입력하세요",
        )

        summary_df = (
            filtered_df.groupby(
                ["상품코드", "패턴", "사이즈", "로케이션"],
                dropna=False,
                as_index=False,
            )["수량"]
            .sum()
            .sort_values(["상품코드", "패턴", "사이즈"])
        )

        if search_query:
            q = search_query.strip()
            summary_df = summary_df[
                summary_df.astype(str).apply(
                    lambda x: x.str.contains(q, case=False, regex=False, na=False).any(),
                    axis=1,
                )
            ]

        st.metric("총 출고수량", f"{format_qty(summary_df['수량'].sum())}개")
        st.markdown(f"**상품 집계: {len(summary_df):,}건**")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
        st.stop()

    # 출력/미출력 필터
    filter_print = st.selectbox("출력 여부", ["전체", "출력", "미출력"])
    if filter_print != "전체":
        filtered_df = filtered_df[filtered_df["출력여부"] == filter_print]

    search_query = st.text_input(
        "🔍 검색 (차량번호, 점포명, 주소, 상품코드)",
        placeholder="검색어를 입력하세요",
    )
    if search_query:
        q = search_query.strip()
        searchable = ["차량번호", "점포명", "주소", "상품코드", "패턴", "사이즈"]
        mask = pd.Series(False, index=filtered_df.index)
        for col in searchable:
            mask = mask | filtered_df[col].astype(str).str.contains(
                q, case=False, regex=False, na=False
            )
        filtered_df = filtered_df[mask]

    total_qty = filtered_df["수량"].sum() if not filtered_df.empty else 0
    st.metric("조회 물량", f"{format_qty(total_qty)}개")

    grouped = group_for_list(filtered_df)
    st.markdown(f"**배송처 기준 {len(grouped):,}건**")
    st.caption("행을 터치/클릭하면 품목별 피킹정보가 열립니다.")

    if grouped.empty:
        st.info("조건에 맞는 조회 결과가 없습니다.")
        st.stop()

    # 현장 양식과 동일한 순서로 배송처 목록 표시
    table_df = grouped[[
        "피킹내림",
        "차량정보",
        "기사명",
        "연락처",
        "점포명",
        "주소",
        "수량",
        "가지수",
        "바코드",
        "윈터",
        "비고",
    ]].copy()
    table_df = table_df.rename(
        columns={
            "점포명": "거래처명",
            "주소": "배송처주소",
        }
    )

    event = st.dataframe(
        table_df,
        use_container_width=True,
        hide_index=True,
        height=min(700, 42 + 36 * max(1, len(table_df))),
        on_select="rerun",
        selection_mode="single-row",
        key=f"dispatch_table_{category}_{menu}_{st.session_state.table_nonce}",
        column_config={
            "피킹내림": st.column_config.TextColumn("피킹내림", width="small"),
            "차량정보": st.column_config.TextColumn("차량정보", width="medium"),
            "기사명": st.column_config.TextColumn("기사명", width="small"),
            "연락처": st.column_config.TextColumn("연락처", width="medium"),
            "거래처명": st.column_config.TextColumn("거래처명", width="large"),
            "배송처주소": st.column_config.TextColumn("배송처주소", width="large"),
            "수량": st.column_config.NumberColumn("수량", format="%d"),
            "가지수": st.column_config.NumberColumn("가지수", format="%d"),
            "바코드": st.column_config.TextColumn("바코드", width="small"),
            "윈터": st.column_config.TextColumn("윈터", width="small"),
            "비고": st.column_config.TextColumn("비고", width="medium"),
        },
    )

    selected_rows = event.selection.rows if event and hasattr(event, "selection") else []
    if selected_rows:
        selected_idx = int(selected_rows[0])
        if 0 <= selected_idx < len(grouped):
            st.session_state.selected_group = grouped.iloc[selected_idx].to_dict()
            st.session_state.step = "detail"
            st.rerun()


# =========================================================
# 4단계: 상세
# =========================================================
elif st.session_state.step == "detail":
    render_top_nav(show_back=True)
    render_data_status()
    st.markdown("### 🔍 상세 세부사항")
    st.markdown("")

    group = st.session_state.selected_group
    if not group:
        st.warning("선택된 항목이 없습니다.")
        if st.button("목록으로 돌아가기", use_container_width=True):
            st.session_state.step = "list"
            st.rerun()
        st.stop()

    detail_df = rows_for_selected_group(df, group)
    if detail_df.empty:
        st.warning("현재 데이터에서 해당 항목을 찾지 못했습니다. 최신 데이터로 다시 조회해주세요.")
        st.stop()

    vehicle = clean_text(group.get("차량정보")) or "미배차"
    total_qty = detail_df["수량"].sum()

    st.info(f"**{clean_text(group.get('점포명'))}**")
    c1, c2 = st.columns(2)
    c1.metric("총 물량", f"{format_qty(total_qty)}개")
    c2.metric("품목", f"{detail_df['상품코드'].nunique():,}종")

    st.markdown(f"• **차량정보:** {vehicle}")
    if clean_text(group.get("기사명")):
        st.markdown(f"• **기사명:** {clean_text(group.get('기사명'))}")
    if clean_text(group.get("연락처")):
        st.markdown(f"• **연락처:** {clean_text(group.get('연락처'))}")
    st.markdown(f"• **구분:** {clean_text(group.get('상태'))}")
    st.markdown(f"• **출력여부:** {clean_text(group.get('출력여부'))}")
    if clean_text(group.get("출고일시")):
        st.markdown(f"• **출고일시:** {clean_text(group.get('출고일시'))}")
    if clean_text(group.get("예상도착일시")):
        st.markdown(f"• **예상도착:** {clean_text(group.get('예상도착일시'))}")
    st.markdown(f"• **배송 주소:** {clean_text(group.get('주소'))}")

    st.markdown("---")
    st.markdown("#### 📦 품목별 피킹정보")
    st.caption("로케이션 내림차순 → 품목코드 순")

    picking_detail = build_picking_detail(picking_df, detail_df)
    if picking_detail.empty:
        st.info("표시할 피킹 상세정보가 없습니다.")
    else:
        st.dataframe(
            picking_detail,
            use_container_width=True,
            hide_index=True,
            column_config={
                "로케이션": st.column_config.TextColumn("로케이션", width="medium"),
                "상품코드": st.column_config.TextColumn("품목코드", width="medium"),
                "수량": st.column_config.NumberColumn("수량", format="%d"),
                "DOT": st.column_config.TextColumn("디오티", width="medium"),
                "패턴": st.column_config.TextColumn("패턴", width="medium"),
                "사이즈": st.column_config.TextColumn("사이즈", width="medium"),
            },
        )
