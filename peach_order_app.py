# =============================================================================
# 🍑 실크로드 복숭아 농장 주문관리 시스템 (Streamlit + Google Sheets)
# =============================================================================
#
# ┌─────────────────────────────────────────────────────────────────────────┐
# │  【 Google Sheets 최초 세팅 가이드 】                                      │
# │                                                                         │
# │  1단계: Google Cloud 프로젝트 & 서비스 계정 생성                            │
# │  ─────────────────────────────────────────────────────────────────────  │
# │  1. https://console.cloud.google.com 접속 → 새 프로젝트 생성              │
# │  2. 왼쪽 메뉴 "API 및 서비스" → "라이브러리" 클릭                          │
# │  3. "Google Sheets API" 검색 → 활성화                                    │
# │  4. "Google Drive API" 검색 → 활성화                                     │
# │  5. 왼쪽 메뉴 "API 및 서비스" → "사용자 인증 정보" 클릭                    │
# │  6. "사용자 인증 정보 만들기" → "서비스 계정" 선택                          │
# │  7. 서비스 계정 이름 입력 (예: peach-farm) → 완료                          │
# │  8. 생성된 서비스 계정 클릭 → "키" 탭 → "키 추가" → "새 키 만들기"         │
# │  9. JSON 형식 선택 → 다운로드 (이 파일이 서비스 계정 키입니다)              │
# │                                                                         │
# │  2단계: Google Sheets 스프레드시트 생성                                    │
# │  ─────────────────────────────────────────────────────────────────────  │
# │  1. https://sheets.google.com 에서 새 스프레드시트 생성                   │
# │  2. 시트 4개 생성: 주문목록 / 상품목록 / 고객목록 / 설정                   │
# │  3. 주문목록 1행 헤더:                                                    │
# │     주문번호|주문일시|주문자이름|주문자전화번호|주문자이메일|               │
# │     받는분이름|받는분전화번호|받는분주소|상품명|수량|배송메모|상태          │
# │  4. 상품목록 1행: 상품명 | 단가 | 설명                                    │
# │     2행부터 상품 입력 (예: 복숭아 4kg 일반용 | 30000 | 신선한 복숭아)      │
# │  5. 고객목록 1행: 이름 | 이메일 | 전화번호                                 │
# │  6. 설정 시트 A/B열:                                                     │
# │     order_start    | 2026-07-01 09:00                                   │
# │     order_end      | 2026-07-10 18:00                                   │
# │     bank           | 농협                                                │
# │     account_number | 000-0000-0000                                      │
# │     holder         | 장명숙                                              │
# │                                                                         │
# │  3단계: 서비스 계정 이메일로 스프레드시트 공유                              │
# │  ─────────────────────────────────────────────────────────────────────  │
# │  스프레드시트 → 공유 → 서비스 계정 client_email 입력 → 편집자 권한 부여     │
# │                                                                         │
# │  4단계: .streamlit/secrets.toml 작성                                     │
# │  ─────────────────────────────────────────────────────────────────────  │
# │  [gcp_service_account]                                                  │
# │  type = "service_account"                                               │
# │  project_id = "your-project-id"                                         │
# │  private_key_id = "..."                                                 │
# │  private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END..."      │
# │  client_email = "peach@project.iam.gserviceaccount.com"                │
# │  client_id = "..."                                                      │
# │  auth_uri = "https://accounts.google.com/o/oauth2/auth"                │
# │  token_uri = "https://oauth2.googleapis.com/token"                     │
# │                                                                         │
# │  [app]                                                                  │
# │  admin_password = "복숭아1234"                                           │
# │  spreadsheet_id = "1ABC...xyz"                                          │
# │  farm_name = "실크로드 복숭아 농장"                                          │
# │  order_url = "https://your-app.streamlit.app"                           │
# │                                                                         │
# │  [email]                                                                │
# │  sender = "your-gmail@gmail.com"                                        │
# │  app_password = "xxxx xxxx xxxx xxxx"                                   │
# │                                                                         │
# │  [account]                                                              │
# │  bank = "농협"                                                           │
# │  account_number = "000-0000-0000"                                       │
# │  holder = "장명숙"                                                       │
# │                                                                         │
# │  실행: streamlit run peach_order_system.py                              │
# └─────────────────────────────────────────────────────────────────────────┘
# =============================================================================

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import io
import random
import string
import re
import time

# =============================================================================
# 페이지 기본 설정
# =============================================================================
st.set_page_config(
    page_title="복숭아농장 주문관리",
    page_icon="🍑",
    layout="wide",
    initial_sidebar_state="auto",
)

# =============================================================================
# 전체 테마 CSS — 따뜻한 복숭아 색상
# =============================================================================
st.markdown("""
<style>
/* ── 전체 배경 ── */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #fff8f0 0%, #fff3e8 50%, #ffeedd 100%);
}

/* ── 사이드바 ── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #ff8c42 0%, #e55a00 100%);
}
[data-testid="stSidebar"] * { color: white !important; }
[data-testid="stSidebar"] .stTextInput input {
    background: rgba(255,255,255,0.2) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.4) !important;
    border-radius: 8px !important;
}
[data-testid="stSidebar"] .stTextInput input::placeholder {
    color: rgba(255,255,255,0.7) !important;
}

/* ── 헤더 ── */
.peach-header {
    background: linear-gradient(135deg, #ff8c42, #ff6b1a, #e55a00);
    color: white;
    padding: 2rem 1.5rem;
    border-radius: 16px;
    text-align: center;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 20px rgba(255,107,26,0.3);
}
.peach-header h1 { font-size: 2.2rem; margin: 0; text-shadow: 1px 1px 3px rgba(0,0,0,0.2); }
.peach-header p  { margin: 0.5rem 0 0 0; opacity: 0.9; font-size: 1.05rem; }

/* ── 카드 ── */
.peach-card {
    background: white;
    border-radius: 12px;
    padding: 1.4rem;
    box-shadow: 0 2px 12px rgba(255,107,26,0.12);
    border-left: 4px solid #ff8c42;
    margin-bottom: 1rem;
}

/* ── 공지 박스 ── */
.notice-box {
    background: #fff3e0;
    border: 2px solid #ffb74d;
    border-radius: 12px;
    padding: 1.4rem 1.5rem;
    margin: 1rem 0;
    text-align: center;
}
.notice-box.closed { background: #fce4ec; border-color: #e57373; }
.notice-box.open   { background: #e8f5e9; border-color: #66bb6a; }

/* ── 계좌 안내 ── */
.account-box {
    background: linear-gradient(135deg, #fff8f0, #fff3e8);
    border: 2px solid #ff8c42;
    border-radius: 12px;
    padding: 1.5rem;
    text-align: center;
    margin: 1rem 0;
}
.account-box .bank-name  { font-size: 1.3rem; font-weight: bold; color: #e55a00; }
.account-box .account-num { font-size: 1.8rem; font-weight: bold; color: #333; letter-spacing: 2px; margin: 0.3rem 0; }

/* ── 수령자 박스 ── */
.recipient-box {
    background: #fff8f0;
    border: 1px solid #ffcc99;
    border-radius: 10px;
    padding: 1rem;
    margin-bottom: 0.8rem;
}
.recipient-box-title { font-weight: bold; color: #e55a00; font-size: 1.05rem; margin-bottom: 0.5rem; }

/* ── 카운트다운 ── */
.countdown-box {
    background: linear-gradient(135deg, #e8f5e9, #c8e6c9);
    border: 2px solid #66bb6a;
    border-radius: 12px;
    padding: 1rem;
    text-align: center;
    font-size: 1.1rem;
    color: #2e7d32;
    font-weight: bold;
    margin: 0.5rem 0;
}

/* ── 버튼 ── */
.stButton > button {
    background: linear-gradient(135deg, #ff8c42, #ff6b1a) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.5rem 1.5rem !important;
    font-weight: bold !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #ff6b1a, #e55a00) !important;
    box-shadow: 0 3px 10px rgba(255,107,26,0.4) !important;
    transform: translateY(-1px) !important;
}

/* ── 탭 ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: white;
    border-radius: 10px;
    padding: 4px;
}
.stTabs [data-baseweb="tab"]    { border-radius: 8px; padding: 0.4rem 1rem; }
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #ff8c42, #ff6b1a) !important;
    color: white !important;
}

/* ── 지표 카드 ── */
.metric-card {
    background: white;
    border-radius: 10px;
    padding: 1rem;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border-top: 3px solid #ff8c42;
}
.metric-card .metric-value { font-size: 2rem; font-weight: bold; }
.metric-card .metric-label { color: #666; font-size: 0.9rem; }

/* ── 모바일 반응형 ── */
@media (max-width: 768px) {
    .peach-header h1 { font-size: 1.5rem; }
    .peach-header p  { font-size: 0.9rem; }
}
</style>
""", unsafe_allow_html=True)


# =============================================================================
# Google Sheets 연결 헬퍼
# =============================================================================

@st.cache_resource(ttl=300)
def get_gspread_client():
    """
    Google Sheets API 클라이언트를 생성합니다.
    secrets.toml의 [gcp_service_account] 항목을 읽습니다.
    연결 실패 시 None을 반환합니다 (앱이 중단되지 않음).
    """
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=scopes,
        )
        return gspread.authorize(creds)
    except Exception:
        return None


def get_spreadsheet():
    """스프레드시트 객체를 반환합니다. 실패 시 None."""
    client = get_gspread_client()
    if client is None:
        return None
    try:
        return client.open_by_key(st.secrets["app"]["spreadsheet_id"])
    except Exception:
        return None


def get_sheet(name: str):
    """시트 이름으로 워크시트 객체를 반환합니다. 실패 시 None."""
    ss = get_spreadsheet()
    if ss is None:
        return None
    try:
        return ss.worksheet(name)
    except Exception:
        return None


# =============================================================================
# 설정 시트 읽기 / 쓰기
# =============================================================================

def load_settings() -> dict:
    """
    '설정' 시트에서 key|value 형태로 설정값을 읽습니다.
    시트 연결 실패 시 secrets.toml의 [account] 값을 기본값으로 사용합니다.
    """
    # secrets에서 기본값 읽기 (secrets.toml 없어도 기본값으로 동작)
    try:
        def_bank  = st.secrets["account"]["bank"]
        def_acct  = st.secrets["account"]["account_number"]
        def_holder = st.secrets["account"]["holder"]
    except Exception:
        def_bank, def_acct, def_holder = "농협", "000-0000-0000", "장명숙"

    defaults = {
        "order_start":    "2026-07-01 09:00",
        "order_end":      "2026-07-10 18:00",
        "bank":           def_bank,
        "account_number": def_acct,
        "holder":         def_holder,
        "farm_phone":     "",
    }

    sheet = get_sheet("설정")
    if sheet is None:
        return defaults

    try:
        rows = sheet.get_all_values()
        settings = defaults.copy()
        for row in rows:
            if len(row) >= 2 and row[0].strip():
                settings[row[0].strip()] = row[1].strip()
        return settings
    except Exception:
        return defaults


def save_settings(settings: dict) -> bool:
    """설정 딕셔너리를 '설정' 시트에 key|value 형태로 저장합니다."""
    sheet = get_sheet("설정")
    if sheet is None:
        return False
    try:
        sheet.clear()
        rows = [[k, v] for k, v in settings.items()]
        if rows:
            sheet.update("A1", rows)
        return True
    except Exception as e:
        st.error(f"설정 저장 실패: {e}")
        return False


# =============================================================================
# 상품 목록
# =============================================================================

def load_products() -> list:
    """'상품목록' 시트 A열(2행~)에서 상품명을 읽어 반환합니다."""
    sheet = get_sheet("상품목록")
    if sheet is None:
        # 시트 연결 전 기본 상품 (데모용)
        return [
            "복숭아 4kg 일반용",
            "복숭아 4kg 선물용",
            "복숭아 10kg 일반용",
            "복숭아 10kg 선물용",
        ]
    try:
        rows = sheet.get_all_values()
        # 1행은 헤더(상품명|단가|설명), 2행부터 데이터
        products = [row[0].strip() for row in rows[1:] if row and row[0].strip()]
        return products if products else ["복숭아 4kg 일반용", "복숭아 4kg 선물용"]
    except Exception:
        return ["복숭아 4kg 일반용", "복숭아 4kg 선물용"]


# =============================================================================
# 주문 저장 / 로드
# =============================================================================

def generate_order_number() -> str:
    """주문번호 생성: PEACH-YYYYMMDD-XXXX (4자리 랜덤 숫자)"""
    today  = datetime.now().strftime("%Y%m%d")
    suffix = "".join(random.choices(string.digits, k=4))
    return f"PEACH-{today}-{suffix}"


def save_orders(order_rows: list) -> bool:
    """
    '주문목록' 시트에 주문 행을 추가합니다.
    시트가 비어있으면 헤더를 먼저 작성합니다.
    """
    sheet = get_sheet("주문목록")
    if sheet is None:
        return False
    try:
        existing = sheet.get_all_values()
        if not existing:
            # 헤더 자동 생성
            header = [
                "주문번호", "주문일시", "주문자이름", "주문자전화번호",
                "보내는분이름", "보내는분전화번호", "보내는분주소",
                "받는분이름", "받는분전화번호", "받는분주소",
                "상품명", "수량", "배송메모", "상태",
            ]
            sheet.append_row(header)
        for row in order_rows:
            sheet.append_row(row)
        return True
    except Exception as e:
        st.error(f"주문 저장 중 오류가 발생했습니다: {e}")
        return False


def load_orders() -> pd.DataFrame:
    """'주문목록' 시트 전체를 DataFrame으로 반환합니다."""
    sheet = get_sheet("주문목록")
    if sheet is None:
        return pd.DataFrame()
    try:
        rows = sheet.get_all_values()
        if len(rows) < 2:
            return pd.DataFrame()
        return pd.DataFrame(rows[1:], columns=rows[0])
    except Exception:
        return pd.DataFrame()


# =============================================================================
# 고객 목록
# =============================================================================

def load_customers() -> pd.DataFrame:
    """'고객목록' 시트에서 고객 정보를 DataFrame으로 반환합니다."""
    sheet = get_sheet("고객목록")
    if sheet is None:
        return pd.DataFrame(columns=["이름", "이메일", "전화번호"])
    try:
        rows = sheet.get_all_values()
        if len(rows) < 2:
            return pd.DataFrame(columns=["이름", "이메일", "전화번호"])
        return pd.DataFrame(rows[1:], columns=rows[0])
    except Exception:
        return pd.DataFrame(columns=["이름", "이메일", "전화번호"])


# =============================================================================
# 이메일 발송
# =============================================================================

def send_email(to_addr: str, subject: str, body: str) -> bool:
    """
    Gmail SMTP(SSL 465포트)를 통해 이메일을 발송합니다.
    secrets.toml의 [email] sender / app_password 를 사용합니다.
    Gmail 앱 비밀번호는 Google 계정 → 보안 → 2단계 인증 → 앱 비밀번호에서 생성합니다.
    """
    try:
        sender   = st.secrets["email"]["sender"]
        app_pw   = st.secrets["email"]["app_password"]
    except Exception:
        # secrets 미설정 시 발송 스킵 (주문은 이미 저장됨)
        return False

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = sender
        msg["To"]      = to_addr
        html_body = "<html><body>" + body.replace("\n", "<br>") + "</body></html>"
        msg.attach(MIMEText(body,      "plain", "utf-8"))
        msg.attach(MIMEText(html_body, "html",  "utf-8"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, app_pw)
            server.sendmail(sender, to_addr, msg.as_string())
        return True
    except Exception as e:
        st.warning(f"이메일 발송 실패 ({to_addr}): {e}")
        return False


def send_confirmation_email(
    orderer_name: str,
    orderer_email: str,
    order_number: str,
    recipients: list,
    settings: dict,
    farm_name: str,
):
    """
    주문 완료 후 주문자에게 확인 이메일을 발송합니다.
    recipients: [{"name":..., "phone":..., "address":..., "product":..., "qty":..., "memo":...}]
    """
    rec_lines = ""
    for i, r in enumerate(recipients, 1):
        rec_lines += (
            f"\n  [{i}번째 수령자]\n"
            f"  이름: {r['name']}\n"
            f"  전화: {r['phone']}\n"
            f"  주소: {r['address']}\n"
            f"  상품: {r['product']} × {r['qty']}박스\n"
            f"  메모: {r.get('memo') or '없음'}\n"
        )

    body = (
        f"안녕하세요, {orderer_name}님! 🍑\n\n"
        f"{farm_name} 주문이 접수되었습니다.\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"주문번호: {order_number}\n"
        f"주문일시: {datetime.now().strftime('%Y년 %m월 %d일 %H:%M')}\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        f"【 수령자 정보 】\n{rec_lines}\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"【 입금 안내 】\n"
        f"은행: {settings.get('bank', '농협')}\n"
        f"계좌: {settings.get('account_number', '000-0000-0000')}\n"
        f"예금주: {settings.get('holder', '장명숙')}\n\n"
        f"※ 입금자명을 주문자 성함으로 해주세요.\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        f"맛있는 복숭아로 보답하겠습니다! 🍑\n{farm_name} 드림"
    )
    send_email(orderer_email, f"[{farm_name}] 주문 접수 확인 - {order_number}", body)


# =============================================================================
# 주문 기간 판단
# =============================================================================

def check_order_period(settings: dict):
    """
    현재 시각과 주문 기간을 비교합니다.
    반환: ("before" | "open" | "closed", start_dt, end_dt)
    """
    now = datetime.now()
    try:
        start_dt = datetime.strptime(
            settings.get("order_start", "2099-01-01 00:00"), "%Y-%m-%d %H:%M"
        )
        end_dt = datetime.strptime(
            settings.get("order_end", "2099-01-01 00:00"), "%Y-%m-%d %H:%M"
        )
    except ValueError:
        return ("closed", None, None)

    if now < start_dt:
        return ("before", start_dt, end_dt)
    elif now > end_dt:
        return ("closed", start_dt, end_dt)
    else:
        return ("open", start_dt, end_dt)


def format_countdown(end_dt: datetime) -> str:
    """마감까지 남은 시간을 '00일 00시간 00분' 형식으로 반환합니다."""
    remaining = end_dt - datetime.now()
    if remaining.total_seconds() <= 0:
        return "마감되었습니다"
    days = remaining.days
    hours, rem = divmod(remaining.seconds, 3600)
    minutes    = rem // 60
    parts = []
    if days > 0:
        parts.append(f"{days}일")
    if hours > 0:
        parts.append(f"{hours}시간")
    parts.append(f"{minutes}분")
    return " ".join(parts)


# =============================================================================
# 입력값 검증
# =============================================================================

def validate_phone(phone: str) -> bool:
    """전화번호 형식 검증: 010-XXXX-XXXX"""
    return bool(re.match(r"^010-\d{3,4}-\d{4}$", phone.strip()))


def validate_email_addr(email: str) -> bool:
    """이메일 형식 검증"""
    return bool(re.match(r"^[\w.+-]+@[\w-]+\.[a-zA-Z]{2,}$", email.strip()))


# =============================================================================
# 로젠택배 엑셀 생성
# =============================================================================

def generate_logen_excel(df: pd.DataFrame, farm_name: str, settings: dict) -> bytes:
    """
    주문 DataFrame을 로젠택배 업로드 양식 엑셀로 변환합니다.
    컬럼 순서: 받는분성명/받는분주소/받는분전화번호/품목/박스수량/
              보내는분성명/보내는분주소/보내는분전화번호/배송메세지
    """
    farm_phone = settings.get("farm_phone", "")

    logen_df = pd.DataFrame({
        "받는분성명":       df["받는분이름"].values      if "받는분이름"      in df.columns else "",
        "받는분주소":       df["받는분주소"].values      if "받는분주소"      in df.columns else "",
        "받는분전화번호":   df["받는분전화번호"].values  if "받는분전화번호"  in df.columns else "",
        "품목":             df["상품명"].values          if "상품명"          in df.columns else "",
        "박스수량":         df["수량"].values            if "수량"            in df.columns else 1,
        "보내는분성명":     df["보내는분이름"].values    if "보내는분이름"    in df.columns else farm_name,
        "보내는분주소":     df["보내는분주소"].values    if "보내는분주소"    in df.columns else "",
        "보내는분전화번호": df["보내는분전화번호"].values if "보내는분전화번호" in df.columns else farm_phone,
        "배송메세지":       df["배송메모"].values        if "배송메모"        in df.columns else "",
    })

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        logen_df.to_excel(writer, index=False, sheet_name="로젠택배")
        ws = writer.sheets["로젠택배"]
        # 컬럼 너비 자동 조정
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max(max_len + 2, 14)
    output.seek(0)
    return output.read()


# =============================================================================
# 사이드바 — 관리자 로그인
# =============================================================================

def render_sidebar() -> bool:
    """
    사이드바에 관리자 비밀번호 입력 UI를 표시합니다.
    올바른 비밀번호 입력 시 True를 반환합니다.
    """
    with st.sidebar:
        st.markdown("## 🍑 복숭아농장")
        st.markdown("---")
        st.markdown("### 🔐 관리자 로그인")
        pw = st.text_input("비밀번호", type="password", placeholder="관리자 비밀번호 입력", key="admin_pw")

        if not pw:
            st.markdown(
                "<div style='font-size:0.85rem;opacity:0.85;'>"
                "관리자만 접근 가능합니다.<br>"
                "고객 주문은 아래 페이지에서 진행해주세요. 🍑"
                "</div>",
                unsafe_allow_html=True,
            )
            return False

        try:
            correct_pw = st.secrets["app"]["admin_password"]
        except Exception:
            # secrets.toml 미설정 시 관리자 접근 차단 (보안)
            st.error("⚠️ 관리자 비밀번호가 설정되지 않았습니다. secrets.toml을 확인해주세요.")
            return False

        if pw == correct_pw:
            st.success("✅ 관리자 모드")
            if st.button("🏠 고객 화면으로", use_container_width=True):
                st.session_state["force_customer"] = True
                st.rerun()
            return True
        else:
            st.error("❌ 비밀번호가 틀렸습니다")
            return False


# =============================================================================
# [A] 고객 주문 페이지
# =============================================================================

def _get_farm_name() -> str:
    """secrets에서 농장 이름을 읽습니다. 없으면 기본값 반환."""
    try:
        return st.secrets["app"]["farm_name"]
    except Exception:
        return "실크로드 복숭아 농장"


def _empty_recipient() -> dict:
    """빈 수령자 딕셔너리를 반환합니다."""
    return {"name": "", "phone": "", "address": "", "product": "", "qty": 1, "memo": ""}


def render_customer_page(settings: dict, products: list):
    """고객이 복숭아를 주문하는 공개 페이지를 렌더링합니다."""
    farm_name = _get_farm_name()

    # ── 헤더 ──
    st.markdown(
        f"<div class='peach-header'>"
        f"<h1>🍑 {farm_name}</h1>"
        f"<p>신선한 복숭아를 직접 산지에서 배달해드립니다</p>"
        f"</div>",
        unsafe_allow_html=True,
    )

    # ── 주문 기간 확인 ──
    status, start_dt, end_dt = check_order_period(settings)

    if status == "before":
        st.markdown(
            f"<div class='notice-box'>"
            f"<h2>🕐 주문이 곧 시작됩니다</h2>"
            f"<p style='font-size:1.2rem;'>"
            f"주문 시작: <strong>{start_dt.strftime('%Y년 %m월 %d일 %H:%M')}</strong>"
            f"</p>"
            f"<p>잠시만 기다려 주세요! 맛있는 복숭아를 준비 중입니다. 🍑</p>"
            f"</div>",
            unsafe_allow_html=True,
        )
        return

    if status == "closed":
        end_str = end_dt.strftime("%Y년 %m월 %d일 %H:%M") if end_dt else "알 수 없음"
        st.markdown(
            f"<div class='notice-box closed'>"
            f"<h2>🚫 주문이 마감되었습니다</h2>"
            f"<p style='font-size:1.1rem;'>"
            f"마감일시: <strong>{end_str}</strong>"
            f"</p>"
            f"<p>다음 주문 기간을 기다려 주세요. 감사합니다! 🍑</p>"
            f"</div>",
            unsafe_allow_html=True,
        )
        return

    # ── 주문 접수 중: 카운트다운 표시 ──
    countdown = format_countdown(end_dt)
    st.markdown(
        f"<div class='countdown-box'>"
        f"✅ 주문 접수 중 — 마감까지 <strong>{countdown}</strong> 남음"
        f"</div>",
        unsafe_allow_html=True,
    )

    # ── 주문 완료 상태 처리 ──
    if st.session_state.get("order_complete"):
        _render_order_complete(settings, farm_name)
        st.markdown("---")
        if st.button("🔄 새 주문하기"):
            st.session_state["order_complete"] = False
            st.session_state["order_result"]   = None
            st.session_state["recipients"]     = [_empty_recipient()]
            st.rerun()
        return

    # ── 세션 상태 초기화 ──
    if not st.session_state.get("recipients"):
        st.session_state["recipients"] = [_empty_recipient()]

    # ── 주문자(입금자) 정보 ──
    st.markdown("### 👤 주문자(입금자) 정보")
    st.caption("실제 주문하고 입금하시는 분의 정보입니다.")
    with st.container():
        orderer_name  = st.text_input("이름 *",     placeholder="홍길동",        key="orderer_name")
        orderer_phone = st.text_input("전화번호 *", placeholder="010-1234-5678", key="orderer_phone")

    # ── 보내는 분(발송인) 정보 ──
    st.markdown("### 📤 보내는 분(발송인) 정보")
    st.caption("택배 송장에 '보내는 분'으로 표시될 정보입니다.")

    sender_same = st.checkbox("보내는 분이 주문자와 같습니다", key="sender_same", value=True)

    if sender_same:
        sender_name  = orderer_name
        sender_phone = orderer_phone
        sender_address = st.text_input("보내는 분 주소 *", placeholder="경북 김천시 OO로 OO", key="sender_address")
    else:
        sender_name    = st.text_input("보내는 분 이름 *",     placeholder="홍길동",             key="sender_name")
        sender_phone   = st.text_input("보내는 분 전화번호 *", placeholder="010-1234-5678",      key="sender_phone")
        sender_address = st.text_input("보내는 분 주소 *",     placeholder="경북 김천시 OO로 OO", key="sender_address")

    # ── 받는 분 정보 ──
    st.markdown("### 📦 받는 분 정보")

    recipient_same = st.checkbox("받는 분이 보내는 분과 같습니다 (본인 수령)", key="recipient_same")

    recipients = st.session_state["recipients"]

    if recipient_same:
        st.info("✅ 보내는 분 정보로 배송됩니다. 상품과 수량만 선택해주세요.")
        rec = recipients[0]
        rec["name"]    = sender_name
        rec["phone"]   = sender_phone
        rec["address"] = sender_address
        with st.container():
            default_prod_idx = products.index(rec["product"]) if rec["product"] in products else 0
            rec["product"] = st.selectbox("상품 선택 *", products, index=default_prod_idx, key="rprod_self")
            rec["qty"]     = st.number_input("수량 (박스) *", min_value=1, max_value=99, value=int(rec["qty"]), step=1, key="rqty_self")
            rec["memo"]    = st.text_input("배송 메모 (선택)", value=rec.get("memo", ""), key="rmemo_self", placeholder="경비실 맡겨주세요")
    else:
        st.caption("여러 명에게 따로 보낼 수 있습니다. '받는 분 추가' 버튼을 이용해주세요.")
        for i in range(len(recipients)):
            rec = recipients[i]
            with st.expander(f"📮 {i + 1}번째 받는 분", expanded=True):
                col_main, col_del = st.columns([6, 1])
                with col_del:
                    if len(recipients) > 1:
                        if st.button("🗑️", key=f"del_{i}", help="이 수령자 삭제"):
                            st.session_state["recipients"].pop(i)
                            st.rerun()
                with col_main:
                    rec["name"]    = st.text_input("받는 분 이름 *",     value=rec["name"],    key=f"rname_{i}",  placeholder="홍길동")
                    rec["phone"]   = st.text_input("받는 분 전화번호 *",  value=rec["phone"],   key=f"rphone_{i}", placeholder="010-0000-0000")
                    rec["address"] = st.text_input("받는 분 주소 *",      value=rec["address"], key=f"raddr_{i}",  placeholder="서울시 강남구 테헤란로 123")
                    default_prod_idx = products.index(rec["product"]) if rec["product"] in products else 0
                    rec["product"] = st.selectbox("상품 선택 *", products, index=default_prod_idx, key=f"rprod_{i}")
                    rec["qty"]     = st.number_input("수량 (박스) *", min_value=1, max_value=99, value=int(rec["qty"]), step=1, key=f"rqty_{i}")
                    rec["memo"]    = st.text_input("배송 메모 (선택)", value=rec.get("memo", ""), key=f"rmemo_{i}", placeholder="경비실 맡겨주세요")

        if st.button("➕ 받는 분 추가"):
            st.session_state["recipients"].append(_empty_recipient())
            st.rerun()

    # ── 계좌 안내 ──
    st.markdown("---")
    bank   = settings.get("bank", "농협")
    acct   = settings.get("account_number", "000-0000-0000")
    holder = settings.get("holder", "장명숙")
    st.markdown(
        f"<div class='account-box'>"
        f"<p style='margin:0 0 0.3rem 0;color:#666;'>📌 주문 후 아래 계좌로 입금해 주세요</p>"
        f"<div class='bank-name'>{bank}은행</div>"
        f"<div class='account-num'>{acct}</div>"
        f"<div style='color:#555;'>예금주: <strong>{holder}</strong></div>"
        f"<p style='font-size:0.85rem;color:#888;margin-top:0.5rem;'>"
        f"입금자명을 주문자 성함으로 해주세요</p>"
        f"</div>",
        unsafe_allow_html=True,
    )

    # ── 주문 제출 버튼 ──
    if st.button("🍑 주문 완료하기", use_container_width=True, type="primary"):
        errors = _validate_order(
            orderer_name, orderer_phone,
            sender_name, sender_phone, sender_address,
            recipients
        )
        if errors:
            for err in errors:
                st.error(err)
        else:
            _submit_order(
                orderer_name, orderer_phone,
                sender_name, sender_phone, sender_address,
                recipients, settings, farm_name,
            )


def _validate_order(orderer_name, orderer_phone, sender_name, sender_phone, sender_address, recipients) -> list:
    """주문 폼 전체 유효성 검사. 오류 메시지 리스트 반환."""
    errors = []
    # 주문자 검증
    if not orderer_name.strip():
        errors.append("❗ 주문자 이름을 입력해주세요.")
    if not orderer_phone.strip():
        errors.append("❗ 주문자 전화번호를 입력해주세요.")
    elif not validate_phone(orderer_phone):
        errors.append("❗ 주문자 전화번호 형식이 올바르지 않습니다. (예: 010-1234-5678)")
    # 보내는 분 검증
    if not sender_name.strip():
        errors.append("❗ 보내는 분 이름을 입력해주세요.")
    if not sender_phone.strip():
        errors.append("❗ 보내는 분 전화번호를 입력해주세요.")
    elif not validate_phone(sender_phone):
        errors.append("❗ 보내는 분 전화번호 형식이 올바르지 않습니다.")
    if not sender_address.strip():
        errors.append("❗ 보내는 분 주소를 입력해주세요.")

    for i, rec in enumerate(recipients, 1):
        if not rec["name"].strip():
            errors.append(f"❗ {i}번째 받는 분의 이름을 입력해주세요.")
        if not rec["phone"].strip():
            errors.append(f"❗ {i}번째 받는 분의 전화번호를 입력해주세요.")
        elif not validate_phone(rec["phone"]):
            errors.append(f"❗ {i}번째 받는 분의 전화번호 형식이 올바르지 않습니다.")
        if not rec["address"].strip():
            errors.append(f"❗ {i}번째 받는 분의 주소를 입력해주세요.")
        if not rec.get("product"):
            errors.append(f"❗ {i}번째 받는 분의 상품을 선택해주세요.")
    return errors


def _submit_order(orderer_name, orderer_phone, sender_name, sender_phone, sender_address, recipients, settings, farm_name):
    """유효성 검사를 통과한 주문을 시트에 저장합니다."""
    order_number = generate_order_number()
    now_str      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 수령자별로 한 행씩 저장 (주문번호는 동일)
    rows = []
    for rec in recipients:
        rows.append([
            order_number,
            now_str,
            orderer_name.strip(),
            orderer_phone.strip(),
            sender_name.strip(),
            sender_phone.strip(),
            sender_address.strip(),
            rec["name"].strip(),
            rec["phone"].strip(),
            rec["address"].strip(),
            rec["product"],
            str(rec["qty"]),
            rec.get("memo", "").strip(),
            "대기",
        ])

    with st.spinner("주문을 저장하는 중입니다..."):
        success = save_orders(rows)

    if success:
        st.session_state["order_complete"] = True
        st.session_state["order_result"]   = {
            "order_number": order_number,
            "name":         orderer_name,
            "address":      orderer_address,
            "recipients":   recipients,
        }
        st.balloons()
        st.rerun()
    else:
        st.error(
            "⚠️ 주문 저장에 실패했습니다. "
            "잠시 후 다시 시도하거나 전화로 문의해주세요."
        )


def _render_order_complete(settings: dict, farm_name: str):
    """주문 완료 화면을 렌더링합니다."""
    result       = st.session_state.get("order_result", {})
    order_number = result.get("order_number", "")
    name         = result.get("name", "")
    recipients   = result.get("recipients", [])

    st.markdown(
        f"<div class='notice-box open'>"
        f"<h2>🎉 주문이 완료되었습니다!</h2>"
        f"<p style='font-size:1.1rem;'>주문번호: <strong>{order_number}</strong></p>"
        f"<p>{name}님, 주문해주셔서 감사합니다! 🍑</p>"
        f"<p style='font-size:0.9rem;color:#555;'>"
        f"주문 접수 후 곧 메시지로 안내드리겠습니다."
        f"</p>"
        f"</div>",
        unsafe_allow_html=True,
    )

    # 계좌 안내 재표시
    bank   = settings.get("bank", "농협")
    acct   = settings.get("account_number", "000-0000-0000")
    holder = settings.get("holder", "장명숙")
    st.markdown(
        f"<div class='account-box'>"
        f"<p style='margin:0 0 0.3rem 0;color:#666;'>💳 아래 계좌로 입금해 주세요</p>"
        f"<div class='bank-name'>{bank}은행</div>"
        f"<div class='account-num'>{acct}</div>"
        f"<div style='color:#555;'>예금주: <strong>{holder}</strong></div>"
        f"</div>",
        unsafe_allow_html=True,
    )

    # 주문 내역 요약
    st.markdown("#### 📋 주문 내역 요약")
    for i, rec in enumerate(recipients, 1):
        memo_html = f"<div>배송메모: {rec.get('memo','')}</div>" if rec.get("memo") else ""
        st.markdown(
            f"<div class='recipient-box'>"
            f"<div class='recipient-box-title'>📮 {i}번째 수령자: {rec['name']}</div>"
            f"<div>전화번호: {rec['phone']}</div>"
            f"<div>주소: {rec['address']}</div>"
            f"<div>상품: {rec['product']} × {rec['qty']}박스</div>"
            f"{memo_html}"
            f"</div>",
            unsafe_allow_html=True,
        )


# =============================================================================
# [B] 관리자 탭 1 — 주문 현황
# =============================================================================

def _metric_card(col, label: str, value, color: str):
    """컬럼 안에 지표 카드를 렌더링합니다."""
    with col:
        st.markdown(
            f"<div class='metric-card'>"
            f"<div class='metric-value' style='color:{color};'>{value}</div>"
            f"<div class='metric-label'>{label}</div>"
            f"</div>",
            unsafe_allow_html=True,
        )


def render_admin_orders():
    """주문 현황 탭: 지표 요약 + 편집 가능 테이블 + CSV 다운로드"""
    st.markdown("### 📊 주문 현황")

    df = load_orders()
    if df.empty:
        st.info("ℹ️ 아직 접수된 주문이 없습니다.")
        return

    # ── 지표 ──
    total   = len(df)
    waiting = int((df["상태"] == "대기").sum())    if "상태" in df.columns else 0
    confirm = int((df["상태"] == "확인").sum())    if "상태" in df.columns else 0
    shipped = int((df["상태"] == "발송완료").sum()) if "상태" in df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    _metric_card(c1, "전체 주문", total,   "#ff8c42")
    _metric_card(c2, "대기",      waiting, "#ffc107")
    _metric_card(c3, "확인",      confirm, "#2196f3")
    _metric_card(c4, "발송완료",  shipped, "#4caf50")

    st.markdown("---")
    st.caption("아래 표에서 '상태' 열을 직접 클릭하여 수정할 수 있습니다.")

    # ── 편집 가능한 데이터 테이블 ──
    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            "상태": st.column_config.SelectboxColumn(
                "상태",
                options=["대기", "확인", "발송완료"],
                required=True,
            )
        },
        key="orders_editor",
    )

    col_save, col_dl = st.columns(2)
    with col_save:
        if st.button("💾 상태 변경 저장", use_container_width=True):
            sheet = get_sheet("주문목록")
            if sheet:
                saved = 0
                try:
                    for idx in range(len(df)):
                        orig = df.iloc[idx]["상태"] if "상태" in df.columns else ""
                        new  = edited_df.iloc[idx]["상태"]
                        if orig != new:
                            sheet.update_cell(idx + 2, 12, new)  # 헤더행=1, 데이터=2+
                            saved += 1
                    if saved:
                        st.success(f"✅ {saved}건의 상태를 저장했습니다.")
                        st.cache_resource.clear()
                    else:
                        st.info("변경된 항목이 없습니다.")
                except Exception as e:
                    st.error(f"저장 실패: {e}")
            else:
                st.error("Google Sheets 연결 실패")

    with col_dl:
        csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            "📥 CSV 다운로드",
            data=csv_bytes,
            file_name=f"주문목록_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )


# =============================================================================
# [B] 관리자 탭 2 — 주문 기간 설정
# =============================================================================

def render_admin_period(settings: dict):
    """주문 기간 설정 탭: 날짜/시간 입력 + 즉시 열기/닫기"""
    st.markdown("### ⏰ 주문 기간 설정")

    status, _, _ = check_order_period(settings)
    labels = {"before": "⏳ 주문 시작 전", "open": "✅ 주문 접수 중", "closed": "🚫 주문 마감"}
    st.info(f"현재 상태: **{labels.get(status, '알 수 없음')}**")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**📅 주문 시작**")
        try:
            d_start = datetime.strptime(settings.get("order_start", "2026-07-01 09:00"), "%Y-%m-%d %H:%M")
        except ValueError:
            d_start = datetime(2026, 7, 1, 9, 0)
        start_date = st.date_input("시작 날짜", value=d_start.date(), key="pd_start_date")
        start_time = st.time_input("시작 시간", value=d_start.time(), key="pd_start_time")

    with col2:
        st.markdown("**📅 주문 마감**")
        try:
            d_end = datetime.strptime(settings.get("order_end", "2026-07-10 18:00"), "%Y-%m-%d %H:%M")
        except ValueError:
            d_end = datetime(2026, 7, 10, 18, 0)
        end_date = st.date_input("마감 날짜", value=d_end.date(), key="pd_end_date")
        end_time = st.time_input("마감 시간", value=d_end.time(), key="pd_end_time")

    st.markdown("---")
    ca, cb, cc = st.columns(3)

    with ca:
        if st.button("💾 기간 저장", use_container_width=True):
            settings["order_start"] = f"{start_date} {start_time.strftime('%H:%M')}"
            settings["order_end"]   = f"{end_date} {end_time.strftime('%H:%M')}"
            if save_settings(settings):
                st.success("✅ 주문 기간이 저장되었습니다.")
            else:
                st.error("저장에 실패했습니다. Google Sheets 연결을 확인해주세요.")

    with cb:
        if st.button("🟢 지금 바로 열기", use_container_width=True):
            now = datetime.now()
            settings["order_start"] = now.strftime("%Y-%m-%d %H:%M")
            settings["order_end"]   = (now + timedelta(days=7)).strftime("%Y-%m-%d %H:%M")
            if save_settings(settings):
                st.success("✅ 주문이 열렸습니다! (7일 후 자동 마감)")
            else:
                st.error("저장 실패")

    with cc:
        if st.button("🔴 지금 바로 닫기", use_container_width=True):
            past = datetime.now() - timedelta(minutes=1)
            settings["order_end"] = past.strftime("%Y-%m-%d %H:%M")
            if save_settings(settings):
                st.success("✅ 주문이 마감되었습니다.")
            else:
                st.error("저장 실패")


# =============================================================================
# [B] 관리자 탭 3 — 로젠택배 엑셀
# =============================================================================

def render_admin_logen(settings: dict):
    """로젠택배 엑셀 생성 탭: 필터 → 미리보기 → 다운로드 → 발송완료 처리"""
    st.markdown("### 📦 로젠택배 발송 엑셀")

    df = load_orders()
    if df.empty:
        st.info("ℹ️ 주문 데이터가 없습니다.")
        return

    # ── 필터 ──
    col1, col2 = st.columns(2)
    with col1:
        status_filter = st.multiselect(
            "상태 필터",
            options=["대기", "확인", "발송완료"],
            default=["대기", "확인"],
            key="logen_status",
        )
    with col2:
        date_range = None
        if "주문일시" in df.columns:
            try:
                df["_dt"] = pd.to_datetime(df["주문일시"], errors="coerce")
                min_d = df["_dt"].min().date()
                max_d = df["_dt"].max().date()
                date_range = st.date_input("주문일 범위", value=(min_d, max_d), key="logen_date")
            except Exception:
                pass

    # 필터 적용
    filtered = df.copy()
    if "상태" in df.columns and status_filter:
        filtered = filtered[filtered["상태"].isin(status_filter)]
    if date_range and len(date_range) == 2 and "_dt" in filtered.columns:
        s_d, e_d = date_range
        filtered = filtered[
            (filtered["_dt"].dt.date >= s_d) &
            (filtered["_dt"].dt.date <= e_d)
        ]

    st.caption(f"필터 결과: **{len(filtered)}건**")
    if filtered.empty:
        st.warning("조건에 맞는 주문이 없습니다.")
        return

    # 미리보기
    preview_cols = [c for c in ["주문번호", "받는분이름", "받는분주소", "상품명", "수량", "상태"] if c in filtered.columns]
    st.dataframe(filtered[preview_cols], use_container_width=True)

    farm_name = _get_farm_name()
    col_dl, col_mark = st.columns(2)

    with col_dl:
        excel_bytes = generate_logen_excel(filtered, farm_name, settings)
        st.download_button(
            "📥 로젠택배 엑셀 다운로드",
            data=excel_bytes,
            file_name=f"로젠택배_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_mark:
        if st.button("✅ 선택 주문 '발송완료' 처리", use_container_width=True):
            sheet = get_sheet("주문목록")
            if sheet:
                try:
                    all_rows = sheet.get_all_values()
                    target_ids = set(filtered["주문번호"].tolist())
                    updated = 0
                    for row_idx, row in enumerate(all_rows[1:], start=2):
                        if row[0] in target_ids and row[11] != "발송완료":
                            sheet.update_cell(row_idx, 12, "발송완료")
                            updated += 1
                    st.success(f"✅ {updated}건을 '발송완료' 처리했습니다.")
                except Exception as e:
                    st.error(f"처리 실패: {e}")
            else:
                st.error("Google Sheets 연결 실패")


# =============================================================================
# [B] 관리자 탭 4 — 메시지 발송 안내
# =============================================================================

def render_admin_email(settings: dict):
    """고객 공지 메시지 발송 탭 — 문자/카카오톡 발송용 문구 생성"""
    st.markdown("### 💬 고객 공지 메시지 발송")

    farm_name = _get_farm_name()
    try:
        order_url = st.secrets["app"]["order_url"]
    except Exception:
        order_url = "https://your-app.streamlit.app"

    order_start = settings.get("order_start", "2026-07-01 09:00")
    order_end   = settings.get("order_end",   "2026-07-10 18:00")
    bank        = settings.get("bank",         "농협")
    acct        = settings.get("account_number", "000-0000-0000")
    holder      = settings.get("holder",         "장명숙")

    st.info("아래 메시지를 복사하여 카카오톡, 문자 등으로 고객에게 보내주세요.")

    # ── 메시지 템플릿 ──
    st.markdown("**메시지 내용 작성**")
    default_body = (
        f"안녕하세요! 🍑\n"
        f"{farm_name}입니다.\n\n"
        f"올해도 맛있는 복숭아 주문을 받습니다.\n\n"
        f"▶ 주문 기간: {order_start} ~ {order_end}\n"
        f"▶ 주문 링크: {order_url}\n\n"
        f"입금 계좌 안내\n"
        f"─────────────\n"
        f"은행: {bank}\n"
        f"계좌: {acct}\n"
        f"예금주: {holder}\n"
        f"─────────────\n\n"
        f"신선하고 달콤한 복숭아로\n"
        f"무더운 여름 보내세요! 🍑\n\n"
        f"감사합니다.\n"
        f"{farm_name}"
    )

    body = st.text_area("메시지 본문 (수정 가능)", value=default_body, height=320, key="msg_body")

    st.markdown("---")

    # ── 고객 목록 및 전화번호 표시 ──
    customers_df = load_customers()
    if not customers_df.empty:
        st.markdown(f"**📋 고객 목록 ({len(customers_df)}명) — 전화번호 확인 후 직접 발송하세요**")
        st.dataframe(customers_df, use_container_width=True)
    else:
        st.warning("고객목록 시트에 고객 정보가 없습니다. '이름 / 전화번호' 형태로 입력해주세요.")


# =============================================================================
# [B] 관리자 탭 5 — 설정
# =============================================================================

def render_admin_settings(settings: dict):
    """농장 정보 및 앱 설정 탭"""
    st.markdown("### ⚙️ 설정")

    st.markdown("**농장 / 계좌 정보**")
    col1, col2 = st.columns(2)
    with col1:
        new_bank   = st.text_input("은행명",           value=settings.get("bank", "농협"),         key="cfg_bank")
        new_holder = st.text_input("예금주",           value=settings.get("holder", "장명숙"),     key="cfg_holder")
    with col2:
        new_acct   = st.text_input("계좌번호",         value=settings.get("account_number", ""),  key="cfg_acct")
        new_phone  = st.text_input("농장 전화번호",    value=settings.get("farm_phone", ""),       key="cfg_phone")

    if st.button("💾 설정 저장"):
        settings.update({
            "bank":           new_bank,
            "holder":         new_holder,
            "account_number": new_acct,
            "farm_phone":     new_phone,
        })
        if save_settings(settings):
            st.success("✅ 설정이 저장되었습니다.")
        else:
            st.error("저장 실패. Google Sheets 연결을 확인해주세요.")

    st.markdown("---")
    st.markdown("**secrets.toml 설정 안내**")
    st.info(
        "아래 항목은 `.streamlit/secrets.toml` 파일 (또는 Streamlit Cloud Secrets)에서 설정합니다.\n\n"
        "- `[gcp_service_account]` — Google 서비스 계정 JSON 내용\n"
        "- `[app] spreadsheet_id` — Google Sheets URL의 `/d/` 뒤 ID\n"
        "- `[app] admin_password` — 관리자 비밀번호\n"
        "- `[email]` — Gmail 계정과 앱 비밀번호\n"
        "- `[account]` — 기본 계좌 정보 (설정 시트보다 우선순위 낮음)"
    )

    st.markdown("---")
    st.markdown("**Google Sheets 연결 상태**")
    client = get_gspread_client()
    if client is not None:
        ss = get_spreadsheet()
        if ss is not None:
            st.success("✅ Google Sheets 연결 성공")
            try:
                ws_names = [ws.title for ws in ss.worksheets()]
                st.write("시트 목록:", ", ".join(ws_names))
                # 필수 시트 확인
                required = {"주문목록", "상품목록", "고객목록", "설정"}
                missing  = required - set(ws_names)
                if missing:
                    st.warning(f"누락된 시트: {', '.join(missing)}")
                else:
                    st.success("✅ 필수 시트(주문목록/상품목록/고객목록/설정) 모두 존재")
            except Exception:
                pass
        else:
            st.error("❌ 스프레드시트를 열 수 없습니다. spreadsheet_id를 확인해주세요.")
    else:
        st.error("❌ Google Sheets 연결 실패. gcp_service_account 설정을 확인해주세요.")


# =============================================================================
# [B] 관리자 페이지 — 탭 통합
# =============================================================================

def render_admin_page(settings: dict, products: list):
    """관리자 모드 전체 레이아웃: 헤더 + 5개 탭"""
    farm_name = _get_farm_name()
    st.markdown(
        f"<div class='peach-header'>"
        f"<h1>🔧 {farm_name} 관리자</h1>"
        f"<p>주문 관리 및 운영 설정 페이지</p>"
        f"</div>",
        unsafe_allow_html=True,
    )

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 주문 현황",
        "⏰ 주문 기간 설정",
        "📦 로젠택배 엑셀",
        "💬 메시지 발송",
        "⚙️ 설정",
    ])

    with tab1: render_admin_orders()
    with tab2: render_admin_period(settings)
    with tab3: render_admin_logen(settings)
    with tab4: render_admin_email(settings)
    with tab5: render_admin_settings(settings)


# =============================================================================
# 메인 진입점
# =============================================================================

def main():
    """앱 진입점: 세션 초기화 → 사이드바 로그인 → 관리자/고객 분기"""

    # 세션 상태 초기화 (최초 실행 시)
    for key, default in [
        ("order_complete",  False),
        ("order_result",    None),
        ("recipients",      None),
        ("force_customer",  False),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    # 사이드바에서 관리자 여부 판단
    is_admin = render_sidebar()

    # "고객 화면으로" 버튼 클릭 시 강제로 고객 화면 표시
    if st.session_state.get("force_customer"):
        st.session_state["force_customer"] = False
        is_admin = False

    # 설정 및 상품 목록 로드 (캐시 활용)
    settings = load_settings()
    products = load_products()

    if is_admin:
        render_admin_page(settings, products)
    else:
        render_customer_page(settings, products)


if __name__ == "__main__":
    main()
