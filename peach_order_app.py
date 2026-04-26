"""
복숭아 농장 주문 자동 관리 및 로젠택배 엑셀 변환기
====================================================
SMS Backup & Restore 앱으로 추출한 XML 파일을 업로드하면
GPT-4o가 주문 문자를 자동으로 분석하여
로젠택배 g-로지스 시스템용 엑셀 파일로 변환합니다.

실행 방법:
    pip install streamlit pandas openpyxl lxml openai
    streamlit run peach_order_app.py
"""

import streamlit as st
import pandas as pd
import json
import io
from lxml import etree
from openai import OpenAI

# ─────────────────────────────────────────────
# 페이지 기본 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="복숭아 주문 관리 & 로젠택배 변환기",
    page_icon="🍑",
    layout="wide",
)

# ─────────────────────────────────────────────
# 사이드바: OpenAI API Key 입력
# ─────────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ 설정")
    st.markdown("---")
    api_key = st.text_input(
        "🔑 OpenAI API Key",
        type="password",
        placeholder="sk-...",
        help="OpenAI 플랫폼(platform.openai.com)에서 발급받은 API 키를 입력하세요.",
    )
    st.markdown("---")
    st.markdown("### 📖 사용 방법")
    st.markdown(
        """
1. OpenAI API 키를 입력하세요
2. SMS XML 파일을 업로드하세요
3. AI가 주문을 자동 분석합니다
4. 내용을 확인·수정하세요
5. 로젠택배 엑셀을 다운로드하세요
"""
    )
    st.markdown("---")
    st.caption("💡 SMS Backup & Restore 앱으로 추출한 XML 파일을 사용하세요.")

# ─────────────────────────────────────────────
# 메인 화면 타이틀
# ─────────────────────────────────────────────
st.title("🍑 복숭아 주문 자동 관리 & 로젠택배 변환기")
st.markdown(
    "스마트폰 문자(SMS) 백업 XML 파일을 업로드하면 AI가 주문을 자동으로 분석하여 로젠택배 송장 엑셀 파일을 만들어 드립니다."
)
st.markdown("---")


# ─────────────────────────────────────────────
# [함수 1] XML 파싱: <sms> 태그에서 문자 목록 추출
# ─────────────────────────────────────────────
def parse_sms_xml(xml_bytes: bytes) -> list[dict]:
    """
    SMS Backup & Restore XML 파일에서 문자 메시지를 파싱합니다.

    반환값: [{"address": "010-xxxx-xxxx", "body": "문자 내용"}, ...]
    """
    messages = []
    try:
        root = etree.fromstring(xml_bytes)
        # <sms> 태그를 모두 찾아 순회
        for sms in root.iter("sms"):
            address = sms.get("address", "").strip()
            body = sms.get("body", "").strip()
            # 발신번호와 본문이 있는 경우만 수집
            if address and body:
                messages.append({"address": address, "body": body})
    except Exception as e:
        st.error(f"XML 파싱 오류: {e}")
    return messages


# ─────────────────────────────────────────────
# [함수 2] AI 필터링: 주문 관련 문자만 선별
# ─────────────────────────────────────────────
def filter_order_messages(client: OpenAI, messages: list[dict]) -> list[dict]:
    """
    GPT-4o를 사용하여 전체 문자 목록 중 '주문 관련' 문자만 골라냅니다.
    한 번의 API 호출로 처리해 비용을 최소화합니다.
    """
    if not messages:
        return []

    # AI에게 넘길 문자 목록을 번호와 함께 텍스트로 만들기
    msg_text = "\n\n".join(
        [f"[{i}] 발신번호: {m['address']}\n내용: {m['body']}" for i, m in enumerate(messages)]
    )

    prompt = f"""당신은 복숭아 농장의 택배 주문 관리 시스템입니다.
아래 문자 메시지 목록에서 복숭아(또는 과일) 구매/주문과 관련된 메시지의 번호(인덱스)만 JSON 배열로 반환하세요.

판단 기준:
- 포함: 복숭아/과일 주문, 수량 언급, 배송 주소 포함, 선물 배송 요청 등
- 제외: 인증번호, 광고, 안부인사, 단순 문의(가격만 묻는 경우), 배송 완료 알림 등

반환 형식 (JSON만 반환, 다른 설명 없이):
{{"order_indices": [0, 3, 5, ...]}}

문자 목록:
{msg_text}
"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
        temperature=0,
    )

    result = json.loads(response.choices[0].message.content)
    order_indices = result.get("order_indices", [])

    # 인덱스가 유효한 것만 필터링
    return [messages[i] for i in order_indices if 0 <= i < len(messages)]


# ─────────────────────────────────────────────
# [함수 3] AI 구조화: 주문 정보 JSON 추출
# ─────────────────────────────────────────────
def extract_order_info(client: OpenAI, address: str, body: str) -> list[dict]:
    """
    단일 주문 문자에서 구조화된 주문 정보를 추출합니다.
    한 명이 여러 곳에 선물하는 경우 복수의 row를 반환합니다.

    반환값: [
        {
            "받는사람이름": ...,
            "받는사람주소": ...,
            "받는사람전화번호": ...,
            "상품구분": ...,
            "수량(박스)": ...,
            "보내는사람이름": ...,
            "보내는사람주소": ...,
            "보내는사람전화번호": ...,
            "특이사항": ...
        },
        ...
    ]
    """
    prompt = f"""당신은 복숭아 농장의 택배 주문 분석 전문가입니다.
아래 SMS 문자에서 주문 정보를 추출하여 JSON으로 반환하세요.

규칙:
1. 보내는 사람(주문자) 정보와 받는 사람(수령자) 정보를 구분하세요.
2. 주소는 도로명 주소를 우선으로 하세요. 지번 주소만 있으면 그대로 사용하세요.
3. 한 명이 여러 주소로 보내는 경우, "recipients" 배열에 각각 넣으세요.
4. 상품구분은 "선물용 복숭아 X박스" 또는 "일반 복숭아 X박스" 형태로 정리하세요.
5. 특이사항에는 배송 메모(예: "경비실 맡겨주세요", "부재시 문앞")를 넣으세요.
6. 알 수 없는 정보는 빈 문자열("")로 남기세요.
7. 수량은 숫자만 입력하세요 (예: 2).

발신번호: {address}
문자내용:
{body}

반환 형식 (JSON만 반환):
{{
  "sender": {{
    "이름": "",
    "주소": "",
    "전화번호": "{address}"
  }},
  "recipients": [
    {{
      "이름": "",
      "주소": "",
      "전화번호": "",
      "상품구분": "",
      "수량": 1,
      "특이사항": ""
    }}
  ]
}}
"""

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
        temperature=0,
    )

    data = json.loads(response.choices[0].message.content)
    sender = data.get("sender", {})
    recipients = data.get("recipients", [])

    # 받는 사람별로 별도 행(row) 생성
    rows = []
    for r in recipients:
        rows.append(
            {
                "받는사람이름": r.get("이름", ""),
                "받는사람주소": r.get("주소", ""),
                "받는사람전화번호": r.get("전화번호", ""),
                "상품구분": r.get("상품구분", "복숭아"),
                "수량(박스)": r.get("수량", 1),
                "보내는사람이름": sender.get("이름", ""),
                "보내는사람주소": sender.get("주소", ""),
                "보내는사람전화번호": sender.get("전화번호", address),
                "특이사항": r.get("특이사항", ""),
            }
        )

    return rows


# ─────────────────────────────────────────────
# [함수 4] 엑셀 변환: 로젠택배 양식으로 저장
# ─────────────────────────────────────────────
def convert_to_excel(df: pd.DataFrame) -> bytes:
    """
    DataFrame을 로젠택배 g-로지스 업로드용 엑셀 파일로 변환합니다.
    컬럼 순서를 로젠택배 규격에 맞게 정렬합니다.
    """
    # 로젠택배 업로드 전용 컬럼 순서
    column_order = [
        "받는사람이름",
        "받는사람주소",
        "받는사람전화번호",
        "상품구분",
        "수량(박스)",
        "보내는사람이름",
        "보내는사람주소",
        "보내는사람전화번호",
        "특이사항",
    ]

    # 순서에 맞게 컬럼 재배열 (없는 컬럼은 빈 열로 추가)
    for col in column_order:
        if col not in df.columns:
            df[col] = ""
    df_out = df[column_order]

    # 메모리 내 엑셀 파일 생성 (파일 저장 없이 바이트로 반환)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="로젠택배_송장")

        # 엑셀 열 너비 자동 조정
        worksheet = writer.sheets["로젠택배_송장"]
        for col_idx, col_name in enumerate(df_out.columns, start=1):
            max_len = max(
                df_out[col_name].astype(str).map(len).max(),
                len(col_name),
            ) + 4
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            worksheet.column_dimensions[col_letter].width = min(max_len, 40)

    return output.getvalue()


# ─────────────────────────────────────────────
# 세션 상태 초기화
# (Streamlit은 재실행 시 상태가 초기화되므로 session_state에 저장)
# ─────────────────────────────────────────────
if "order_df" not in st.session_state:
    st.session_state.order_df = None  # 편집 가능한 DataFrame
if "processing_done" not in st.session_state:
    st.session_state.processing_done = False

# ─────────────────────────────────────────────
# [Step 1] XML 파일 업로드
# ─────────────────────────────────────────────
st.subheader("📂 1단계: SMS XML 파일 업로드")
uploaded_file = st.file_uploader(
    "SMS Backup & Restore로 추출한 XML 파일을 업로드하세요",
    type=["xml"],
    help="갤럭시 스마트폰에서 'SMS Backup & Restore' 앱으로 백업한 XML 파일",
)

if uploaded_file and api_key:
    xml_bytes = uploaded_file.read()

    # ── XML 파싱 ──
    with st.spinner("📖 XML 파일을 읽는 중..."):
        all_messages = parse_sms_xml(xml_bytes)

    if not all_messages:
        st.error("⚠️ XML 파일에서 문자 메시지를 찾을 수 없습니다. 파일 형식을 확인하세요.")
        st.stop()

    st.success(f"✅ 총 **{len(all_messages)}개**의 문자를 읽었습니다.")

    # 분석 시작 버튼
    if st.button("🤖 AI 분석 시작", type="primary", use_container_width=True):
        client = OpenAI(api_key=api_key)

        # ─────────────────────────────────────────────
        # [Step 2] 주문 관련 문자 필터링
        # ─────────────────────────────────────────────
        with st.spinner("🔍 AI가 주문 문자를 찾는 중... (잠시 기다려 주세요)"):
            order_messages = filter_order_messages(client, all_messages)

        if not order_messages:
            st.warning("⚠️ 주문 관련 문자를 찾지 못했습니다. 파일 내용을 확인해 주세요.")
            st.stop()

        st.info(f"📬 총 {len(all_messages)}개 중 **{len(order_messages)}개의 주문 문자**를 발견했습니다.")

        # ─────────────────────────────────────────────
        # [Step 3] 주문 정보 추출 (문자별 AI 분석)
        # ─────────────────────────────────────────────
        all_rows = []
        progress_bar = st.progress(0, text="주문 내용 분석 중...")

        for i, msg in enumerate(order_messages):
            try:
                rows = extract_order_info(client, msg["address"], msg["body"])
                all_rows.extend(rows)
            except Exception as e:
                st.warning(f"⚠️ {msg['address']} 문자 분석 실패: {e}")

            # 진행률 업데이트
            progress_bar.progress(
                (i + 1) / len(order_messages),
                text=f"분석 중... ({i+1}/{len(order_messages)})",
            )

        progress_bar.empty()

        if not all_rows:
            st.error("❌ 주문 정보를 추출하지 못했습니다.")
            st.stop()

        # DataFrame으로 변환 후 세션에 저장
        st.session_state.order_df = pd.DataFrame(all_rows)
        st.session_state.processing_done = True
        st.success(f"🎉 분석 완료! 총 **{len(all_rows)}건**의 주문을 추출했습니다.")

elif uploaded_file and not api_key:
    st.warning("⬅️ 왼쪽 사이드바에서 OpenAI API 키를 먼저 입력해 주세요.")

elif not uploaded_file:
    # 예시 화면 (파일 미업로드 시)
    st.info("👆 XML 파일을 업로드하면 분석을 시작할 수 있습니다.")
    with st.expander("📝 지원하는 문자 형식 예시 보기"):
        st.markdown(
            """
**예시 1 - 일반 주문:**
```
안녕하세요, 복숭아 2박스 주문할게요.
받는 사람: 홍길동, 010-1234-5678
주소: 서울시 강남구 테헤란로 123
메모: 경비실에 맡겨주세요
```

**예시 2 - 선물용 다수 배송:**
```
선물용 복숭아 보내주세요!
1. 김철수 / 010-2222-3333 / 부산시 해운대구 마린시티1로 1 / 1박스
2. 이영희 / 010-4444-5555 / 대구시 수성구 달구벌대로 111 / 1박스
보내는 사람: 박민준 / 010-9999-8888
```
"""
        )

# ─────────────────────────────────────────────
# [Step 4] 데이터 편집기 표시
# ─────────────────────────────────────────────
if st.session_state.processing_done and st.session_state.order_df is not None:
    st.markdown("---")
    st.subheader("✏️ 2단계: 주문 내용 확인 및 수정")
    st.markdown("아래 표에서 잘못된 내용을 직접 클릭하여 수정할 수 있습니다.")

    # st.data_editor: 셀 직접 편집 가능
    edited_df = st.data_editor(
        st.session_state.order_df,
        use_container_width=True,
        num_rows="dynamic",  # 행 추가/삭제 가능
        column_config={
            "수량(박스)": st.column_config.NumberColumn(
                "수량(박스)",
                min_value=1,
                max_value=100,
                step=1,
                format="%d 박스",
            ),
            "받는사람전화번호": st.column_config.TextColumn("받는사람전화번호"),
            "보내는사람전화번호": st.column_config.TextColumn("보내는사람전화번호"),
            "받는사람주소": st.column_config.TextColumn("받는사람주소", width="large"),
            "상품구분": st.column_config.TextColumn("상품구분", width="medium"),
        },
        height=400,
    )

    # 편집된 데이터 세션에 저장
    st.session_state.order_df = edited_df

    # 요약 통계
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("총 배송 건수", f"{len(edited_df)}건")
    with col2:
        total_boxes = edited_df["수량(박스)"].sum() if "수량(박스)" in edited_df.columns else 0
        st.metric("총 박스 수", f"{int(total_boxes)}박스")
    with col3:
        unique_senders = edited_df["보내는사람전화번호"].nunique() if "보내는사람전화번호" in edited_df.columns else 0
        st.metric("주문자 수", f"{unique_senders}명")

    # ─────────────────────────────────────────────
    # [Step 5] 로젠택배 엑셀 다운로드
    # ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📥 3단계: 로젠택배 엑셀 다운로드")
    st.markdown("아래 버튼을 클릭하면 **로젠택배 g-로지스** 시스템에 바로 업로드할 수 있는 엑셀 파일이 다운로드됩니다.")

    excel_bytes = convert_to_excel(edited_df.copy())

    st.download_button(
        label="📦 로젠택배 양식 다운로드 (.xlsx)",
        data=excel_bytes,
        file_name="로젠택배_복숭아주문.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.caption("💡 다운로드된 파일을 로젠택배 g-로지스(g-logis.co.kr) → 일괄 송장 등록 메뉴에서 업로드하세요.")
