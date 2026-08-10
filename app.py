import streamlit as st
import google.generativeai as genai
import json
import os
import io
import re
from docx import Document

# ==========================================
# 1. 환경 설정 및 보안
# ==========================================
st.set_page_config(page_title="재무제표 AI 분석 시스템", page_icon="📊", layout="wide")

st.sidebar.title("🔑 시스템 설정")
api_key_input = st.sidebar.text_input("Gemini API 키를 입력하세요", type="password")
API_KEY = api_key_input or st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")

if not API_KEY:
    st.info("👈 왼쪽 사이드바에 발급받은 Gemini API 키를 입력하면 시스템이 활성화됩니다.")
    st.stop()

genai.configure(api_key=API_KEY)

# ==========================================
# 2. 유틸리티 함수
# ==========================================
def safe_div(numerator, denominator):
    try:
        if denominator == 0 or denominator is None:
            return 0.0
        return float(numerator) / float(denominator)
    except (TypeError, ValueError):
        return 0.0

def extract_financial_data(pdf_bytes):
    model = genai.GenerativeModel('gemini-3.6-flash') 
    prompt = """
    당신은 20년 경력의 재무 분석가입니다. 첨부된 재무제표 PDF를 읽고 다음 9가지 항목의 숫자를 추출하세요:
    1. 총자산, 2. 총부채, 3. 자기자본(자본총계), 4. 유동자산, 5. 유동부채, 6. 장단기차입금(장기차입금+단기차입금), 7. 매출액, 8. 영업이익, 9. 이자비용
    [조건]
    - 단위(예: 백만원)를 파악하고, 무조건 '1원 단위의 절대금액(정수)'으로 변환하세요.
    - 반드시 아래의 순수 JSON 형식으로만 답변하세요.
    {"총자산": 0, "총부채": 0, "자기자본": 0, "유동자산": 0, "유동부채": 0, "장단기차입금": 0, "매출액": 0, "영업이익": 0, "이자비용": 0}
    """
    response = model.generate_content([{"mime_type": "application/pdf", "data": pdf_bytes}, prompt])
    try:
        clean_text = response.text.replace("```json", "").replace("```", "").strip()
        return json.loads(clean_text)
    except:
        return None

def generate_financial_report(metrics_data):
    model = genai.GenerativeModel('gemini-3.6-flash')
    debt_ratio_str = metrics_data['부채비율'] if isinstance(metrics_data['부채비율'], str) else f"{metrics_data['부채비율']:.1f}%"
    prompt = f"""
    당신은 기업의 재무 건전성을 평가하는 시니어 재무 분석가입니다. 
    다음 산출된 5대 재무 지표를 심층 진단하세요.
    [산출 데이터]
    - 부채비율: {debt_ratio_str}
    - 유동비율: {metrics_data['유동비율']:.1f}%
    - 차입금의존도: {metrics_data['차입금의존도']:.1f}%
    - 매출액영업이익률: {metrics_data['매출액영업이익률']:.1f}%
    - 이자보상배율: {metrics_data['이자보상배율']:.1f}배

    [작성 규칙]
    1. 각 지표별로 "양호", "보통", "주의", "위험", "적자", "부도 위험", "자본잠식" 중 가장 적합한 키워드를 골라 반드시 `- [키워드]` 형태로 요약 결론을 내려주세요.
    2. 긍정적 내용은 `:blue[파란색]`, 부정적 내용은 `:red[빨간색]`, 중간은 `:orange[주황색]` Streamlit 태그로 시각적 강조를 해주세요.

    [작성 예시 - 반드시 이 형식을 지키세요!]
    (1) 부채비율: 292.5% - :red[위험]
    - 기준: 일반 제조업 권장 기준 200% 이하
    - 진단: 권장 기준을 크게 초과하여 타인자본 의존도가 매우 높은 :red[불안정한 상태]입니다.
    (2) 유동비율: 109.2% - :orange[주의]
    - 기준: 일반 제조업 권장 기준 150% 이상
    - 진단: 100%를 겨우 넘어 단기 채무를 가까스로 상환할 수 있는 :orange[아슬아슬한 유동성 수준]입니다.

    한국통계시스템의 최신 중견기업 제조업 평균 통계를 모를 경우 "일반적인 제조업 재무 건전성 기준"으로 평가한다고 명시하세요.
    """
    response = model.generate_content(prompt)
    return response.text

def extract_status_from_report(report_text, metric_name):
    pattern = re.compile(f".*?{metric_name}.*?(-|:).*?(:red|:blue|:orange)\[(.*?)\]", re.DOTALL)
    match = pattern.search(report_text)
    if match:
        status_text = match.group(3)
        if "위험" in status_text or "적자" in status_text or "부도" in status_text or "주의" in status_text:
             return f"🚨 {status_text}"
        elif "양호" in status_text or "안전" in status_text:
             return f"🔵 {status_text}"
        else:
             return f"🟠 {status_text}"
    return "⚪️ 분석중"

def generate_word_document(metrics, report_text):
    """워드 문서 생성 함수 (텍스트 래핑 지원)"""
    doc = Document()
    doc.add_heading('📊 재무제표 AI 분석 & 심층 진단 리포트', 0)
    
    doc.add_heading('1. 핵심 재무 지표 요약', level=1)
    
    debt_str = metrics['부채비율'] if isinstance(metrics['부채비율'], str) else f"{metrics['부채비율']:.1f}%"
    doc.add_paragraph(f"• 부채비율: {debt_str}")
    doc.add_paragraph(f"• 유동비율: {metrics['유동비율']:.1f}%")
    doc.add_paragraph(f"• 차입금의존도: {metrics['차입금의존도']:.1f}%")
    doc.add_paragraph(f"• 매출액영업이익률: {metrics['매출액영업이익률']:.1f}%")
    doc.add_paragraph(f"• 이자보상배율: {metrics['이자보상배율']:.2f}배")
    
    doc.add_heading('2. AI 심층 진단 상세 내역', level=1)
    
    # 웹용 색상 태그(:red[위험] 등)를 워드 문서에 맞게 깔끔한 텍스트로 변환
    clean_report = re.sub(r':(red|blue|orange)\[(.*?)\]', r'\2', report_text)
    doc.add_paragraph(clean_report)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# 3. UI 및 워크플로우 구성
# ==========================================
fields = ["총자산", "총부채", "자기자본", "유동자산", "유동부채", "장단기차입금", "매출액", "영업이익", "이자비용"]
for field in fields:
    if field not in st.session_state:
        st.session_state[field] = 0

st.title("📊 재무제표 AI 분석 & 5대 지표 진단 시스템")
st.warning("⚠️ 주의: 본 시스템에는 DART(전자공시시스템) 등에 공시된 공개용 재무제표 PDF만 업로드해 주시기 바랍니다.")

st.subheader("1. 재무제표 PDF 업로드")
uploaded_file = st.file_uploader("협력사 재무제표 PDF 파일을 업로드하세요", type=["pdf"])
if uploaded_file is not None:
    if st.button("📄 AI로 숫자 자동 추출하기"):
        with st.spinner("AI가 문서를 분석 중입니다..."):
            extracted_data = extract_financial_data(uploaded_file.read())
            if extracted_data:
                for field in fields:
                    st.session_state[field] = extracted_data.get(field, 0)
                st.success("추출 완료! 아래에서 숫자를 확인하세요.")

st.markdown("---")

st.subheader("2. 재무 데이터 검증 및 수정 (단위: 원)")
cols = st.columns(3)
for idx, field in enumerate(fields):
    col = cols[idx % 3]
    st.session_state[field] = col.number_input(
        f"{field}", 
        value=int(st.session_state[field]), 
        step=1000000,
        format="%d"
    )
    col.caption(f"💸 **{st.session_state[field]:,}** 원")

st.markdown("---")

st.subheader("3. 5대 재무 지표 심층 분석")
if st.button("🚀 5대 지표 심층 분석 실행", type="primary"):
    with st.spinner("AI가 재무 리포트를 작성 중입니다..."):
        ta, tl, eq = st.session_state["총자산"], st.session_state["총부채"], st.session_state["자기자본"]
        ca, cl, borr = st.session_state["유동자산"], st.session_state["유동부채"], st.session_state["장단기차입금"]
        rev, op, ie = st.session_state["매출액"], st.session_state["영업이익"], st.session_state["이자비용"]
        
        debt_ratio_val = "자본잠식" if eq < 0 else safe_div(tl, eq) * 100
        metrics = {
            "부채비율": debt_ratio_val,
            "유동비율": safe_div(ca, cl) * 100,
            "차입금의존도": safe_div(borr, ta) * 100,
            "매출액영업이익률": safe_div(op, rev) * 100,
            "이자보상배율": safe_div(op, ie)
        }
        
        report = generate_financial_report(metrics)
        st.session_state['report'] = report 
        
        st.write("### 🧮 핵심 지표 계산 결과 (AI 진단 연동)")
        c1, c2, c3, c4, c5 = st.columns(5)
        
        c1.metric("부채비율", debt_ratio_val if isinstance(debt_ratio_val, str) else f"{debt_ratio_val:.1f}%", extract_status_from_report(report, "부채비율"), delta_color="off")
        c2.metric("유동비율", f"{metrics['유동비율']:.1f}%", extract_status_from_report(report, "유동비율"), delta_color="off")
        c3.metric("차입금의존도", f"{metrics['차입금의존도']:.1f}%", extract_status_from_report(report, "차입금의존도"), delta_color="off")
        c4.metric("영업이익률", f"{metrics['매출액영업이익률']:.1f}%", extract_status_from_report(report, "매출액영업이익률"), delta_color="off")
        c5.metric("이자보상배율", f"{metrics['이자보상배율']:.2f}배", extract_status_from_report(report, "이자보상배율"), delta_color="off")
        
        st.markdown("---")
        st.write("### 🤖 AI 심층 진단 리포트 (색상 시각화 적용)")
        st.write(report)
        
        word_data = generate_word_document(metrics, report)
        
        st.download_button(
            label="📄 분석 결과 워드(Word) 리포트 다운로드",
            data=word_data,
            file_name="재무분석_심층리포트.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
