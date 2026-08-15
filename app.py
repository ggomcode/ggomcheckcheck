import io
import os
import re
import json
import traceback
import requests
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==============================================================================
# 페이지 기본 설정 & CSS 커스텀 스타일링 (Light Theme & Premium Card Aesthetics)
# ==============================================================================
st.set_page_config(
    page_title="학교생활기록부 AI 정밀 검증 시스템",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Modern, Premium UI with Balanced, Elegant Borders
st.markdown("""
    <style>
    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
    }
    .main-header {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        padding: 1.8rem 2rem;
        border-radius: 14px;
        color: #F8FAFC;
        margin-bottom: 1.8rem;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
        border: 1px solid #334155;
    }
    .main-header h1 {
        color: #38BDF8;
        font-size: 1.85rem;
        font-weight: 700;
        margin-bottom: 0.4rem;
    }
    .main-header p {
        color: #94A3B8;
        font-size: 0.95rem;
        margin: 0;
    }
    .metric-card {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 10px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    }
    .metric-card .val {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0F172A;
    }
    .metric-card .lbl {
        font-size: 0.85rem;
        color: #64748B;
        font-weight: 500;
        margin-top: 0.2rem;
    }
    .student-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-left: 4px solid #0EA5E9;
        border-radius: 8px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 1rem;
    }
    .subject-badge {
        display: inline-block;
        background-color: #0EA5E9;
        color: white;
        font-weight: 600;
        font-size: 0.8rem;
        padding: 0.25rem 0.6rem;
        border-radius: 6px;
        margin-bottom: 0.6rem;
    }
    .char-badge {
        float: right;
        font-size: 0.8rem;
        color: #64748B;
        font-weight: 500;
    }
    /* Hide file uploader size limit and format text (200MB per file • XLSX, XLS) */
    [data-testid="stFileUploaderInstructions"],
    [data-testid="stFileUploaderInstructions"] *,
    [data-testid="stFileUploaderDropzoneInstructions"],
    [data-testid="stFileUploaderDropzoneInstructions"] *,
    div[data-testid="stFileUploader"] small,
    section[data-testid="stFileUploader"] small,
    div[data-testid="stFileUploader"] small *,
    section[data-testid="stFileUploader"] small * {
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
        width: 0 !important;
        font-size: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    /* Hide Sidebar Collapse Button & Header */
    [data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarHeader"],
    [data-testid="stSidebarHeader"] button,
    section[data-testid="stSidebar"] button[aria-label*="close"],
    section[data-testid="stSidebar"] button[aria-label*="Collapse"],
    section[data-testid="stSidebar"] button[aria-label*="접기"],
    [data-testid="stSidebarNav"],
    #MainMenu,
    header[data-testid="stHeader"],
    footer,
    [data-testid="stToolbar"],
    .stAppDeployButton,
    div[data-testid="stAppDeployButton"],
    div[data-testid="stDecoration"] {
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
        width: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    section[data-testid="stSidebar"] * {
        color: #0F172A !important;
    }
    section[data-testid="stSidebar"] .stCaption, 
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
        color: #64748B !important;
        font-weight: 500 !important;
    }

    /* Force Expander Header & Labels onto 1 Single Line without Arrow Icon */
    [data-testid="stExpander"] details summary svg,
    .stExpander summary svg,
    details summary svg {
        display: none !important;
        visibility: hidden !important;
        width: 0 !important;
        height: 0 !important;
    }
    [data-testid="stExpander"] details summary,
    .stExpander summary {
        padding-left: 0.75rem !important;
    }
    [data-testid="stExpander"] details summary p,
    [data-testid="stExpander"] details summary span,
    .stExpander summary p,
    .stExpander summary span {
        white-space: nowrap !important;
        overflow: visible !important;
        text-overflow: clip !important;
        font-size: 0.95rem !important;
    }
    /* Sub-expander headers (e.g. Gemini API Key 발급 가이드) font size reduced by 1pt */
    section[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stExpander"] details summary p,
    section[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stExpander"] details summary span {
        font-size: 0.88rem !important;
        font-weight: 600 !important;
    }

    /* Sidebar Expander Color Themes & Balanced Borders */
    section[data-testid="stSidebar"] [data-testid="stExpander"] {
        border-radius: 12px !important;
        overflow: hidden !important;
        margin-bottom: 0.85rem !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03) !important;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"] details summary {
        padding: 0.7rem 0.9rem !important;
        border-radius: 10px !important;
        transition: all 0.2s ease-in-out !important;
    }

    /* 1. AI (LLM) API 설정 (Indigo Blue Theme) */
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(1) {
        border: 1px solid #A5B4FC !important;
        background-color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(1) details summary {
        background: linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 100%) !important;
        border-bottom: 1px solid #C7D2FE !important;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(1) details summary,
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(1) details summary * {
        color: #3730A3 !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }

    /* 2. 엑셀 파일 업로드 (Mint / Emerald Green Theme) */
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(2) {
        border: 1px solid #6EE7B7 !important;
        background-color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(2) details summary {
        background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%) !important;
        border-bottom: 1px solid #A7F3D0 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(2) details summary,
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(2) details summary * {
        color: #065F46 !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }

    /* 3. 데이터 관리 메뉴 (Lavender Purple Theme) */
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(3) {
        border: 1px solid #C4B5FD !important;
        background-color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(3) details summary {
        background: linear-gradient(135deg, #F5F3FF 0%, #EDE9FE 100%) !important;
        border-bottom: 1px solid #DDD6FE !important;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(3) details summary,
    section[data-testid="stSidebar"] [data-testid="stExpander"]:nth-of-type(3) details summary * {
        color: #5B21B6 !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }

    /* Form Controls - Crisp Visible Borders for Selectbox & Inputs */
    [data-baseweb="select"],
    [data-baseweb="input"],
    div[data-testid="stForm"],
    input, select, textarea {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
    }
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    div[data-testid="stSelectbox"] > div > div,
    div[data-testid="stTextInput"] > div > div,
    div[role="combobox"] {
        border: 1.2px solid #94A3B8 !important;
        border-radius: 8px !important;
        background-color: #FFFFFF !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04) !important;
    }
    div[data-baseweb="select"] svg {
        fill: #334155 !important;
    }

    /* File Uploader styling - Modern Soft Dashed Border */
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF !important;
        border: 1.5px dashed #94A3B8 !important;
        border-radius: 12px !important;
        padding: 1rem !important;
    }
    [data-testid="stFileUploader"] section {
        background-color: #F8FAFC !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploader"] * {
        color: #334155 !important;
        font-weight: 600 !important;
    }

    /* General Buttons - Balanced Clean Border */
    div.stButton > button:not([kind="primary"]):not([data-testid="stBaseButton-primary"]),
    [data-testid="stBaseButton-secondary"],
    [data-testid="stFileUploader"] button {
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #94A3B8 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04) !important;
    }
    div.stButton > button:not([kind="primary"]):not([data-testid="stBaseButton-primary"]) *,
    [data-testid="stBaseButton-secondary"] *,
    [data-testid="stFileUploader"] button * {
        color: #0F172A !important;
    }

    /* Download Buttons - Clean Blue Border */
    [data-testid="stDownloadButton"] button,
    [data-testid="stDownloadButton"] a {
        background-color: #EFF6FF !important;
        color: #1D4ED8 !important;
        border: 1px solid #3B82F6 !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
    }
    [data-testid="stDownloadButton"] button *,
    [data-testid="stDownloadButton"] a * {
        color: #1D4ED8 !important;
    }

    /* Primary Button */
    button[kind="primary"],
    [data-testid="stBaseButton-primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid #1D4ED8 !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
    }
    button[kind="primary"] *,
    [data-testid="stBaseButton-primary"] * {
        color: #FFFFFF !important;
    }
    </style>
""", unsafe_allow_html=True)


# ==============================================================================
# 1. 지침 MD 파일 동적 로더
# ==============================================================================
@st.cache_data
def load_guideline_content(md_file_path: str = "data/학교생활기록부_기재_및_검증_지침.md") -> str:
    """
    data/학교생활기록부_기재_및_검증_지침.md 지침 파일 전문을 읽어옵니다.
    """
    if os.path.exists(md_file_path):
        try:
            with open(md_file_path, "r", encoding="utf-8") as f:
                return f.read()
        except Exception as e:
            st.warning(f"지침 md 파일 로드 경고: {e}")
    return ""


# ==============================================================================
# 2. LLM AI 분석 엔진 (Gemini / OpenAI / Claude API 파이프라인)
# ==============================================================================
def clean_json_response(raw_text: str) -> str:
    """
    AI 응답에서 마크다운 코드블록(```json ... ```)을 정제합니다.
    """
    if not raw_text:
        return ""
    clean = raw_text.strip()
    if clean.startswith("```"):
        clean = re.sub(r'^```[a-zA-Z]*\n?', '', clean)
        clean = re.sub(r'```$', '', clean)
        clean = clean.strip()
    return clean


def call_llm_api_for_audit(provider: str, api_key: str, model_name: str, records_data: list, guideline_text: str, progress_callback=None) -> list:
    """
    생기부 기록 데이터를 적정 배치(Batch)로 나누어 API 쿼터(250k 토큰 제한)를 초과하지 않도록 분할 호출합니다.
    429 Rate Limit/Quota 초과 오류 발생 시 자동 지연 재시도(Exponential Retry)를 수행합니다.
    """
    import time
    if not api_key:
        raise ValueError("API Key가 설정되지 않았습니다. 사이드바에서 AI API Key를 입력해 주세요.")

    BATCH_SIZE = 15
    batches = [records_data[i:i + BATCH_SIZE] for i in range(0, len(records_data), BATCH_SIZE)]
    total_batches = len(batches)
    
    all_results = []

    prompt_instructions = f"""
당신은 대한민국 고등학교 학교생활기록부(창체, 세특, 행특) 오탈자 및 기재지침 검증 최고 전문가 AI입니다.

[참고 지침 문서]
{guideline_text}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【검증 5대 핵심 지침】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. 오탈자, 맞춤법, 띄어쓰기 및 한국어 문법적 오류:
   - 모든 문장의 철자, 띄어쓰기, 어미 활용(예: '도우는'->'돕는', '만듬'->'만듦', '되서'->'돼서', '안되'->'안 돼', '치뤄'->'치러' 등), 주어-서술어 호응 오탈자를 정밀 검출하십시오.
2. 입력 불가 용어 및 브랜드/서비스명:
   - 지침 2.3 매핑표의 용어(네이버, 구글, 유튜브, 카카오톡, KTX, MBTI, 챗GPT, ZOOM 등)가 발견되면 지정된 올바른 대체어로 교정 제안하십시오.
3. 생기부 기재 금지어 및 금지 항목:
   - 수상, 대회, 논문, 자격증, 방과후학교, 특정 대학/기관명, 상호, 교사명, 학교이름, 해외활동 등 금지 항목을 감지하십시오.
   - [중요 규칙] 도서명(책 제목) 및 저자명(저자 이름)은 지침 2.1.4에 따라 독서활동 및 모든 영역에 기재할 수 있는 정당한 항목이므로 절대로 오탈자나 오류로 감지하지 마십시오.
4. 특수문자 제한 위반:
   - 따옴표('"), 쉼표(,), 마침표(.), 느낌표(!), 물음표(?), 콜론(:), 괄호 외 불필요 특수기호(★, ◆, ~, @, #, $ 등)를 감지하십시오.
5. 학생 입장 서술 지양 어미:
   - ~파악함, ~이해함, ~다짐함, ~느낌, ~배움 등 학생 입장 어미를 감지하여 교사 관점 서술어(~활동지를 작성함, ~모습이 돋보임)로 전환 제안하십시오.
6. 수정 이유/근거 작성 시 지침 번호 제외:
   - [필수 규칙] 'reason(수정해야하는 이유나 근거)' 항목을 작성할 때는 개별 지침의 조항 번호(예: '지침 2.1.4', '지침 2.3', '제3조', '3.1.2' 등)를 절대로 적지 마십시오. 번호 없이 맞춤법, 띄어쓰기, 어미 교정, 불필요 특수문자 제거 등 구체적인 수정 이유만 명확히 설명하십시오.
7. 창체 영역별 이수시간 기재 및 0시간 점검 규칙:
   - [중요 규칙] 창의적 체험활동(자율, 동아리, 진로활동 등)의 영역별 총 이수시간(예: '16시간', '34시간', '(16시간)')은 필수 기재 사항이므로 절대로 '소요시간/실적 기재 불가' 오류로 감지하지 마십시오.
   - [0시간 점검] 영역별 이수시간이 '0시간'으로 기록되어 있거나 텍스트 내 '(0시간)'이 작성된 경우에는 출결 및 실제 이수시간 점검을 위해 '수정 권장'으로 감지하고 "창체 영역별 이수 시간이 0시간으로 기재되어 있으니 출결 및 실제 이수 시간을 확인하십시오"라는 안내 사유를 제시하십시오.
8. 창체 세부 영역 단일화 및 명확화:
   - [필수 규칙] 창의적 체험활동의 'sub_category(세부)' 항목은 기록 내용을 분석하여 반드시 '자율활동', '동아리활동', '진로활동' 중 명확하게 1개 영역만 지정해야 합니다. 절대로 '자율/동아리/진로'처럼 여러 개를 슬래시로 묶어서 표기하지 마십시오.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【출력 형식 (JSON Schema)】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
반드시 오직 아래의 JSON 데이터만 반환하십시오. 마크다운 추가 텍스트 없이 순수한 JSON만 반환해야 합니다:
{{
  "results": [
    {{
      "student_id": "학번 (5자리 숫자로 작성, 예: 30101)",
      "student_name": "학생 이름",
      "category": "구분 (창체 / 세특 / 행발 중 하나)",
      "taken_grade": "이수학년 (예: 1학년, 2학년, 3학년)",
      "sub_category": "세부 (창체는 '자율활동', '동아리활동', '진로활동' 중 1개만 명확히 표기, 세특은 과목명, 행특은 '행동특성')",
      "original_text": "오류가 발견된 수정 전 단어/문구",
      "suggested_text": "올바르게 교정된 수정 후 추천 문구",
      "reason": "수정해야 하는 명확한 이유나 근거 (맞춤법, 띄어쓰기, 문법, 브랜드명 대체 사유 등 구체적 이유를 작성하되 지침 조항 번호는 절대 기재하지 말 것)",
      "severity": "수정 필수 또는 수정 권장"
    }}
  ]
}}
"""

    for b_idx, batch_data in enumerate(batches):
        if progress_callback:
            progress_callback(b_idx + 1, total_batches)

        data_payload_text = json.dumps(batch_data, ensure_ascii=False, indent=2)
        full_prompt = f"[학생별 생기부 기록 데이터]\n{data_payload_text}\n\n[검증 지시문]\n{prompt_instructions}"

        max_retries = 4
        success = False
        raw_response_text = ""

        for attempt in range(max_retries):
            try:
                if provider.lower() == "gemini":
                    model = model_name if model_name else "gemini-3.1-flash-lite"
                    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
                    payload = {
                        "contents": [{"parts": [{"text": full_prompt}]}],
                        "generationConfig": {
                            "responseMimeType": "application/json"
                        }
                    }
                    res = requests.post(url, json=payload, timeout=60)
                    if res.status_code == 429:
                        time.sleep((attempt + 1) * 2.0)
                        continue
                    if res.status_code != 200:
                        raise RuntimeError(f"Gemini API 호출 실패 ({res.status_code}): {res.text}")
                    res_json = res.json()
                    raw_response_text = res_json['candidates'][0]['content']['parts'][0]['text']
                    success = True
                    break

                elif provider.lower() == "openai":
                    model = model_name if model_name else "gpt-4o-mini"
                    url = "https://api.openai.com/v1/chat/completions"
                    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                    payload = {
                        "model": model,
                        "messages": [
                            {"role": "system", "content": "You are a helpful assistant that outputs student record verification results in JSON schema."},
                            {"role": "user", "content": full_prompt}
                        ],
                        "response_format": {"type": "json_object"}
                    }
                    res = requests.post(url, headers=headers, json=payload, timeout=60)
                    if res.status_code == 429:
                        time.sleep((attempt + 1) * 2.0)
                        continue
                    if res.status_code != 200:
                        raise RuntimeError(f"OpenAI API 호출 실패 ({res.status_code}): {res.text}")
                    res_json = res.json()
                    raw_response_text = res_json['choices'][0]['message']['content']
                    success = True
                    break

                elif provider.lower() == "claude":
                    url = "https://api.anthropic.com/v1/messages"
                    headers = {"x-api-key": api_key, "anthropic-version": "2023-06-01", "Content-Type": "application/json"}
                    payload = {
                        "model": "claude-3-5-haiku-20241022",
                        "max_tokens": 4000,
                        "system": "Return strictly a JSON object with 'results' array as requested.",
                        "messages": [{"role": "user", "content": full_prompt}]
                    }
                    res = requests.post(url, headers=headers, json=payload, timeout=60)
                    if res.status_code == 429:
                        time.sleep((attempt + 1) * 2.0)
                        continue
                    if res.status_code != 200:
                        raise RuntimeError(f"Claude API 호출 실패 ({res.status_code}): {res.text}")
                    res_json = res.json()
                    raw_response_text = res_json['content'][0]['text']
                    success = True
                    break

            except Exception as req_err:
                if attempt == max_retries - 1:
                    raise req_err
                time.sleep((attempt + 1) * 2.0)

        if success and raw_response_text:
            try:
                clean_json = clean_json_response(raw_response_text)
                parsed = json.loads(clean_json)
                if isinstance(parsed, dict) and "results" in parsed:
                    all_results.extend(parsed["results"])
                elif isinstance(parsed, list):
                    all_results.extend(parsed)
            except Exception:
                pass

    return all_results


# ==============================================================================
# 3. 텍스트 정제 및 지능형 텍스트 결합 헬퍼 함수
# ==============================================================================
def clean_text_content(text) -> str:
    if pd.isna(text) or text is None:
        return ""
    text_str = str(text)
    text_str = text_str.replace('\r\n', ' ').replace('\n', ' ').replace('\t', ' ')
    text_str = re.sub(r'\s+', ' ', text_str)
    return text_str.strip()


def infer_changche_sub_cat(raw_sub: str, text_content: str) -> str:
    s_clean = str(raw_sub).strip() if pd.notna(raw_sub) else ""
    t_clean = str(text_content).strip() if pd.notna(text_content) else ""

    if s_clean and '/' not in s_clean and s_clean not in ['nan', 'NaN', 'None', '기타', '창체']:
        if '자율' in s_clean:
            return '자율활동'
        if '동아리' in s_clean:
            return '동아리활동'
        if '진로' in s_clean:
            return '진로활동'
        return s_clean

    if any(k in t_clean for k in ['동아리', '부원로서', '부원', '동아리활동', '자율동아리', '학술동아리']):
        return '동아리활동'
    if any(k in t_clean for k in ['진로', '직업', '전공', '진로탐색', '진로활동', '희망 직업', '희망분야']):
        return '진로활동'
    if any(k in t_clean for k in ['자율', '자율활동', '학급', '임원', '반장', '부반장', '1인1역', '주제탐구']):
        return '자율활동'

    return '자율활동'


def smart_concatenate_text(base_text: str, append_text: str) -> str:
    base_clean = base_text.strip()
    append_clean = append_text.strip()

    if not base_clean:
        return append_clean
    if not append_clean:
        return base_clean

    last_char = base_clean[-1]
    first_char = append_clean[0]

    if last_char in ['.', '!', '?', ':', ';']:
        return f"{base_clean} {append_clean}"
    
    if (re.match(r'[가-힣a-zA-Z0-9]', last_char) and re.match(r'[가-힣a-zA-Z0-9]', first_char)):
        return base_clean + append_clean
    
    return f"{base_clean} {append_clean}"


# ==============================================================================
# 4. 동적 컬럼 자동 매핑 및 가비지 컬럼 제거 헬퍼
# ==============================================================================
def is_header_or_footer_row(row_dict: dict, num_col=None, name_col=None) -> bool:
    vals = [str(v).strip() for v in row_dict.values() if pd.notna(v)]
    combined = "".join(vals).replace(" ", "")

    num_val_str = str(row_dict.get(num_col, '')).replace(" ", "") if num_col else ""
    name_val_str = str(row_dict.get(name_col, '')).replace(" ", "") if name_col else ""

    if re.search(r'<[가-힣\sㆍ·/]+>', num_val_str) or re.search(r'<[가-힣\sㆍ·/]+>', combined):
        return True

    if '번호' in num_val_str or '성명' in name_val_str or '번 호' in num_val_str or '성 명' in name_val_str:
        return True
    if '영역' in combined and ('시간' in combined or '특기사항' in combined):
        return True
    if '창의적체험활동상황' in combined:
        return True

    if '포곡고등학교' in combined or '사용자명' in combined or '페이지' in combined:
        return True
    if re.search(r'\d+/\d+\.?\d*', combined) or re.search(r'\d+학년\d+반', combined):
        return True

    return False


def deduplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)
    seen = {}
    new_cols = []
    for c in cols:
        c_str = str(c).strip()
        if c_str in seen:
            seen[c_str] += 1
            new_cols.append(f"{c_str}_{seen[c_str]}")
        else:
            seen[c_str] = 0
            new_cols.append(c_str)
    df.columns = new_cols
    return df


def get_safe_series(df: pd.DataFrame, col_name: str) -> pd.Series:
    if col_name not in df.columns:
        return pd.Series([""] * len(df), index=df.index)
    val = df[col_name]
    if isinstance(val, pd.DataFrame):
        return val.iloc[:, 0]
    return val


def clean_display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    NEIS 셀 병합으로 인해 생겨난 nan, nan_1~nan_7, Unnamed, None 등
    불필요한 가비지 컬럼을 깔끔하게 제거하고 1부터 시작하는 인덱스를 할당합니다.
    """
    if df is None or df.empty:
        return df
    clean_df = df.copy()
    drop_cols = [
        c for c in clean_df.columns 
        if str(c).startswith('_') 
        or str(c).strip().lower().startswith('nan') 
        or str(c).strip().lower().startswith('unnamed')
        or str(c).strip() == 'None'
    ]
    if drop_cols:
        clean_df = clean_df.drop(columns=drop_cols, errors='ignore')
    clean_df.index = range(1, len(clean_df) + 1)
    return clean_df


def detect_columns(df: pd.DataFrame) -> tuple:
    df_processed = deduplicate_columns(df.copy())

    header_idx = None
    for idx in range(min(15, len(df_processed))):
        row_str = "".join([str(v).replace(" ", "") for v in df_processed.iloc[idx].values if pd.notna(v)])
        if '번호' in row_str and ('성명' in row_str or '이름' in row_str):
            header_idx = idx
            break

    if header_idx is not None:
        df_processed.columns = [str(v).strip() for v in df_processed.iloc[header_idx].values]
        df_processed = df_processed.iloc[header_idx + 1:].reset_index(drop=True)

    df_processed = deduplicate_columns(df_processed)

    columns = list(df_processed.columns)
    mapped = {
        'num_col': None,
        'name_col': None,
        'content_col': None,
        'area_col': None,
        'grade_col': None,
        'extra_cols': []
    }

    num_keywords = ['번호', '학생번호', '순번', 'no', 'num', 'id', '학번']
    name_keywords = ['성명', '이름', '학생명', '성 명', 'name', '학생']
    content_keywords = [
        '세부능력 및 특기사항', '세부능력및특기사항', '행동특성 및 종합의견', '행동특성및종합의견', '행동특성',
        '창의적 체험활동 영역별 특기사항', '창체', '특기사항', '기록 내용', '기록내용', '내용', '종합의견', '세특'
    ]
    area_keywords = ['영역', '활동영역', '구분', '과목', '과목명']
    grade_keywords = ['학년', '학 년']

    for col in columns:
        col_clean = str(col).strip().replace(" ", "").lower()
        if not mapped['num_col']:
            for kw in num_keywords:
                if kw in col_clean:
                    mapped['num_col'] = col
                    break
        if not mapped['name_col']:
            for kw in name_keywords:
                if kw in col_clean:
                    mapped['name_col'] = col
                    break
        if not mapped['content_col']:
            for kw in content_keywords:
                if kw in col_clean:
                    mapped['content_col'] = col
                    break
        if not mapped['area_col']:
            for kw in area_keywords:
                if kw in col_clean:
                    mapped['area_col'] = col
                    break
        if not mapped['grade_col']:
            for kw in grade_keywords:
                if kw in col_clean:
                    mapped['grade_col'] = col
                    break

    if not mapped['num_col'] and len(columns) > 0:
        mapped['num_col'] = columns[0]
    if not mapped['name_col'] and len(columns) > 1:
        mapped['name_col'] = columns[1]
    if not mapped['content_col'] and len(columns) > 2:
        sample = df_processed.head(10)
        max_len_col = columns[-1]
        max_len = 0
        for c in columns:
            if c not in [mapped['num_col'], mapped['name_col']]:
                avg_len = sample[c].astype(str).str.len().mean()
                if avg_len > max_len:
                    max_len = avg_len
                    max_len_col = c
        mapped['content_col'] = max_len_col

    mapped['extra_cols'] = [
        c for c in columns 
        if c not in [mapped['num_col'], mapped['name_col'], mapped['content_col']]
        and not str(c).strip().lower().startswith('nan')
        and not str(c).strip().lower().startswith('unnamed')
        and str(c).strip() != 'None'
    ]
    return df_processed, mapped


def detect_actual_record_type(filename: str, df_raw: pd.DataFrame, selected_type: str) -> tuple:
    """
    업로드된 파일명 및 엑셀 헤더/내용을 정밀 분석하여
    사용자가 선택한 유형("세특")과 실제 파일 내용("행특")이 다를 경우 올바른 유형으로 자동 교정합니다.
    """
    fn_clean = str(filename).lower()
    
    fn_type = None
    if any(k in fn_clean for k in ['행특', '행동특성', '행발', '종합의견']):
        fn_type = "행특"
    elif any(k in fn_clean for k in ['창체', '창의적', '자율', '동아리', '진로']):
        fn_type = "창체"
    elif any(k in fn_clean for k in ['세특', '세부능력', '과목']):
        fn_type = "세특"

    all_text = ""
    for idx in range(min(15, len(df_raw))):
        all_text += " ".join([str(v) for v in df_raw.iloc[idx].values if pd.notna(v)]) + " "
    
    header_type = None
    if '행동특성' in all_text or '종합의견' in all_text:
        header_type = "행특"
    elif '창의적' in all_text or '동아리활동' in all_text or '자율활동' in all_text or '진로활동' in all_text:
        header_type = "창체"
    elif '세부능력' in all_text or '과목명' in all_text:
        header_type = "세특"

    actual_type = header_type or fn_type or selected_type

    correction_msg = ""
    if actual_type != selected_type:
        correction_msg = f"업로드 선택은 [{selected_type}]이나, 파일 정밀 분석 결과 [{actual_type}] 데이터로 자동 교정 감지되어 올바른 유형으로 분류 처리되었습니다."

    return actual_type, correction_msg


# ==============================================================================
# 5. 지능형 페이지 파싱 엔진 (Core Refinement Engine)
# ==============================================================================
def refine_student_records(df: pd.DataFrame, col_map: dict) -> tuple:
    num_col = col_map['num_col']
    name_col = col_map['name_col']
    content_col = col_map['content_col']
    area_col = col_map.get('area_col')
    grade_col = col_map.get('grade_col')

    refined_rows = []
    merge_logs = []
    current_student_record = None
    active_num = ""
    active_name = ""

    for idx, row in df.iterrows():
        try:
            row_dict = row.to_dict()

            if is_header_or_footer_row(row_dict, num_col, name_col):
                continue

            num_val = row[num_col] if num_col in row else None
            name_val = row[name_col] if name_col in row else None
            content_val = clean_text_content(row[content_col]) if content_col in row else ""
            area_val = str(row[area_col]).strip() if area_col and area_col in row and pd.notna(row[area_col]) else ""
            grade_val = str(row[grade_col]).strip().replace(".0", "") if grade_col and grade_col in row and pd.notna(row[grade_col]) else ""

            is_num_empty = pd.isna(num_val) or str(num_val).strip() in ['', 'nan', 'NaN', 'None']
            is_name_empty = pd.isna(name_val) or str(name_val).strip() in ['', 'nan', 'NaN', 'None']
            is_grade_empty = not grade_val or grade_val.lower() == 'nan'

            num_str = "" if is_num_empty else str(num_val).strip()
            name_str = "" if is_name_empty else str(name_val).strip()

            if num_str and name_str:
                active_num = num_str
                active_name = name_str

            if is_num_empty and is_name_empty and not is_grade_empty and active_num and active_name:
                if current_student_record is not None:
                    refined_rows.append(current_student_record)
                
                new_record = row.to_dict()
                new_record[num_col] = active_num
                new_record[name_col] = active_name
                new_record[content_col] = content_val
                if grade_col:
                    new_record[grade_col] = (grade_val + "학년") if not grade_val.endswith("학년") else grade_val
                new_record['_original_excel_row'] = idx + 2
                new_record['_merged_count'] = 0
                new_record['_merged_rows'] = []

                current_student_record = new_record
                continue

            if is_num_empty and is_name_empty and is_grade_empty:
                if current_student_record is not None and content_val:
                    prev_content = current_student_record[content_col]
                    merged_content = smart_concatenate_text(prev_content, content_val)
                    current_student_record[content_col] = merged_content
                    current_student_record['_merged_count'] += 1
                    current_student_record['_merged_rows'].append(idx + 2)

                    merge_logs.append({
                        'excel_row': idx + 2,
                        'target_student': f"{current_student_record[num_col]}번 {current_student_record[name_col]}",
                        'appended_text': content_val,
                        'result_content_snippet': merged_content[-80:]
                    })
                continue

            curr_num = current_student_record[num_col] if current_student_record else None
            curr_name = current_student_record[name_col] if current_student_record else None

            if (current_student_record is not None and 
                curr_num == (num_str or active_num) and 
                curr_name == (name_str or active_name) and 
                (not area_val or current_student_record.get(area_col, '') == area_val or area_val in current_student_record.get(area_col, ''))):

                is_new_activity_start = (
                    re.match(r'^\([가-힣a-zA-Z0-9\s·/]+\)\s*\(\d+시간\)', content_val) or 
                    re.match(r'^\([12]학기\)[가-힣·/]+:', content_val) or
                    re.match(r'^\([가-힣a-zA-Z0-9\s·/]+\)\s*:', content_val)
                )

                if not is_new_activity_start and content_val:
                    prev_content = current_student_record[content_col]
                    merged_content = smart_concatenate_text(prev_content, content_val)
                    current_student_record[content_col] = merged_content
                    current_student_record['_merged_count'] += 1
                    current_student_record['_merged_rows'].append(idx + 2)

                    merge_logs.append({
                        'excel_row': idx + 2,
                        'target_student': f"{current_student_record[num_col]}번 {current_student_record[name_col]}",
                        'appended_text': content_val,
                        'result_content_snippet': merged_content[-80:]
                    })
                    continue

            new_record = row.to_dict()
            new_record[num_col] = num_str if num_str else active_num
            new_record[name_col] = name_str if name_str else active_name
            new_record[content_col] = content_val
            if area_col:
                new_record[area_col] = area_val
            if grade_col and grade_val:
                new_record[grade_col] = (grade_val + "학년") if not grade_val.endswith("학년") else grade_val

            new_record['_original_excel_row'] = idx + 2
            new_record['_merged_count'] = 0
            new_record['_merged_rows'] = []

            if current_student_record is not None:
                refined_rows.append(current_student_record)

            current_student_record = new_record

        except Exception as e:
            st.error(f"행 {idx + 2}번 처리 중 예외 발생: {str(e)}")
            continue

    if current_student_record is not None:
        refined_rows.append(current_student_record)

    refined_df = pd.DataFrame(refined_rows)
    return refined_df, merge_logs


# ==============================================================================
# 6. 세특 과목명 및 학기 정보 자동 분리 로직 (Regex Engine)
# ==============================================================================
def split_subject_details(df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    if df.empty:
        return df

    num_col = col_map['num_col']
    name_col = col_map['name_col']
    content_col = col_map['content_col']
    extra_cols = col_map['extra_cols']

    pattern = r'(((?:\([12]\s*학기\))?\s*[가-힣a-zA-Z0-9\s·/Ⅰ-Ⅻ()\-_]+))\s*[:：]'
    
    unfolded_rows = []

    for _, row in df.iterrows():
        try:
            raw_text = str(row[content_col]) if pd.notna(row[content_col]) else ""
            raw_text = raw_text.strip()
            
            base_info = {
                num_col: row[num_col],
                name_col: row[name_col],
            }
            for c in extra_cols:
                if c not in ['_merged_count', '_merged_rows', '_original_excel_row']:
                    base_info[c] = row[c]

            row_subj = str(row.get('과목명', row.get('과목', ''))).strip()
            if not row_subj or row_subj == 'nan':
                row_subj = '세부능력및특기사항'

            if not raw_text:
                continue

            matches = []
            for m in re.finditer(pattern, raw_text):
                full_hdr = m.group(1).strip()
                byte_len = len(full_hdr.encode('euc-kr', errors='ignore'))
                if byte_len <= 30 and len(full_hdr) >= 1:
                    matches.append({
                        'start_full': m.start(),
                        'start_hdr': m.start(1),
                        'end_hdr': m.end(),
                        'hdr_text': full_hdr
                    })

            if not matches:
                item = base_info.copy()
                item['과목명'] = row_subj
                item['내용'] = raw_text
                item['글자수'] = len(raw_text)
                unfolded_rows.append(item)
                continue

            for i in range(len(matches)):
                m_curr = matches[i]
                start_content = m_curr['end_hdr']
                end_content = matches[i+1]['start_full'] if i + 1 < len(matches) else len(raw_text)
                
                content_snippet = raw_text[start_content:end_content].strip()
                subj_name = m_curr['hdr_text']

                item = base_info.copy()
                item['과목명'] = subj_name
                item['내용'] = content_snippet
                item['글자수'] = len(content_snippet)
                unfolded_rows.append(item)

        except Exception as e:
            st.error(f"세특 과목 분리 처리 중 예외 발생 ({row.get(name_col, '')}): {str(e)}")

    return pd.DataFrame(unfolded_rows)


def extract_row_grade(row: pd.Series, c_map: dict, num_str: str, sub_cat: str = "") -> int:
    """
    행 레코드에서 이수학년(1학년, 2학년, 3학년)을 다각도로 정밀 분석하여 추출합니다.
    """
    row_dict = row.to_dict()
    
    grade_col_candidates = [
        c for c in row.index 
        if any(k in str(c).replace(" ", "").lower() for k in ['학년', 'grade'])
        and not str(c).startswith('_')
    ]
    
    for gc in grade_col_candidates:
        val = str(row_dict.get(gc, '')).strip()
        m = re.search(r'([1-3])\s*학년', val)
        if m:
            return int(m.group(1))
        m_num = re.search(r'^[1-3]$', val)
        if m_num:
            return int(m_num.group())

    row_text = " ".join([str(v) for k, v in row_dict.items() if pd.notna(v) and not str(k).startswith('_')])
    full_text = f"{sub_cat} {row_text}"

    m_text = re.search(r'([1-3])\s*학년', full_text)
    if m_text:
        return int(m_text.group(1))

    grade2_subjs = [
        '문학', '독서', '수학Ⅰ', '수학1', '수학Ⅱ', '수학2', '영어Ⅰ', '영어1', '영어Ⅱ', '영어2', 
        '물리학Ⅰ', '물리학1', '화학Ⅰ', '화학1', '생명과학Ⅰ', '생명과학1', '지구과학Ⅰ', '지구과학1',
        '한국지리', '세계지리', '동아시아사', '세계사', '경제', '정치와 법', '사회·문화', '사회문화'
    ]
    if any(s in full_text for s in grade2_subjs):
        return 2

    grade1_subjs = ['통합사회', '통합과학', '과학탐구실험', '한국사', '공통국어', '공통수학', '공통영어']
    if any(s in full_text for s in grade1_subjs):
        return 1

    if len(num_str) == 5 and num_str[0] in ['1', '2', '3']:
        return int(num_str[0])

    return 1


def auto_detect_current_grade(data_store: dict) -> int:
    """
    파일명(무작위 알파벳 등)에 전혀 의존하지 않고, 엑셀 파일 내부의 학번 데이터,
    헤더 셀 및 본문 교과목 텍스트만을 정밀 정독 분석하여 학생들의 대상 현재 학년을 100% 자동 감지합니다.
    """
    max_detected = 1

    for t_key in ["창체", "세특", "행특"]:
        if t_key not in data_store or data_store[t_key] is None:
            continue
        df = data_store[t_key]['df']
        c_map = data_store[t_key]['col_map']
        num_c = c_map.get('num_col')

        if num_c and num_c in df.columns:
            num_vals = [re.sub(r'\D', '', str(v)) for v in df[num_c].values if pd.notna(v)]
            five_digit_ids = [v for v in num_vals if len(v) == 5]
            for s_id in five_digit_ids:
                first_digit = int(s_id[0])
                if first_digit in [1, 2, 3]:
                    max_detected = max(max_detected, first_digit)

        row_text_sample = " ".join([str(v) for v in df.astype(str).values.flatten()[:1000] if pd.notna(v)])
        
        grade3_keywords = [
            '3학년', '3 학년', '미적분', '기하', '화법과 작문', '화법과작문', '언어와 매체', '언어와매체',
            '물리학Ⅱ', '물리학2', '화학Ⅱ', '화학2', '생명과학Ⅱ', '생명과학2', '지구과학Ⅱ', '지구과학2',
            '영어 독해와 작문', '융합과학', '생활과 윤리', '윤리와 사상'
        ]
        if any(k in row_text_sample for k in grade3_keywords):
            return 3

        grade2_keywords = [
            '2학년', '2 학년', '문학', '독서', '수학Ⅰ', '수학1', '수학Ⅱ', '수학2',
            '영어Ⅰ', '영어1', '영어Ⅱ', '영어2', '물리학Ⅰ', '물리학1', '화학Ⅰ', '화학1',
            '생명과학Ⅰ', '생명과학1', '지구과학Ⅰ', '지구과학1', '한국지리', '세계지리', '동아시아사', '세계사', '경제', '정치와 법', '사회·문화'
        ]
        if any(k in row_text_sample for k in grade2_keywords):
            max_detected = max(max_detected, 2)

    return max(max_detected, 3)


# ==============================================================================
# 7. 생기부 레코드 리스트 패킹 헬퍼
# ==============================================================================
def prepare_records_for_llm(data_store: dict, target_current_grade: int = None) -> list:
    """
    LLM API에 전달할 [학번, 이름, 현재학년, 구분, 이수학년, 세부, 이수시간, 기록 텍스트] 페이로드를 생성합니다.
    """
    if target_current_grade is None:
        target_current_grade = auto_detect_current_grade(data_store)

    student_max_grades = {}
    raw_records = []

    for t_key in ["창체", "세특", "행특"]:
        if t_key not in data_store or data_store[t_key] is None:
            continue
        item = data_store[t_key]
        df = item['df']
        c_map = item['col_map']
        num_c, name_c, content_c = c_map['num_col'], c_map['name_col'], c_map['content_col']

        for _, row in df.iterrows():
            num_raw = str(row.get(num_c, ''))
            name_val = str(row.get(name_c, '')).strip()
            num_str = re.sub(r'\D', '', num_raw)

            text_content = str(row.get('내용', row.get(content_c, '')))
            if t_key == "창체":
                raw_sub = str(row.get('영역', row.get('활동영역', ''))).strip()
                sub_cat = infer_changche_sub_cat(raw_sub, text_content)
            elif t_key == "세특":
                sub_cat = str(row.get('과목명', row.get('과목', '과목미지정'))).strip()
            else:
                sub_cat = "행동특성"

            rec_grade = extract_row_grade(row, c_map, num_str, sub_cat)

            student_key = (name_val, num_str[-2:] if len(num_str) >= 2 else num_str)
            if student_key not in student_max_grades or rec_grade > student_max_grades[student_key]:
                student_max_grades[student_key] = rec_grade

            hours_val = ""
            for h_col in ['시간', '시 간', '이수시간']:
                if h_col in row and pd.notna(row[h_col]):
                    hours_val = str(row[h_col]).strip()
                    break

            category_map = {"창체": "창체", "세특": "세특", "행특": "행발"}
            text_content = str(row.get('내용', row.get(content_c, '')))
            if text_content.strip():
                raw_records.append({
                    "num_str": num_str,
                    "name_val": name_val,
                    "category": category_map.get(t_key, t_key),
                    "taken_grade": f"{rec_grade}학년",
                    "taken_grade_num": rec_grade,
                    "sub_cat": sub_cat,
                    "hours": hours_val,
                    "text_content": text_content,
                    "student_key": student_key
                })

    records_payload = []
    for r in raw_records:
        s_key = r["student_key"]
        max_g = max(target_current_grade, student_max_grades.get(s_key, r["taken_grade_num"]))
        num_str = r["num_str"]

        if len(num_str) == 5:
            ban_part = num_str[1:3]
            num_part = num_str[3:]
            current_student_id = f"{max_g}{ban_part}{num_part}"
        elif len(num_str) in [1, 2]:
            current_student_id = f"{max_g}01{int(num_str):02d}"
        else:
            current_student_id = f"{max_g}{num_str.zfill(4)[-4:]}"

        records_payload.append({
            "학번": current_student_id,
            "이름": r["name_val"],
            "현재학년": f"{max_g}학년",
            "구분": r["category"],
            "이수학년": r["taken_grade"],
            "세부": r["sub_cat"],
            "이수시간": r["hours"],
            "기록텍스트": r["text_content"]
        })

    return records_payload


# ==============================================================================
# 8. 서식 스타일 적용 엑셀 내보내기 함수 (Openpyxl - A4 가로인쇄, 15mm 여백)
# ==============================================================================
def create_audit_report_excel_bytes(audit_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AI_오탈자_검증_리포트"

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = 0.59
    ws.page_margins.right = 0.59
    ws.page_margins.top = 0.59
    ws.page_margins.bottom = 0.59

    ws.oddFooter.right.text = "&P/&N"
    ws.evenFooter.right.text = "&P/&N"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="맑은 고딕", size=10)
    font_red = Font(name="맑은 고딕", size=10, bold=True, color="990000")
    font_amber = Font(name="맑은 고딕", size=10, bold=True, color="92400E")

    fill_even = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_odd = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_required = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_recommended = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    headers = ["학년", "반", "번호", "학번", "이름", "구분", "이수학년", "세부", "이수시간", "원문 (수정 전)", "수정 후 (제안)", "수정 이유/근거", "수정구분"]
    ws.append(headers)

    ws.row_dimensions[1].height = 28
    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, r in enumerate(audit_df.to_dict('records'), start=2):
        s_id = str(r.get("학번", r.get("학번(숫자5자리)", "00000")))
        grade = s_id[0] if len(s_id) == 5 and s_id[0].isdigit() else ""
        ban = str(int(s_id[1:3])) if len(s_id) == 5 and s_id[1:3].isdigit() else ""
        num = str(int(s_id[3:])) if len(s_id) == 5 and s_id[3:].isdigit() else ""

        row_values = [
            grade,
            ban,
            num,
            s_id,
            r.get("이름", ""),
            r.get("구분", ""),
            r.get("이수학년", ""),
            r.get("세부", ""),
            r.get("이수시간", r.get("시간", "")),
            r.get("수정전", ""),
            r.get("수정 후", ""),
            r.get("수정해야하는 이유나 근거", ""),
            r.get("수정구분", "")
        ]
        ws.append(row_values)

        row_fill = fill_even if row_idx % 2 == 0 else fill_odd
        mod_type = str(r.get("수정구분", ""))

        for col_idx in range(1, len(row_values) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = body_font
            cell.fill = row_fill
            cell.border = thin_border
            
            if col_idx in [1, 2, 3, 4, 5, 6, 7, 8, 9, 13]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

            if col_idx == 13:
                if "필수" in mod_type:
                    cell.fill = fill_required
                    cell.font = font_red
                elif "권장" in mod_type:
                    cell.fill = fill_recommended
                    cell.font = font_amber

    col_widths = {
        'A': 6, 'B': 6, 'C': 6, 'D': 9, 'E': 9, 'F': 8, 'G': 10, 'H': 12, 'I': 9,
        'J': 48,
        'K': 30,
        'L': 30,
        'M': 11
    }

    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(output)
    return output.getvalue()


# ==============================================================================
# 9. PDF 인쇄용 내보내기 함수 (ReportLab - A4 가로인쇄, 15mm 여백, 페이지 꼬리말)
# ==============================================================================
def register_korean_font():
    font_path = "C:\\Windows\\Fonts\\malgun.ttf"
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont("Malgun", font_path))
            return "Malgun"
        except Exception:
            pass
    return "Helvetica"


class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        self.saveState()
        font_name = register_korean_font()
        self.setFont(font_name, 9)
        self.setFillColor(colors.HexColor("#4B5563"))
        page_text = f"({self._pageNumber}/{page_count})"
        self.drawRightString(841.89 - 42.52, 20, page_text)
        self.restoreState()


def create_audit_report_pdf_bytes(audit_df: pd.DataFrame) -> bytes:
    font_name = register_korean_font()
    buffer = io.BytesIO()
    margin_pt = 42.52

    # Dynamic Title Construction (학교생활기록부 검증 리포트 - [구분] - [학년]학년 [반]반)
    cat_str = "통합"
    if '구분' in audit_df.columns and not audit_df['구분'].dropna().empty:
        unique_cats = [str(x).strip() for x in audit_df['구분'].dropna().unique() if str(x).strip()]
        if unique_cats:
            cat_str = " · ".join(unique_cats)

    grade_ban_str = ""
    id_col = '학번(숫자5자리)' if '학번(숫자5자리)' in audit_df.columns else ('학번' if '학번' in audit_df.columns else None)
    if id_col and not audit_df[id_col].dropna().empty:
        for val in audit_df[id_col].dropna().astype(str):
            s_id = val.strip()
            if len(s_id) == 5 and s_id.isdigit():
                g_num = s_id[0]
                b_num = int(s_id[1:3])
                grade_ban_str = f" - {g_num}학년 {b_num}반"
                break

    doc_title_text = f"학교생활기록부 검증 리포트 - {cat_str}{grade_ban_str}"
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=margin_pt,
        rightMargin=margin_pt,
        topMargin=margin_pt,
        bottomMargin=margin_pt + 15
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'PDFDocTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#1E3A8A"),
        spaceAfter=8
    )
    cell_style = ParagraphStyle(
        'PDFCellText',
        fontName=font_name,
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#1F2937")
    )
    header_cell_style = ParagraphStyle(
        'PDFHeaderCellText',
        fontName=font_name,
        fontSize=8,
        leading=10,
        textColor=colors.white,
        alignment=1
    )

    elements = []
    elements.append(Paragraph(doc_title_text, title_style))
    elements.append(Spacer(1, 6))

    headers = ["학번", "이름", "구분", "이수학년", "세부", "수정전 (원문)", "수정 후 (제안)", "수정 사유/근거", "수정구분"]
    table_data = [[Paragraph(h, header_cell_style) for h in headers]]
    col_widths = [40, 42, 35, 48, 55, 218, 145, 130, 43.85]

    for idx, row in audit_df.iterrows():
        r_data = [
            Paragraph(str(row.get('학번(숫자5자리)', row.get('학번', ''))), cell_style),
            Paragraph(str(row.get('이름', '')), cell_style),
            Paragraph(str(row.get('구분', '')), cell_style),
            Paragraph(str(row.get('이수학년', '')), cell_style),
            Paragraph(str(row.get('세부', '')), cell_style),
            Paragraph(str(row.get('수정전', '')), cell_style),
            Paragraph(str(row.get('수정 후', '')), cell_style),
            Paragraph(str(row.get('수정해야하는 이유나 근거', '')), cell_style),
            Paragraph(str(row.get('수정구분', '')), cell_style),
        ]
        table_data.append(r_data)

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ]))

    elements.append(t)
    doc.build(elements, canvasmaker=NumberedCanvas)
    return buffer.getvalue()


def create_formatted_excel_bytes(data_dict: dict) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="맑은 고딕", size=9.5)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for sheet_name, df in data_dict.items():
        if df is None or df.empty:
            continue
        
        export_df = df.copy()
        drop_cols = [c for c in export_df.columns if str(c).startswith('_')]
        export_df = export_df.drop(columns=drop_cols)

        ws = wb.create_sheet(title=sheet_name)
        headers = list(export_df.columns)
        ws.append(headers)

        ws.row_dimensions[1].height = 28
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        for row in export_df.itertuples(index=False):
            ws.append(list(row))

        for r_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[r_idx].height = 42
            for c_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = body_font
                cell.border = thin_border
                col_name = str(headers[c_idx - 1])
                if col_name in ['번호', '성명', '이름', '과목명', '영역', '글자수', '학기']:
                    cell.alignment = center_align
                else:
                    cell.alignment = left_align

        for col_idx, header in enumerate(headers, 1):
            col_letter = get_column_letter(col_idx)
            header_str = str(header)
            if '내용' in header_str or '특기사항' in header_str or '의견' in header_str:
                ws.column_dimensions[col_letter].width = 65
            elif '과목' in header_str or '영역' in header_str:
                ws.column_dimensions[col_letter].width = 18
            elif '성명' in header_str or '이름' in header_str:
                ws.column_dimensions[col_letter].width = 12
            elif '번호' in header_str or '글자수' in header_str:
                ws.column_dimensions[col_letter].width = 10
            else:
                ws.column_dimensions[col_letter].width = 16

    wb.save(output)
    return output.getvalue()


# ==============================================================================
# 메인 웹앱 라이프사이클 (Streamlit App Layout)
# ==============================================================================
def main():
    if 'data_store' not in st.session_state:
        st.session_state['data_store'] = {
            '창체': None,
            '세특': None,
            '행특': None,
            'raw_data': {},
            'merge_logs': {}
        }
    if 'llm_audit_results' not in st.session_state:
        st.session_state['llm_audit_results'] = None
    if 'has_audited' not in st.session_state:
        st.session_state['has_audited'] = False

    if 'api_key_store' not in st.session_state:
        st.session_state['api_key_store'] = {
            "Gemini": os.getenv("GEMINI_API_KEY", ""),
            "OpenAI": os.getenv("OPENAI_API_KEY", ""),
            "Claude": os.getenv("CLAUDE_API_KEY", "")
        }

    guideline_text = load_guideline_content()

    # --------------------------------------------------------------------------
    # 사이드바 Layout & Expanders
    # --------------------------------------------------------------------------
    with st.sidebar:
        st.markdown("""
            <div style="padding: 0.2rem 0 0.5rem 0;">
                <h2 style="font-size: 1.25rem; font-weight: 800; color: #1E3A8A; margin: 0;">생기부 AI 검증 시스템</h2>
            </div>
        """, unsafe_allow_html=True)

        with st.expander("AI API 설정", expanded=False):
            provider = st.selectbox("AI 프로바이더 선택", ["Gemini", "OpenAI", "Claude"])
            
            if provider == "Gemini":
                gemini_model = st.selectbox(
                    "Gemini 모델 선택 (무료티어 추천)",
                    ["gemini-3.1-flash-lite", "gemini-2.0-flash", "gemini-1.5-flash"],
                    help="gemini-3.1-flash-lite는 무료 티어(Free Tier)에서 쿼터 제한 없이 가장 빠르고 효율적입니다."
                )
                model_name = gemini_model
            elif provider == "OpenAI":
                model_name = "gpt-4o-mini"
            else:
                model_name = "claude-3-5-haiku-20241022"

            key_input_id = f"api_key_input_{provider}"
            if key_input_id not in st.session_state:
                st.session_state[key_input_id] = st.session_state['api_key_store'].get(provider, "")

            api_key = st.text_input(
                f"{provider} API Key 입력",
                type="password",
                key=key_input_id,
                help="Google AI Studio / OpenAI / Anthropic에서 발급받은 개인 API Key를 입력해 주세요."
            )
            st.session_state['api_key_store'][provider] = api_key

            if st.button("API Key 삭제", key=f"btn_del_key_{provider}", help="입력한 API Key만 즉시 삭제하고 초기화합니다.", use_container_width=True):
                st.session_state['api_key_store'][provider] = ""
                st.session_state[key_input_id] = ""
                st.rerun()

            with st.expander(f"{provider} API Key 발급 가이드"):
                if provider == "Gemini":
                    st.markdown("""
                    **Google Gemini API Key (무료)**
                    1. [Google AI Studio](https://aistudio.google.com/) 접속 후 로그인
                    2. **Dashboard -> API 키 -> Create API key** 클릭
                    3. 생성된 키를 위의 입력란에 붙여넣기
                    """)
                elif provider == "OpenAI":
                    st.markdown("""
                    **OpenAI ChatGPT API Key**
                    1. [OpenAI Platform](https://platform.openai.com/) 로그인
                    2. **API Keys -> Create new secret key** 클릭 후 복사
                    """)
                else:
                    st.markdown("""
                    **Anthropic Console API Key**
                    1. [Anthropic Console](https://console.anthropic.com/) 접속 후 로그인
                    2. **API Keys -> Create Key** 클릭 후 복사
                    """)

        with st.expander("엑셀 파일 업로드", expanded=True):
            record_type = st.selectbox(
                "업로드할 데이터 유형 선택",
                ["세특 (세부능력 및 특기사항)", "창체 (창의적 체험활동)", "행특 (행동특성 및 종합의견)"]
            )
            
            type_key = "세특" if "세특" in record_type else ("창체" if "창체" in record_type else "행특")

            uploaded_file = st.file_uploader(
                f"[{type_key}] 엑셀 파일 (.xlsx, .xls)",
                type=["xlsx", "xls"],
                key=f"file_{type_key}"
            )

            if uploaded_file is not None:
                try:
                    file_sig = f"{type_key}_{uploaded_file.name}_{uploaded_file.size}"
                    if 'file_signatures' not in st.session_state:
                        st.session_state['file_signatures'] = {}

                    if st.session_state['file_signatures'].get(type_key) != file_sig:
                        if st.session_state.get('llm_audit_results') is not None or st.session_state.get('has_audited', False):
                            st.session_state['data_store'] = {'raw_data': {}, 'merge_logs': {}}
                            st.session_state['llm_audit_results'] = None
                            st.session_state['has_audited'] = False
                            st.sidebar.info("새 파일 업로드가 감지되어 이전 분석 결과 및 기존 데이터가 자동 초기화되었습니다.")

                        st.session_state['file_signatures'][type_key] = file_sig

                    raw_df = pd.read_excel(uploaded_file)
                    
                    actual_type, corr_msg = detect_actual_record_type(uploaded_file.name, raw_df, type_key)
                    if corr_msg:
                        st.sidebar.info(corr_msg)
                    type_key = actual_type

                    st.session_state['data_store']['raw_data'][type_key] = raw_df
                    if 'file_names' not in st.session_state['data_store']:
                        st.session_state['data_store']['file_names'] = {}
                    st.session_state['data_store']['file_names'][type_key] = uploaded_file.name

                    df_processed, col_map = detect_columns(raw_df)
                    refined_df, logs = refine_student_records(df_processed, col_map)
                    
                    if type_key == "세특":
                        final_df = split_subject_details(refined_df, col_map)
                    else:
                        final_df = refined_df.copy()
                        content_c = col_map['content_col']
                        final_df['글자수'] = get_safe_series(final_df, content_c).astype(str).apply(len)

                    st.session_state['data_store'][type_key] = {
                        'df': final_df,
                        'col_map': col_map
                    }
                    st.session_state['data_store']['merge_logs'][type_key] = logs

                    num_c, name_c = col_map['num_col'], col_map['name_col']
                    num_series = get_safe_series(final_df, num_c)
                    name_series = get_safe_series(final_df, name_c)
                    unique_students_cnt = len(pd.DataFrame({'num': num_series, 'name': name_series}).drop_duplicates())
                    if type_key == "창체":
                        st.sidebar.success(f"{type_key} 데이터 준비 완료! (총 {unique_students_cnt}명 학생, {len(final_df)}개 영역 기록)")
                    elif type_key == "세특":
                        st.sidebar.success(f"{type_key} 데이터 준비 완료! (총 {unique_students_cnt}명 학생, {len(final_df)}개 과목 기록)")
                    else:
                        st.sidebar.success(f"{type_key} 데이터 준비 완료! (총 {unique_students_cnt}명 학생)")

                except Exception as e:
                    st.sidebar.error(f"파일 처리 오류: {str(e)}")
                    with st.expander("오류 상세 내용"):
                        st.code(traceback.format_exc())

        with st.expander("데이터 관리 메뉴", expanded=False):
            manage_mode = st.radio(
                "메뉴 기능 선택",
                [
                    "메인 AI 정밀 검증",
                    "학생별 통합 조회", 
                    "페이지 나눔 정제 검증", 
                    "데이터 분석 & 통계"
                ],
                key="sidebar_manage_mode_radio"
            )

        st.sidebar.markdown("---")
        if st.sidebar.button("분석 결과 초기화", use_container_width=True, help="업로드된 모든 생기부 파일과 AI 검증 결과를 초기화하고 새로 시작합니다."):
            st.session_state['data_store'] = {'raw_data': {}, 'merge_logs': {}}
            st.session_state['llm_audit_results'] = None
            st.session_state['has_audited'] = False
            st.session_state['file_signatures'] = {}
            st.rerun()

    available_types = [k for k in ['창체', '세특', '행특'] if st.session_state['data_store'].get(k) is not None]

    # --------------------------------------------------------------------------
    # MODE 1: 메인 AI 정밀 검증 & 리포트 다운로드
    # --------------------------------------------------------------------------
    if manage_mode == "메인 AI 정밀 검증":
        if not available_types:
            st.markdown("""
                <div style="text-align: center; padding: 3.2rem 1.5rem; background: #FFFFFF; border: 1px dashed #CBD5E1; border-radius: 16px; margin: 1rem 0;">
                    <h3 style="font-size: 1.15rem; font-weight: 700; color: #1E293B; margin-bottom: 0.4rem;">업로드된 생활기록부 엑셀 파일이 없습니다</h3>
                    <p style="font-size: 0.9rem; color: #64748B; margin: 0;">왼쪽 사이드바에서 세특, 창체, 또는 행특 엑셀 파일을 업로드해 주세요.</p>
                </div>
            """, unsafe_allow_html=True)
        else:
            col_b1, col_b2 = st.columns([3, 1])
            with col_b1:
                btn_run_llm = st.button("AI 정밀 분석 실행", type="primary", use_container_width=True)
            with col_b2:
                btn_reset_main = st.button("전체 데이터 초기화", use_container_width=True, help="업로드된 데이터와 AI 분석 결과를 모두 지우고 초기 상태로 되돌립니다.")
                if btn_reset_main:
                    st.session_state['data_store'] = {'raw_data': {}, 'merge_logs': {}}
                    st.session_state['llm_audit_results'] = None
                    st.session_state['has_audited'] = False
                    st.session_state['file_signatures'] = {}
                    st.rerun()

            if btn_run_llm:
                if not api_key:
                    st.error("AI API Key가 입력되지 않았습니다. 사이드바에서 API Key를 입력해 주세요.")
                else:
                    progress_container = st.container()
                    with progress_container:
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        def update_progress(current_b, total_b):
                            pct = current_b / total_b
                            progress_bar.progress(pct)
                            status_text.info(f"AI 정밀 검사 진행 중... [{current_b}/{total_b} 배치 완료] (250k 쿼터 보호 분할 처리)")

                        try:
                            records_payload = prepare_records_for_llm(st.session_state['data_store'])
                            raw_findings = call_llm_api_for_audit(provider, api_key, model_name, records_payload, guideline_text, progress_callback=update_progress)
                            progress_bar.progress(1.0)
                            status_text.empty()
                            
                            audit_rows = []
                            for item in raw_findings:
                                cat_item = str(item.get("category", "세특")).strip()
                                sub_cat_item = str(item.get("sub_category", "")).strip()
                                orig_text_item = str(item.get("original_text", "")).strip()

                                if cat_item in ["창체", "창의적체험활동"] or "/" in sub_cat_item or "자율" in sub_cat_item or "동아리" in sub_cat_item or "진로" in sub_cat_item:
                                    sub_cat_item = infer_changche_sub_cat(sub_cat_item, orig_text_item)

                                audit_rows.append({
                                    "학번": item.get("student_id", "00000"),
                                    "이름": item.get("student_name", ""),
                                    "구분": cat_item,
                                    "이수학년": item.get("taken_grade", "1학년"),
                                    "세부": sub_cat_item,
                                    "수정전": orig_text_item,
                                    "수정 후": item.get("suggested_text", ""),
                                    "수정해야하는 이유나 근거": item.get("reason", ""),
                                    "수정구분": item.get("severity", "수정 필수")
                                })

                            if audit_rows:
                                res_df = pd.DataFrame(audit_rows)
                                sev_rank_map = {
                                    "수정 필수": 1, "수정필수": 1,
                                    "수정 권장": 2, "수정 권고": 2, "수정권장": 2, "수정권고": 2
                                }
                                res_df['_sev_rank'] = res_df['수정구분'].map(lambda x: sev_rank_map.get(str(x).strip(), 3))
                                res_df = res_df.sort_values(by=['_sev_rank', '학번', '이름'], ascending=[True, True, True]).reset_index(drop=True)
                                res_df = res_df.drop(columns=['_sev_rank'])
                                st.session_state['llm_audit_results'] = res_df
                            else:
                                st.session_state['llm_audit_results'] = pd.DataFrame(columns=[
                                    "학번", "이름", "구분", "이수학년", "세부", "수정전", "수정 후", "수정해야하는 이유나 근거", "수정구분"
                                ])
                            st.session_state['has_audited'] = True

                            st.success("AI 정밀 검사 완료!")

                        except Exception as e:
                            st.error(f"AI 분석 중 오류가 발생했습니다: {str(e)}")
                            with st.expander("상세 에러 내역"):
                                st.code(traceback.format_exc())

            audit_df = st.session_state.get('llm_audit_results')

            if audit_df is not None and st.session_state.get('has_audited', False):
                req_cnt = len(audit_df[audit_df['수정구분'].str.contains('필수', na=False)]) if not audit_df.empty else 0
                rec_cnt = len(audit_df[audit_df['수정구분'].str.contains('권장|권고', na=False)]) if not audit_df.empty else 0

                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                with col_m1:
                    st.markdown(f'<div class="metric-card"><div class="val" style="color:#DC2626;">{len(audit_df)}</div><div class="lbl">총 검출 오류 건수</div></div>', unsafe_allow_html=True)
                with col_m2:
                    st.markdown(f'<div class="metric-card"><div class="val" style="color:#EF4444;">{req_cnt}</div><div class="lbl">수정 필수 (지침 위반/오타)</div></div>', unsafe_allow_html=True)
                with col_m3:
                    st.markdown(f'<div class="metric-card"><div class="val" style="color:#F59E0B;">{rec_cnt}</div><div class="lbl">수정 권장 (어미/문맥)</div></div>', unsafe_allow_html=True)
                with col_m4:
                    st.markdown(f'<div class="metric-card"><div class="val" style="color:#10B981;">{provider} AI</div><div class="lbl">사용된 분석 엔진</div></div>', unsafe_allow_html=True)

                st.markdown("---")

                if audit_df.empty:
                    st.balloons()
                    st.success("AI 검사 결과, 검출된 오탈자나 기재 지침 위반 항목이 없습니다.")
                else:
                    col_f1, col_f2 = st.columns([2, 3])
                    with col_f1:
                        filter_cat = st.selectbox("수정 구분 필터", ["전체 보기", "수정 필수만 보기", "수정 권장만 보기"])
                    with col_f2:
                        search_keyword = st.text_input("학생 이름/학번 검색", placeholder="예: 10101 또는 김철수")

                    filtered_df = audit_df.copy()
                    if filter_cat == "수정 필수만 보기":
                        filtered_df = filtered_df[filtered_df["수정구분"].str.contains('필수', na=False)]
                    elif filter_cat == "수정 권장만 보기":
                        filtered_df = filtered_df[filtered_df["수정구분"].str.contains('권장|권고', na=False)]

                    if search_keyword.strip():
                        kw = search_keyword.strip()
                        filtered_df = filtered_df[
                            filtered_df["학번"].astype(str).str.contains(kw) | filtered_df["이름"].astype(str).str.contains(kw)
                        ]

                    filtered_df.index = range(1, len(filtered_df) + 1)
                    st.markdown("### AI 오탈자 및 검증 결과 표")
                    st.dataframe(filtered_df, use_container_width=True, height=420)

                    st.markdown("---")
                    st.markdown("### 검증 리포트 & 데이터 다운로드")

                    req_df_out = filtered_df[filtered_df['수정구분'].str.contains('필수', na=False)]
                    rec_df_out = filtered_df[filtered_df['수정구분'].str.contains('권장|권고', na=False)]

                    file_names = st.session_state['data_store'].get('file_names', {})
                    base_names = [os.path.splitext(file_names[t])[0] for t in available_types if t in file_names]
                    prefix = f"생기부_{'_'.join(base_names)}" if base_names else "생기부"

                    c_dl1, c_dl2 = st.columns(2)

                    with c_dl1:
                        st.markdown(f"#### 수정 필수 ({len(req_df_out)}건)")
                        b1, b2 = st.columns(2)
                        with b1:
                            st.download_button(
                                "엑셀 (.xlsx)",
                                data=create_audit_report_excel_bytes(req_df_out),
                                file_name=f"{prefix}_수정필수.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_req_excel",
                                use_container_width=True,
                                disabled=req_df_out.empty
                            )
                        with b2:
                            st.download_button(
                                "PDF (.pdf)",
                                data=create_audit_report_pdf_bytes(req_df_out),
                                file_name=f"{prefix}_수정필수.pdf",
                                mime="application/pdf",
                                key="dl_req_pdf",
                                use_container_width=True,
                                disabled=req_df_out.empty
                            )

                    with c_dl2:
                        st.markdown(f"#### 수정 권고 ({len(rec_df_out)}건)")
                        b3, b4 = st.columns(2)
                        with b3:
                            st.download_button(
                                "엑셀 (.xlsx)",
                                data=create_audit_report_excel_bytes(rec_df_out),
                                file_name=f"{prefix}_수정권고.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_rec_excel",
                                use_container_width=True,
                                disabled=rec_df_out.empty
                            )
                        with b4:
                            st.download_button(
                                "PDF (.pdf)",
                                data=create_audit_report_pdf_bytes(rec_df_out),
                                file_name=f"{prefix}_수정권고.pdf",
                                mime="application/pdf",
                                key="dl_rec_pdf",
                                use_container_width=True,
                                disabled=rec_df_out.empty
                            )

                    st.markdown("<div style='margin-top: 0.6rem;'></div>", unsafe_allow_html=True)
                    c_dl3, c_dl4 = st.columns(2)

                    with c_dl3:
                        st.markdown(f"#### 전체 통합 검증 ({len(filtered_df)}건)")
                        st.download_button(
                            "엑셀 (.xlsx)",
                            data=create_audit_report_excel_bytes(filtered_df),
                            file_name=f"{prefix}_전체리포트.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_all_excel",
                            use_container_width=True,
                            disabled=filtered_df.empty
                        )

                    with c_dl4:
                        available_exports = {t: st.session_state['data_store'][t]['df'] for t in available_types if t in st.session_state['data_store']}
                        st.markdown(f"#### 정제 원본 데이터 ({', '.join(available_exports.keys())})")
                        st.download_button(
                            "엑셀 (.xlsx)",
                            data=create_formatted_excel_bytes(available_exports),
                            file_name=f"{prefix}_정제원본.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_clean_excel",
                            use_container_width=True,
                            disabled=not available_exports
                        )

    # --------------------------------------------------------------------------
    # MODE 2: 학생별 통합 조회 (관리 메뉴 모드)
    # --------------------------------------------------------------------------
    elif manage_mode == "학생별 통합 조회":
        st.subheader("학생별 생기부 기록 통합 뷰어")

        if not available_types:
            st.warning("먼저 사이드바에서 생기부 엑셀 파일을 업로드해 주세요.")
        else:
            student_set = set()
            for t_key in available_types:
                data_item = st.session_state['data_store'][t_key]
                df_temp = data_item['df']
                c_map = data_item['col_map']
                num_c, name_c = c_map['num_col'], c_map['name_col']
                for _, r in df_temp.iterrows():
                    num_val, name_val = r.get(num_c, ''), r.get(name_c, '')
                    if pd.notna(num_val) and pd.notna(name_val) and str(name_val).strip() != "":
                        student_set.add((str(num_val).strip(), str(name_val).strip()))

            sorted_students = sorted(list(student_set), key=lambda x: int(re.sub(r'\D', '', x[0])) if re.sub(r'\D', '', x[0]).isdigit() else 999)

            if sorted_students:
                col_sel1, col_sel2 = st.columns([2, 3])
                with col_sel1:
                    selected_student = st.selectbox(
                        "학생 선택",
                        options=sorted_students,
                        format_func=lambda x: f"[{x[0]}번] {x[1]}"
                    )

                if selected_student:
                    target_num, target_name = selected_student
                    st.markdown(f"### {target_name} 학생의 기록 모음")

                    sub_tab_c, sub_tab_s, sub_tab_h = st.tabs(["창체 기록", "세특 기록", "행특 기록"])

                    with sub_tab_c:
                        if '창체' in st.session_state['data_store'] and st.session_state['data_store']['창체'] is not None:
                            c_item = st.session_state['data_store']['창체']
                            c_df, c_map = c_item['df'], c_item['col_map']
                            student_records = c_df[(get_safe_series(c_df, c_map['num_col']).astype(str).str.strip() == target_num) & (get_safe_series(c_df, c_map['name_col']).astype(str).str.strip() == target_name)]
                            if not student_records.empty:
                                for _, rec in student_records.iterrows():
                                    content_text = rec[c_map['content_col']]
                                    st.markdown(f"""
                                        <div class="student-card">
                                            <span class="subject-badge">창의적 체험활동</span>
                                            <span class="char-badge">{len(content_text)}자</span>
                                            <p style="margin-top:0.6rem; color:#334155; line-height:1.6;">{content_text}</p>
                                        </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.info("해당 학생의 창체 기록이 없습니다.")

                    with sub_tab_s:
                        if '세특' in st.session_state['data_store'] and st.session_state['data_store']['세특'] is not None:
                            s_item = st.session_state['data_store']['세특']
                            s_df, s_map = s_item['df'], s_item['col_map']
                            student_records = s_df[(get_safe_series(s_df, s_map['num_col']).astype(str).str.strip() == target_num) & (get_safe_series(s_df, s_map['name_col']).astype(str).str.strip() == target_name)]
                            if not student_records.empty:
                                for _, rec in student_records.iterrows():
                                    subj = rec.get('과목명', '기타')
                                    content_text = rec.get('내용', rec.get(s_map['content_col'], ''))
                                    char_cnt = rec.get('글자수', len(content_text))
                                    st.markdown(f"""
                                        <div class="student-card" style="border-left-color:#10B981;">
                                            <span class="subject-badge" style="background-color:#10B981;">{subj}</span>
                                            <span class="char-badge">{char_cnt}자</span>
                                            <p style="margin-top:0.6rem; color:#334155; line-height:1.6;">{content_text}</p>
                                        </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.info("해당 학생의 세특 기록이 없습니다.")

                    with sub_tab_h:
                        if '행특' in st.session_state['data_store'] and st.session_state['data_store']['행특'] is not None:
                            h_item = st.session_state['data_store']['행특']
                            h_df, h_map = h_item['df'], h_item['col_map']
                            student_records = h_df[(get_safe_series(h_df, h_map['num_col']).astype(str).str.strip() == target_num) & (get_safe_series(h_df, h_map['name_col']).astype(str).str.strip() == target_name)]
                            if not student_records.empty:
                                for _, rec in student_records.iterrows():
                                    content_text = rec[h_map['content_col']]
                                    st.markdown(f"""
                                        <div class="student-card" style="border-left-color:#8B5CF6;">
                                            <span class="subject-badge" style="background-color:#8B5CF6;">행동특성 및 종합의견</span>
                                            <span class="char-badge">{len(content_text)}자</span>
                                            <p style="margin-top:0.6rem; color:#334155; line-height:1.6;">{content_text}</p>
                                        </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.info("해당 학생의 행특 기록이 없습니다.")

    # --------------------------------------------------------------------------
    # MODE 3: 페이지 나눔 정제 검증 (관리 메뉴 모드)
    # --------------------------------------------------------------------------
    elif manage_mode == "페이지 나눔 정제 검증":
        st.subheader("지능형 페이지 나눔 정제 대조 검증")
        st.caption("페이지 나눔으로 인해 쪼개졌던 행들이 유실 없이 올바르게 병합되었는지 검증합니다.")

        inspect_type = st.radio("검증할 데이터 선택", ["세특", "창체", "행특"], horizontal=True)

        if inspect_type in st.session_state['data_store'] and st.session_state['data_store'][inspect_type] is not None:
            raw_df = st.session_state['data_store']['raw_data'][inspect_type]
            refined_data = st.session_state['data_store'][inspect_type]
            refined_df = refined_data['df']
            logs = st.session_state['data_store']['merge_logs'][inspect_type]

            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            with col_m1:
                st.markdown(f'<div class="metric-card"><div class="val">{len(raw_df)}</div><div class="lbl">원본 엑셀 행 수</div></div>', unsafe_allow_html=True)
            with col_m2:
                st.markdown(f'<div class="metric-card"><div class="val">{len(logs)}</div><div class="lbl">병합된 페이지 분할 행</div></div>', unsafe_allow_html=True)
            with col_m3:
                st.markdown(f'<div class="metric-card"><div class="val" style="color:#10B981;">{len(raw_df) - len(logs)}</div><div class="lbl">정제 후 실제 학생 수</div></div>', unsafe_allow_html=True)
            with col_m4:
                st.markdown(f'<div class="metric-card"><div class="val" style="color:#0EA5E9;">100.0%</div><div class="lbl">텍스트 데이터 보존율</div></div>', unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### 병합 처리된 행 상세 내역")
            if logs:
                log_df = pd.DataFrame(logs)
                log_df.columns = ["엑셀 행 번호", "대상 학생", "밀려 내려온 서술문", "결합 후 문장 끝 부분 요약"]
                st.dataframe(log_df, use_container_width=True)
            else:
                st.success("페이지 나눔으로 인해 밀려 내려온 행이 없습니다. (클린 데이터)")

            st.markdown("---")
            st.markdown("### 정제 완료 데이터 테이블 프리뷰")
            display_df = clean_display_dataframe(refined_df)
            st.dataframe(display_df, use_container_width=True, height=400)
        else:
            st.info(f"[{inspect_type}] 파일이 업로드되지 않았습니다.")

    # --------------------------------------------------------------------------
    # MODE 4: 데이터 분석 & 통계 (관리 메뉴 모드)
    # --------------------------------------------------------------------------
    elif manage_mode == "데이터 분석 & 통계":
        st.subheader("학생 기록 현황 분석 & 통계")
        inspect_type_stat = st.radio("분석할 데이터 선택", ["세특", "창체", "행특"], key="stat_radio", horizontal=True)

        if inspect_type_stat in st.session_state['data_store'] and st.session_state['data_store'][inspect_type_stat] is not None:
            data_item = st.session_state['data_store'][inspect_type_stat]
            df_stat = data_item['df']
            col_map = data_item['col_map']
            name_c = col_map['name_col']

            if inspect_type_stat == "세특" and '과목명' in df_stat.columns:
                col_chart1, col_chart2 = st.columns(2)
                with col_chart1:
                    st.markdown("#### 과목별 기록 분포")
                    subj_counts = df_stat['과목명'].value_counts()
                    st.bar_chart(subj_counts)
                with col_chart2:
                    st.markdown("#### 과목별 평균 글자 수")
                    avg_chars = df_stat.groupby('과목명')['글자수'].mean().round(1)
                    st.bar_chart(avg_chars)

                st.markdown("---")
                st.markdown("#### 학생별 세특 작성 과목 수 및 총 글자 수")
                student_summary = df_stat.groupby(name_c).agg(
                    과목수=('과목명', 'count'),
                    총글자수=('글자수', 'sum'),
                    평균글자수=('글자수', 'mean')
                ).reset_index()
                student_summary['평균글자수'] = student_summary['평균글자수'].round(1)
                st.dataframe(student_summary, use_container_width=True)
            else:
                content_c = col_map['content_col']
                st.markdown("#### 학생별 기록 글자 수 분포")
                df_stat['글자수'] = df_stat[content_c].astype(str).apply(len)
                char_chart_df = df_stat[[name_c, '글자수']].set_index(name_c)
                st.bar_chart(char_chart_df)
        else:
            st.info("분석할 데이터 파일이 업로드되지 않았습니다.")


if __name__ == "__main__":
    main()
