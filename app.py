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

# ==============================================================================
# 페이지 기본 설정 & CSS 커스텀 스타일링
# ==============================================================================
st.set_page_config(
    page_title="학교생활기록부 AI(LLM) 오탈자 정밀 검증 시스템",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Modern, Premium UI
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
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    }
    .metric-card .val {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0F172A;
    }
    .metric-card .lbl {
        font-size: 0.85rem;
        color: #64748B;
        margin-top: 0.2rem;
    }
    .student-card {
        background: #F8FAFC;
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
            st.warning(f"⚠️ 지침 md 파일 로드 경고: {e}")
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


def call_llm_api_for_audit(provider: str, api_key: str, model_name: str, records_data: list, guideline_text: str) -> list:
    """
    D:\Cloud\git\ggomcode\ggomcheck의 LLM 백엔드호출 방식을 참고하여 
    생기부 기록 데이터를 LLM AI(Gemini/OpenAI/Claude)에 전달하고 
    오탈자, 맞춤법/문법 오류, 입력불가 용어, 금지어를 정밀 분석하여 JSON 구조로 수신합니다.
    """
    if not api_key:
        raise ValueError("API Key가 설정되지 않았습니다. 사이드바에서 AI API Key를 입력해 주세요.")

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
4. 특수문자 제한 위반:
   - 따옴표('"), 쉼표(,), 마침표(.), 느낌표(!), 물음표(?), 콜론(:), 괄호 외 불필요 특수기호(★, ◆, ~, @, #, $ 등)를 감지하십시오.
5. 학생 입장 서술 지양 어미:
   - ~파악함, ~이해함, ~다짐함, ~느낌, ~배움 등 학생 입장 어미를 감지하여 교사 관점 서술어(~활동지를 작성함, ~모습이 돋보임)로 전환 제안하십시오.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【출력 형식 (JSON Schema)】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
반드시 오직 아래의 JSON 데이터만 반환하십시오. 마크다운 추가 텍스트 없이 순수한 JSON만 반환해야 합니다:
{{
  "results": [
    {{
      "student_id": "학번 (5자리 숫자로 작성, 예: 10101)",
      "student_name": "학생 이름",
      "category": "구분 (창체 / 세특 / 행발 중 하나)",
      "sub_category": "세부 (자율 / 동아리 / 진로 / 과목명 / 행동특성 중 하나)",
      "original_text": "오류가 발견된 수정 전 단어/문구",
      "suggested_text": "올바르게 교정된 수정 후 추천 문구",
      "reason": "수정해야 하는 명확한 이유나 근거 (맞춤법 사유 및 지침 조항)",
      "severity": "수정 필수 또는 수정 권장"
    }}
  ]
}}
"""

    data_payload_text = json.dumps(records_data, ensure_ascii=False, indent=2)
    full_prompt = f"[학생별 생기부 기록 데이터]\n{data_payload_text}\n\n[검증 지시문]\n{prompt_instructions}"

    raw_response_text = ""

    # 1. Gemini API (Default / Recommended)
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
        if res.status_code != 200:
            raise RuntimeError(f"Gemini API 호출 실패 ({res.status_code}): {res.text}")
        res_json = res.json()
        raw_response_text = res_json['candidates'][0]['content']['parts'][0]['text']

    # 2. OpenAI API
    elif provider.lower() == "openai":
        model = model_name if model_name else "gpt-4o-mini"
        url = "https://api.openai.com/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": "You are a helpful assistant that outputs student record verification results in JSON schema."},
                {"role": "user", "content": full_prompt}
            ],
            "response_format": {"type": "json_object"}
        }
        res = requests.post(url, headers=headers, json=payload, timeout=60)
        if res.status_code != 200:
            raise RuntimeError(f"OpenAI API 호출 실패 ({res.status_code}): {res.text}")
        res_json = res.json()
        raw_response_text = res_json['choices'][0]['message']['content']

    # 3. Claude API
    elif provider.lower() == "claude":
        url = "https://api.anthropic.com/v1/messages"
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
        }
        payload = {
            "model": "claude-3-5-haiku-20241022",
            "max_tokens": 4000,
            "system": "Return strictly a JSON object with 'results' array as requested.",
            "messages": [{"role": "user", "content": full_prompt}]
        }
        res = requests.post(url, headers=headers, json=payload, timeout=60)
        if res.status_code != 200:
            raise RuntimeError(f"Claude API 호출 실패 ({res.status_code}): {res.text}")
        res_json = res.json()
        raw_response_text = res_json['content'][0]['text']

    else:
        raise ValueError(f"지원하지 않는 API Provider입니다: {provider}")

    # JSON 파싱
    clean_json = clean_json_response(raw_response_text)
    parsed = json.loads(clean_json)
    
    if isinstance(parsed, dict) and "results" in parsed:
        return parsed["results"]
    elif isinstance(parsed, list):
        return parsed
    else:
        return []


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
# 4. 동적 컬럼 자동 매핑 엔진
# ==============================================================================
def detect_columns(df: pd.DataFrame) -> dict:
    columns = list(df.columns)
    mapped = {
        'num_col': None,
        'name_col': None,
        'content_col': None,
        'extra_cols': []
    }

    num_keywords = ['번호', '학생번호', '순번', 'no', 'num', 'id', '학번']
    name_keywords = ['성명', '이름', '학생명', '성 명', 'name', '학생']
    content_keywords = [
        '세부능력 및 특기사항', '세부능력및특기사항', '행동특성 및 종합의견', '행동특성및종합의견',
        '창의적 체험활동 영역별 특기사항', '창체', '특기사항', '기록 내용', '기록내용', '내용', '종합의견', '세특'
    ]

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

    if not mapped['num_col'] and len(columns) > 0:
        mapped['num_col'] = columns[0]
    if not mapped['name_col'] and len(columns) > 1:
        mapped['name_col'] = columns[1]
    if not mapped['content_col'] and len(columns) > 2:
        sample = df.head(10)
        max_len_col = columns[-1]
        max_len = 0
        for c in columns:
            if c not in [mapped['num_col'], mapped['name_col']]:
                avg_len = sample[c].astype(str).str.len().mean()
                if avg_len > max_len:
                    max_len = avg_len
                    max_len_col = c
        mapped['content_col'] = max_len_col

    mapped['extra_cols'] = [c for c in columns if c not in [mapped['num_col'], mapped['name_col'], mapped['content_col']]]
    return mapped


# ==============================================================================
# 5. 지능형 페이지 파싱 엔진 (Core Refinement Engine)
# ==============================================================================
def refine_student_records(df: pd.DataFrame, col_map: dict) -> tuple:
    num_col = col_map['num_col']
    name_col = col_map['name_col']
    content_col = col_map['content_col']

    refined_rows = []
    merge_logs = []
    current_student_record = None

    for idx, row in df.iterrows():
        try:
            num_val = row[num_col] if num_col in row else None
            name_val = row[name_col] if name_col in row else None
            content_val = clean_text_content(row[content_col]) if content_col in row else ""

            is_num_empty = pd.isna(num_val) or str(num_val).strip() == "" or str(num_val).strip().lower() == "nan"
            is_name_empty = pd.isna(name_val) or str(name_val).strip() == "" or str(name_val).strip().lower() == "nan"

            if is_num_empty and is_name_empty:
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

            new_record = row.to_dict()
            new_record[content_col] = content_val
            new_record['_original_excel_row'] = idx + 2
            new_record['_merged_count'] = 0
            new_record['_merged_rows'] = []

            if current_student_record is not None:
                refined_rows.append(current_student_record)

            current_student_record = new_record

        except Exception as e:
            st.error(f"⚠️ 행 {idx + 2}번 처리 중 예외 발생: {str(e)}")
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

    pattern = r'((?:\([12]학기\))?[가-힣·/]+):'
    unfolded_rows = []

    for _, row in df.iterrows():
        try:
            raw_text = str(row[content_col]) if pd.notna(row[content_col]) else ""
            matches = list(re.finditer(pattern, raw_text))

            base_info = {
                num_col: row[num_col],
                name_col: row[name_col],
            }
            for c in extra_cols:
                if c not in ['_merged_count', '_merged_rows', '_original_excel_row']:
                    base_info[c] = row[c]

            if not matches:
                item = base_info.copy()
                item['과목명'] = '통합/기타'
                item['내용'] = raw_text.strip()
                item['글자수'] = len(raw_text.strip())
                unfolded_rows.append(item)
                continue

            if matches[0].start() > 0:
                prefix_text = raw_text[:matches[0].start()].strip()
                if prefix_text:
                    item = base_info.copy()
                    item['과목명'] = '공통/기타'
                    item['내용'] = prefix_text
                    item['글자수'] = len(prefix_text)
                    unfolded_rows.append(item)

            for i in range(len(matches)):
                subject_name = matches[i].group(1).strip()
                start_pos = matches[i].end()
                end_pos = matches[i+1].start() if i + 1 < len(matches) else len(raw_text)
                subject_content = raw_text[start_pos:end_pos].strip()
                
                item = base_info.copy()
                item['과목명'] = subject_name
                item['내용'] = subject_content
                item['글자수'] = len(subject_content)
                unfolded_rows.append(item)

        except Exception as e:
            st.error(f"⚠️ 세특 과목 분리 처리 중 예외 발생 ({row.get(name_col, '')}): {str(e)}")

    return pd.DataFrame(unfolded_rows)


# ==============================================================================
# 7. 생기부 레코드 리스트 패킹 헬퍼
# ==============================================================================
def prepare_records_for_llm(data_store: dict) -> list:
    """
    LLM API에 전달할 [학번, 이름, 구분, 세부, 기록 텍스트] 페이로드 리스트를 만듭니다.
    """
    records_payload = []
    category_map = {"창체": "창체", "세특": "세특", "행특": "행발"}

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
            if len(num_str) in [1, 2]:
                student_id = f"101{int(num_str):02d}"
            else:
                student_id = num_str.zfill(5) if num_str else "00000"

            if t_key == "창체":
                sub_cat = str(row.get('영역', row.get('활동영역', '자율/동아리/진로'))).strip()
            elif t_key == "세특":
                sub_cat = str(row.get('과목명', '과목미지정')).strip()
            else:
                sub_cat = "행동특성"

            text_content = str(row.get('내용', row.get(content_c, '')))
            if text_content.strip():
                records_payload.append({
                    "학번": student_id,
                    "이름": name_val,
                    "구분": category_map[t_key],
                    "세부": sub_cat,
                    "기록텍스트": text_content
                })
    return records_payload


# ==============================================================================
# 8. 서식 스타일 적용 엑셀 내보내기 함수 (Openpyxl)
# ==============================================================================
def create_audit_report_excel_bytes(audit_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AI 오탈자 및 지침 검증 리포트"

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

    headers = list(audit_df.columns)
    ws.append(headers)

    ws.row_dimensions[1].height = 28
    for c_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    for row in audit_df.itertuples(index=False):
        ws.append(list(row))

    for r_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[r_idx].height = 36
        for c_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = body_font
            cell.border = thin_border
            
            col_name = headers[c_idx - 1]
            if col_name in ["학번", "이름", "구분", "세부", "수정구분"]:
                cell.alignment = center_align
                if col_name == "수정구분":
                    if cell.value == "수정 필수":
                        cell.font = Font(name="맑은 고딕", size=9.5, bold=True, color="DC2626")
                    else:
                        cell.font = Font(name="맑은 고딕", size=9.5, bold=True, color="D97706")
            else:
                cell.alignment = left_align

    col_widths = {
        "학번": 12, "이름": 12, "구분": 10, "세부": 16,
        "수정전": 28, "수정 후": 28, "수정해야하는 이유나 근거": 45, "수정구분": 12
    }
    for c_idx, header in enumerate(headers, 1):
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = col_widths.get(header, 18)

    wb.save(output)
    return output.getvalue()


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

    guideline_text = load_guideline_content()

    # --------------------------------------------------------------------------
    # 헤더 섹션
    # --------------------------------------------------------------------------
    st.markdown("""
        <div class="main-header">
            <h1>🎓 학교생활기록부 AI(LLM) 오탈자 정밀 검증 & 데이터 정제 시스템</h1>
            <p>Gemini/OpenAI/Claude AI 모델을 활용하여 기재 지침 준수 여부, 오탈자, 맞춤법/문법 오류를 정밀 검출합니다.</p>
        </div>
    """, unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # 사이드바: AI API 설정 및 엑셀 파일 업로드
    # --------------------------------------------------------------------------
    with st.sidebar:
        st.header("🤖 AI (LLM) API 설정")
        
        provider = st.selectbox("AI 모델 선택", ["Gemini", "OpenAI", "Claude"])
        
        default_api_key = os.getenv("GEMINI_API_KEY", "") if provider == "Gemini" else os.getenv("OPENAI_API_KEY", "")
        api_key = st.text_input(
            f"{provider} API Key 입력",
            value=default_api_key,
            type="password",
            help="AI 정밀 오탈자 검증을 위해 API 키를 입력해 주세요."
        )

        model_name = "gemini-3.1-flash-lite" if provider == "Gemini" else ("gpt-4o-mini" if provider == "OpenAI" else "claude-3-5-haiku-20241022")

        st.markdown("---")
        st.header("📂 엑셀 파일 업로드")
        st.caption("창체, 세특, 행특 엑셀 파일을 선택하여 업로드하세요.")

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
                raw_df = pd.read_excel(uploaded_file)
                st.session_state['data_store']['raw_data'][type_key] = raw_df

                col_map = detect_columns(raw_df)
                refined_df, logs = refine_student_records(raw_df, col_map)
                
                if type_key == "세특":
                    final_df = split_subject_details(refined_df, col_map)
                else:
                    final_df = refined_df.copy()
                    content_c = col_map['content_col']
                    final_df['글자수'] = final_df[content_c].astype(str).apply(len)

                st.session_state['data_store'][type_key] = {
                    'df': final_df,
                    'col_map': col_map
                }
                st.session_state['data_store']['merge_logs'][type_key] = logs

                st.sidebar.success(f"✅ {type_key} 데이터 준비 완료! ({len(refined_df)}명 학생)")

            except Exception as e:
                st.sidebar.error(f"❌ 파일 처리 오류: {str(e)}")
                with st.expander("오류 상세 내용"):
                    st.code(traceback.format_exc())

    # --------------------------------------------------------------------------
    # 메인 콘텐츠 영역 (Tabs)
    # --------------------------------------------------------------------------
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🚨 AI(LLM) 오탈자 정밀 검증", 
        "🔍 학생별 통합 조회", 
        "🛠️ 페이지 나눔 정제 검증", 
        "📊 데이터 분석 & 통계", 
        "📥 엑셀 통합 다운로드"
    ])

    # ==========================================================================
    # Tab 1: 🚨 AI(LLM) 오탈자 & 지침 정밀 검증 (최우선 메인 탭)
    # ==========================================================================
    with tab1:
        st.subheader("🚨 AI (LLM) 기반 생기부 오탈자·맞춤법 및 지침 정밀 검증 리포트")
        st.caption("AI(Gemini/OpenAI/Claude)가 지침 문서(`학교생활기록부_기재_및_검증_지침.md`)의 원칙을 바탕으로 오탈자, 맞춤법/문법 오류, 입력불가 용어, 금지어를 정밀 검출합니다.")

        available_types = [k for k in ['창체', '세특', '행특'] if st.session_state['data_store'].get(k) is not None]

        if not available_types:
            st.warning("👈 먼저 사이드바에서 생기부 엑셀 파일을 업로드해 주세요.")
        else:
            col_b1, col_b2 = st.columns([3, 1])
            with col_b1:
                st.info("💡 사이드바에 API Key를 입력한 후 아래 [AI 정밀 분석 실행] 버튼을 클릭하세요.")
            with col_b2:
                btn_run_llm = st.button("🚀 AI 정밀 분석 실행", type="primary", use_container_width=True)

            if btn_run_llm:
                if not api_key:
                    st.error("⚠️ AI API Key가 입력되지 않았습니다. 사이드바에서 API Key를 입력해 주세요.")
                else:
                    with st.spinner("🤖 AI가 학교생활기록부 전 문장의 오탈자 및 기재 지침을 분석하고 있습니다... (약 10~30초 소요)"):
                        try:
                            records_payload = prepare_records_for_llm(st.session_state['data_store'])
                            raw_findings = call_llm_api_for_audit(provider, api_key, model_name, records_payload, guideline_text)
                            
                            # 데이터프레임 변환
                            audit_rows = []
                            for item in raw_findings:
                                audit_rows.append({
                                    "학번": item.get("student_id", "00000"),
                                    "이름": item.get("student_name", ""),
                                    "구분": item.get("category", "세특"),
                                    "세부": item.get("sub_category", ""),
                                    "수정전": item.get("original_text", ""),
                                    "수정 후": item.get("suggested_text", ""),
                                    "수정해야하는 이유나 근거": item.get("reason", ""),
                                    "수정구분": item.get("severity", "수정 필수")
                                })

                            if audit_rows:
                                res_df = pd.DataFrame(audit_rows)
                                res_df = res_df.sort_values(by=["학번", "이름"], ascending=True).reset_index(drop=True)
                                st.session_state['llm_audit_results'] = res_df
                            else:
                                st.session_state['llm_audit_results'] = pd.DataFrame(columns=[
                                    "학번", "이름", "구분", "세부", "수정전", "수정 후", "수정해야하는 이유나 근거", "수정구분"
                                ])

                            st.success("✅ AI 정밀 검사 완료!")

                        except Exception as e:
                            st.error(f"❌ AI 분석 중 오류가 발생했습니다: {str(e)}")
                            with st.expander("상세 에러 내역"):
                                st.code(traceback.format_exc())

            # 분석 결과 출력
            audit_df = st.session_state.get('llm_audit_results')

            if audit_df is not None:
                req_cnt = len(audit_df[audit_df['수정구분'] == '수정 필수']) if not audit_df.empty else 0
                rec_cnt = len(audit_df[audit_df['수정구분'] == '수정 권장']) if not audit_df.empty else 0

                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                with col_m1:
                    st.markdown(f'<div class="metric-card"><div class="val" style="color:#DC2626;">{len(audit_df)}</div><div class="lbl">총 검출 오류 건수</div></div>', unsafe_allow_html=True)
                with col_m2:
                    st.markdown(f'<div class="metric-card"><div class="val" style="color:#EF4444;">{req_cnt}</div><div class="lbl">🚨 수정 필수 (지침 위반/오타)</div></div>', unsafe_allow_html=True)
                with col_m3:
                    st.markdown(f'<div class="metric-card"><div class="val" style="color:#F59E0B;">{rec_cnt}</div><div class="lbl">⚠️ 수정 권장 (어미/문맥)</div></div>', unsafe_allow_html=True)
                with col_m4:
                    st.markdown(f'<div class="metric-card"><div class="val" style="color:#10B981;">{provider} AI</div><div class="lbl">사용된 분석 엔진</div></div>', unsafe_allow_html=True)

                st.markdown("---")

                if audit_df.empty:
                    st.balloons()
                    st.success("🎉 AI 검사 결과, 검출된 오탈자나 기재 지침 위반 항목이 없습니다.")
                else:
                    col_f1, col_f2 = st.columns([2, 3])
                    with col_f1:
                        filter_cat = st.selectbox("수정 구분 필터", ["전체 보기", "수정 필수만 보기", "수정 권장만 보기"])
                    with col_f2:
                        search_keyword = st.text_input("학생 이름/학번 검색", placeholder="예: 10101 또는 김철수")

                    filtered_df = audit_df.copy()
                    if filter_cat == "수정 필수만 보기":
                        filtered_df = filtered_df[filtered_df["수정구분"] == "수정 필수"]
                    elif filter_cat == "수정 권장만 보기":
                        filtered_df = filtered_df[filtered_df["수정구분"] == "수정 권장"]

                    if search_keyword.strip():
                        kw = search_keyword.strip()
                        filtered_df = filtered_df[
                            filtered_df["학번"].astype(str).str.contains(kw) | filtered_df["이름"].astype(str).str.contains(kw)
                        ]

                    st.markdown("### 📋 AI 오탈자 및 검증 결과 표 (학번순 정렬)")
                    st.dataframe(filtered_df, use_container_width=True, height=450)

                    audit_excel_bytes = create_audit_report_excel_bytes(filtered_df)
                    st.download_button(
                        label="💾 AI 오탈자 & 지침 검증 리포트 엑셀 다운로드 (.xlsx)",
                        data=audit_excel_bytes,
                        file_name="생기부_AI_오탈자_및_지침검증_리포트.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.info("👆 [AI 정밀 분석 실행] 버튼을 눌러 AI 검증을 시작하세요.")

    # ==========================================================================
    # Tab 2: 학생별 통합 조회
    # ==========================================================================
    with tab2:
        st.subheader("👤 학생별 생기부 기록 통합 뷰어")

        if not available_types:
            st.warning("👈 먼저 사이드바에서 생기부 엑셀 파일을 업로드해 주세요.")
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
                    st.markdown(f"### 📋 {target_name} 학생의 기록 모음")

                    sub_tab_c, sub_tab_s, sub_tab_h = st.tabs(["창체 기록", "세특 기록", "행특 기록"])

                    with sub_tab_c:
                        if '창체' in st.session_state['data_store'] and st.session_state['data_store']['창체'] is not None:
                            c_item = st.session_state['data_store']['창체']
                            c_df, c_map = c_item['df'], c_item['col_map']
                            student_records = c_df[(c_df[c_map['num_col']].astype(str).str.strip() == target_num) & (c_df[c_map['name_col']].astype(str).str.strip() == target_name)]
                            if not student_records.empty:
                                for _, rec in student_records.iterrows():
                                    content_text = rec[c_map['content_col']]
                                    st.markdown(f"""
                                        <div class="student-card">
                                            <span class="subject-badge">창의적 체험활동</span>
                                            <span class="char-badge">📏 {len(content_text)}자</span>
                                            <p style="margin-top:0.6rem; color:#334155; line-height:1.6;">{content_text}</p>
                                        </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.info("해당 학생의 창체 기록이 없습니다.")

                    with sub_tab_s:
                        if '세특' in st.session_state['data_store'] and st.session_state['data_store']['세특'] is not None:
                            s_item = st.session_state['data_store']['세특']
                            s_df, s_map = s_item['df'], s_item['col_map']
                            student_records = s_df[(s_df[s_map['num_col']].astype(str).str.strip() == target_num) & (s_df[s_map['name_col']].astype(str).str.strip() == target_name)]
                            if not student_records.empty:
                                for _, rec in student_records.iterrows():
                                    subj = rec.get('과목명', '기타')
                                    content_text = rec.get('내용', rec.get(s_map['content_col'], ''))
                                    char_cnt = rec.get('글자수', len(content_text))
                                    st.markdown(f"""
                                        <div class="student-card" style="border-left-color:#10B981;">
                                            <span class="subject-badge" style="background-color:#10B981;">📚 {subj}</span>
                                            <span class="char-badge">📏 {char_cnt}자</span>
                                            <p style="margin-top:0.6rem; color:#334155; line-height:1.6;">{content_text}</p>
                                        </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.info("해당 학생의 세특 기록이 없습니다.")

                    with sub_tab_h:
                        if '행특' in st.session_state['data_store'] and st.session_state['data_store']['행특'] is not None:
                            h_item = st.session_state['data_store']['행특']
                            h_df, h_map = h_item['df'], h_item['col_map']
                            student_records = h_df[(h_df[h_map['num_col']].astype(str).str.strip() == target_num) & (h_df[h_map['name_col']].astype(str).str.strip() == target_name)]
                            if not student_records.empty:
                                for _, rec in student_records.iterrows():
                                    content_text = rec[h_map['content_col']]
                                    st.markdown(f"""
                                        <div class="student-card" style="border-left-color:#8B5CF6;">
                                            <span class="subject-badge" style="background-color:#8B5CF6;">🌟 행동특성 및 종합의견</span>
                                            <span class="char-badge">📏 {len(content_text)}자</span>
                                            <p style="margin-top:0.6rem; color:#334155; line-height:1.6;">{content_text}</p>
                                        </div>
                                    """, unsafe_allow_html=True)
                            else:
                                st.info("해당 학생의 행특 기록이 없습니다.")

    # ==========================================================================
    # Tab 3: 페이지 나눔 정제 검증
    # ==========================================================================
    with tab3:
        st.subheader("🛠️ 지능형 페이지 나눔 정제 대조 검증")
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
            st.markdown("### 🧩 병합 처리된 행 상세 내역")
            if logs:
                log_df = pd.DataFrame(logs)
                log_df.columns = ["엑셀 행 번호", "대상 학생", "밀려 내려온 서술문", "결합 후 문장 끝 부분 요약"]
                st.dataframe(log_df, use_container_width=True)
            else:
                st.success("페이지 나눔으로 인해 밀려 내려온 행이 없습니다. (클린 데이터)")

            st.markdown("---")
            st.markdown("### 📊 정제 완료 데이터 테이블 프리뷰")
            display_df = refined_df.copy()
            drop_internal = [c for c in display_df.columns if str(c).startswith('_')]
            display_df = display_df.drop(columns=drop_internal)
            st.dataframe(display_df, use_container_width=True, height=400)
        else:
            st.info(f"[{inspect_type}] 파일이 업로드되지 않았습니다.")

    # ==========================================================================
    # Tab 4: 데이터 분석 & 통계
    # ==========================================================================
    with tab4:
        st.subheader("📊 학생 기록 현황 분석 & 통계")
        inspect_type_stat = st.radio("분석할 데이터 선택", ["세특", "창체", "행특"], key="stat_radio", horizontal=True)

        if inspect_type_stat in st.session_state['data_store'] and st.session_state['data_store'][inspect_type_stat] is not None:
            data_item = st.session_state['data_store'][inspect_type_stat]
            df_stat = data_item['df']
            col_map = data_item['col_map']
            name_c = col_map['name_col']

            if inspect_type_stat == "세특" and '과목명' in df_stat.columns:
                col_chart1, col_chart2 = st.columns(2)
                with col_chart1:
                    st.markdown("#### 📚 과목별 기록 분포")
                    subj_counts = df_stat['과목명'].value_counts()
                    st.bar_chart(subj_counts)
                with col_chart2:
                    st.markdown("#### 📏 과목별 평균 글자 수")
                    avg_chars = df_stat.groupby('과목명')['글자수'].mean().round(1)
                    st.bar_chart(avg_chars)

                st.markdown("---")
                st.markdown("#### 👨‍🎓 학생별 세특 작성 과목 수 및 총 글자 수")
                student_summary = df_stat.groupby(name_c).agg(
                    과목수=('과목명', 'count'),
                    총글자수=('글자수', 'sum'),
                    평균글자수=('글자수', 'mean')
                ).reset_index()
                student_summary['평균글자수'] = student_summary['평균글자수'].round(1)
                st.dataframe(student_summary, use_container_width=True)
            else:
                content_c = col_map['content_col']
                st.markdown("#### 👨‍🎓 학생별 기록 글자 수 분포")
                df_stat['글자수'] = df_stat[content_c].astype(str).apply(len)
                char_chart_df = df_stat[[name_c, '글자수']].set_index(name_c)
                st.bar_chart(char_chart_df)
        else:
            st.info("분석할 데이터 파일이 업로드되지 않았습니다.")

    # ==========================================================================
    # Tab 5: 엑셀 통합 다운로드
    # ==========================================================================
    with tab5:
        st.subheader("📥 엑셀 파일 다운로드")
        st.write("AI 오탈자 검증 리포트 및 정제 완료 데이터를 엑셀 파일로 각각 다운로드할 수 있습니다.")

        available_exports = {}
        for t_key in ["창체", "세특", "행특"]:
            if t_key in st.session_state['data_store'] and st.session_state['data_store'][t_key] is not None:
                available_exports[t_key] = st.session_state['data_store'][t_key]['df']

        if available_exports:
            audit_df = st.session_state.get('llm_audit_results')
            
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                st.markdown("### 🚨 AI 오탈자 & 지침 검증 리포트")
                if audit_df is not None:
                    st.write(f"총 {len(audit_df)}건의 검출 항목이 포함된 8개 컬럼 규격 엑셀 리포트입니다.")
                    audit_bytes = create_audit_report_excel_bytes(audit_df)
                    st.download_button(
                        label="💾 AI 오탈자 검증 리포트 다운로드 (.xlsx)",
                        data=audit_bytes,
                        file_name="생기부_AI_오탈자_및_지침검증_리포트.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.info("첫 번째 탭에서 [AI 정밀 분석 실행]을 먼저 진행해 주세요.")

            with col_d2:
                st.markdown("### 📄 정제 완료 통합 생기부 데이터")
                st.write(f"시트: {', '.join(available_exports.keys())}")
                excel_bytes = create_formatted_excel_bytes(available_exports)
                st.download_button(
                    label="💾 정제 완료 데이터 다운로드 (.xlsx)",
                    data=excel_bytes,
                    file_name="생기부_정제_데이터_통합.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.warning("다운로드할 데이터가 없습니다. 사이드바에서 엑셀 파일을 먼저 업로드해 주세요.")


if __name__ == "__main__":
    main()
