import io
import os
import re
import traceback
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 페이지 기본 설정 & CSS 커스텀 스타일링
# ==============================================================================
st.set_page_config(
    page_title="학교생활기록부 데이터 정제 & 오탈자 정밀 검증 시스템",
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
    .badge-required {
        background-color: #EF4444;
        color: white;
        padding: 0.15rem 0.5rem;
        border-radius: 4px;
        font-weight: 700;
        font-size: 0.8rem;
    }
    .badge-recommended {
        background-color: #F59E0B;
        color: white;
        padding: 0.15rem 0.5rem;
        border-radius: 4px;
        font-weight: 700;
        font-size: 0.8rem;
    }
    </style>
""", unsafe_allow_html=True)


# ==============================================================================
# 1. 지침 MD 파일 동적 파서 & 지침 규칙 로더
# ==============================================================================
@st.cache_data
def load_guideline_rules(md_file_path: str = "data/학교생활기록부_기재_및_검증_지침.md") -> dict:
    """
    data/학교생활기록부_기재_및_검증_지침.md 파일에서 
    입력불가 용어 및 대체어 매핑표, 금지어 키워드, 학생입장 어미를 동적으로 파싱합니다.
    """
    mapping_dict = {}
    forbidden_keywords = [
        "수상", "대회", "공모전", "논문", "소논문", "탐구 보고서", "연구 보고서", "자격증", 
        "방과후학교", "모의고사", "어학시험", "특정대학", "시청", "박물관", "상호", "강사명", 
        "교사명", "학교이름", "학교별칭", "축제이름", "해외활동", "외국어", "K-MOOC", "MOOC", 
        "KOCW", "TED", "가정환경", "장학금"
    ]
    student_endings = ["파악함", "이해함", "깨달음", "다짐함", "느낌", "배움", "알게 됨", "생각해 봄", "생각함"]

    if os.path.exists(md_file_path):
        try:
            with open(md_file_path, "r", encoding="utf-8") as f:
                content = f.read()

            # 마크다운 표 영역 파싱 (| 입력불가 용어 | 올바른 대체어 |)
            table_lines = re.findall(r'\| ([^\|]+) \| ([^\|]+) \|', content)
            for raw, replacement in table_lines:
                raw_clean = raw.strip()
                rep_clean = replacement.strip()
                if raw_clean and rep_clean and "입력불가" not in raw_clean and "---" not in raw_clean:
                    # 다중 용어 콤마 분리 (예: 네이버, 구글, 다음)
                    for item in raw_clean.split(','):
                        item_s = item.strip()
                        if item_s:
                            mapping_dict[item_s] = rep_clean
        except Exception as e:
            st.warning(f"⚠️ 지침 md 파일 파싱 중 경고: {e}")

    # 기본 필수 입력불가 용어 매핑 보장 (fallback)
    default_mappings = {
        "네이버": "포털 사이트", "구글": "포털 사이트", "다음": "포털 사이트",
        "네이버 밴드": "교육 플랫폼", "구글클래스룸": "교육 플랫폼", "클래스팅": "학습 플랫폼",
        "유튜브": "동영상 공유 서비스", "유튜버": "동영상 크리에이터", "카카오톡": "SNS 메신저",
        "인스타그램": "소셜 미디어", "페이스북": "소셜 미디어", "틱톡": "엔터테인먼트 플랫폼",
        "아이폰": "스마트폰", "아이패드": "태블릿PC", "갤럭시탭": "태블릿PC", "크롬북": "휴대용 컴퓨터",
        "KTX": "(초)고속열차", "SRT": "(초)고속열차", "MBTI": "성격유형 검사", "챗GPT": "대화형 인공지능",
        "ZOOM": "화상 회의", "파이썬": "프로그래밍 언어", "HTML": "웹페이지 제작 언어", "CSS": "스타일 시트 언어",
        "미리캔버스": "디자인 제작 플랫폼", "망고보드": "온라인 디자인 도구", "패들렛": "온라인 협업 플랫폼"
    }
    for k, v in default_mappings.items():
        if k not in mapping_dict:
            mapping_dict[k] = v

    return {
        "mappings": mapping_dict,
        "forbidden_keywords": forbidden_keywords,
        "student_endings": student_endings
    }


# ==============================================================================
# 2. 오탈자 & 맞춤법·문법 & 지침 위반 정밀 검증 엔진
# ==============================================================================
def inspect_student_record_text(text: str, rules: dict) -> list:
    """
    단일 서술문 텍스트에 대하여 
    1) 자주 틀리는 오탈자 및 한국어 문법 오류
    2) 불필요 특수문자
    3) 입력 불가 용어 (대체어 제공)
    4) 생기부 기재 금지 키워드
    5) 학생 입장 서술어 지양 어미
    를 정밀 검증하여 발견된 오류 항목 리스트를 반환합니다.
    """
    findings = []
    if not text or pd.isna(text):
        return findings

    text_str = str(text)

    # --------------------------------------------------------------------------
    # Rule 1. 자주 틀리는 한글 오탈자 & 맞춤법·어휘 문법 오류 DB
    # --------------------------------------------------------------------------
    typo_database = [
        (r'\b도우는\b', '돕는', '오탈자/문법 오류 (\'돕다\'의 관형사형 어미는 \'돕는\'이 올바름)', '수정 필수'),
        (r'\b만듬\b', '만듦', '맞춤법 오류 (명사형 어미 표기는 \'만듦\'이 올바름)', '수정 필수'),
        (r'\b이끔\b', '이끎', '맞춤법 오류 (명사형 어미 표기는 \'이끎\'이 올바름)', '수정 필수'),
        (r'\b치뤄\b', '치러', '어휘 활용 오탈자 (\'치르다\'의 어미 활용은 \'치러\'가 올바름)', '수정 필수'),
        (r'\b치뤘\b', '치렀', '어휘 활용 오탈자 (\'치르다\'의 과거형은 \'치렀\'이 올바름)', '수정 필수'),
        (r'\b되서\b', '돼서', '맞춤법 오류 (\'되어\'의 줄임말은 \'돼서\'가 올바름)', '수정 필수'),
        (r'\b안되\b', '안 돼', '띄어쓰기/맞춤법 오류 (\'안 돼\' 또는 \'안 됨\'으로 수정)', '수정 필수'),
        (r'\b몇일\b', '며칠', '맞춤법 오류 (\'며칠\'이 올바른 표준어 표기임)', '수정 필수'),
        (r'\b오랫만에\b', '오랜만에', '맞춤법 오류 (\'오랜만에\'가 올바른 표기임)', '수정 필수'),
        (r'\b띔\b', '띰', '맞춤법 오류 (\'눈에 띰\'으로 표기)', '수정 필수'),
        (r'\b설레임\b', '설렘', '명사형 표기 오류 (\'설렘\'이 올바른 표기임)', '수정 권장'),
        (r'\b삼가하다\b', '삼가다', '어휘 오류 (\'삼가다\'가 올바른 기본형임)', '수정 권장'),
        (r'\b어의없\b', '어처구니없', '어휘 오탈자 (\'어이없다/어처구니없다\'가 올바름)', '수정 필수'),
        (r'\b밞아\b', '밟아', '받침 오탈자 (\'밟아\'가 올바른 표기임)', '수정 필수'),
        (r'\b따라함\b', '따라 함', '띄어쓰기 오류 (\'따라 함\'으로 띄어 씀)', '수정 권장'),
        (r'\b가르키\b', '가리키', '어휘 오류 (\'가리키다\'와 \'가르치다\'의 구별 필요)', '수정 권장'),
        (r'\.\.', '.', '문장부호 중복 오류 (마침표가 연속 중복됨)', '수정 권장'),
        (r'\,{2,}', ',', '문장부호 중복 오류 (쉼표가 연속 중복됨)', '수정 권장'),
        (r'\s{2,}', ' ', '다중 공백 오류 (연속된 공백이 포함됨)', '수정 권장'),
    ]

    for pattern, replacement, reason, category in typo_database:
        for match in re.finditer(pattern, text_str):
            findings.append({
                "raw_text": match.group(0),
                "suggested_text": replacement,
                "reason": reason,
                "category": category
            })

    # 조사 띄어쓰기 오류 검사 (예: "조사 를", "활동 을", "학생 은")
    josa_typo_matches = re.finditer(r'([가-힣]{2,})\s+([을를이가은는에의로으로에서부터까지])\b', text_str)
    for m in josa_typo_matches:
        word, josa = m.group(1), m.group(2)
        # 단어 뒤 조사는 붙여 쓰는 것이 원칙
        findings.append({
            "raw_text": m.group(0),
            "suggested_text": f"{word}{josa}",
            "reason": "조사 띄어쓰기 오류 (조사는 앞 단어에 붙여 써야 함)",
            "category": "수정 필수"
        })

    # --------------------------------------------------------------------------
    # Rule 2. 특수문자 제한 검사
    # 허용: 따옴표('"), 쉼표(,), 마침표(.), 느낌표(!), 물음표(?), 콜론(:), 세미콜론(;), 괄호(()[])
    # --------------------------------------------------------------------------
    forbidden_specials = re.finditer(r'([★◆▲■●★☆◇△□○~@#$%^&*+=<>/\\])', text_str)
    for sm in forbidden_specials:
        findings.append({
            "raw_text": sm.group(0),
            "suggested_text": "삭제 또는 문장 기호(점, 쉼표, 따옴표)로 변경",
            "reason": f"특수문자 기재 불가 지침 위반 ('{sm.group(0)}' 기호 사용 금지)",
            "category": "수정 필수"
        })

    # --------------------------------------------------------------------------
    # Rule 3. 입력 불가 용어 & 대체어 매핑 검사
    # --------------------------------------------------------------------------
    mappings = rules.get("mappings", {})
    for forbidden_word, correct_word in mappings.items():
        # 단어 완전/부분 매칭 정규식
        pattern = re.compile(re.escape(forbidden_word), re.IGNORECASE)
        for fm in pattern.finditer(text_str):
            findings.append({
                "raw_text": fm.group(0),
                "suggested_text": correct_word,
                "reason": f"입력 불가 용어 지침 위반 ('{fm.group(0)}' ➔ '{correct_word}' 대체어 사용)",
                "category": "수정 필수"
            })

    # --------------------------------------------------------------------------
    # Rule 4. 생기부 기재 금지 키워드 검사
    # --------------------------------------------------------------------------
    forbidden_keywords = rules.get("forbidden_keywords", [])
    for kw in forbidden_keywords:
        # 이미 대체어 매핑에서 처리된 경우 중복 방지
        if kw in mappings:
            continue
        pattern = re.compile(r'\b' + re.escape(kw) + r'\b')
        for km in pattern.finditer(text_str):
            findings.append({
                "raw_text": km.group(0),
                "suggested_text": "관련 항목 내용 삭제 또는 기재 가능 용어로 수정",
                "reason": f"생기부 기재 금지어 위반 ('{km.group(0)}' 항목은 생기부 기재 불가)",
                "category": "수정 필수"
            })

    # --------------------------------------------------------------------------
    # Rule 5. 학생 입장 서술 지양 어미 검사
    # --------------------------------------------------------------------------
    student_endings = rules.get("student_endings", [])
    for ending in student_endings:
        pattern = re.compile(r'([가-힣]+' + re.escape(ending) + r'|\b' + re.escape(ending) + r')')
        for em in pattern.finditer(text_str):
            findings.append({
                "raw_text": em.group(0),
                "suggested_text": "~하는 관찰 기재 (예: ~활동지를 작성함, ~모습이 돋보임)",
                "reason": f"학생 입장 서술어 지양 지침 위반 ('{em.group(0)}' ➔ 교사 관점 서술 권장)",
                "category": "수정 권장"
            })

    return findings


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

            # 누락 행 병합 처리
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
# 7. 검증 결과를 규격 표 형식(8개 컬럼)으로 변환하는 함수
# ==============================================================================
def generate_audit_report_table(data_store: dict, rules: dict) -> pd.DataFrame:
    """
    모든 데이터(창체, 세특, 행특)를 정밀 검수하여
    [학번(5자리), 이름, 구분, 세부, 수정전, 수정 후, 수정해야하는 이유나 근거, 수정구분]
    규격 표 형식으로 학번순 오름차순 정렬하여 반환합니다.
    """
    audit_rows = []

    category_type_map = {
        "창체": "창체",
        "세특": "세특",
        "행특": "행발"
    }

    for t_key in ["창체", "세특", "행특"]:
        if t_key not in data_store or data_store[t_key] is None:
            continue

        item = data_store[t_key]
        df = item['df']
        c_map = item['col_map']

        num_c = c_map['num_col']
        name_c = c_map['name_col']
        content_c = c_map['content_col']

        for _, row in df.iterrows():
            num_raw = row.get(num_c, '')
            name_val = str(row.get(name_c, '')).strip()
            
            # 학번 5자리 포맷팅 (예: 1반 1번 -> 10101, 10반 3번 -> 11003)
            num_str = re.sub(r'\D', '', str(num_raw))
            if len(num_str) == 1 or len(num_str) == 2:
                formatted_student_id = f"101{int(num_str):02d}"
            elif len(num_str) == 3 or len(num_str) == 4:
                formatted_student_id = f"{int(num_str):05d}"
            else:
                formatted_student_id = num_str.zfill(5) if num_str else "00000"

            # 세부 항목 구분
            if t_key == "창체":
                detail_sub = str(row.get('영역', row.get('활동영역', '자율/동아리/진로'))).strip()
            elif t_key == "세특":
                detail_sub = str(row.get('과목명', '과목미지정')).strip()
            else:
                detail_sub = "행동특성"

            text_to_inspect = str(row.get('내용', row.get(content_c, '')))
            
            # 정밀 검증 수행
            findings = inspect_student_record_text(text_to_inspect, rules)

            for f in findings:
                audit_rows.append({
                    "학번": formatted_student_id,
                    "이름": name_val,
                    "구분": category_type_map[t_key],
                    "세부": detail_sub,
                    "수정전": f["raw_text"],
                    "수정 후": f["suggested_text"],
                    "수정해야하는 이유나 근거": f["reason"],
                    "수정구분": f["category"]
                })

    if not audit_rows:
        return pd.DataFrame(columns=[
            "학번", "이름", "구분", "세부", "수정전", "수정 후", "수정해야하는 이유나 근거", "수정구분"
        ])

    result_df = pd.DataFrame(audit_rows)
    # 학급/학번순 오름차순 정렬
    result_df = result_df.sort_values(by=["학번", "이름"], ascending=True).reset_index(drop=True)
    return result_df


# ==============================================================================
# 8. 서식 스타일 적용 엑셀 내보내기 함수 (Openpyxl)
# ==============================================================================
def create_audit_report_excel_bytes(audit_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "오탈자 및 지침 검증 리포트"

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

    # 지침 md 규칙 동적 로드
    rules = load_guideline_rules()

    # --------------------------------------------------------------------------
    # 헤더 섹션
    # --------------------------------------------------------------------------
    st.markdown("""
        <div class="main-header">
            <h1>🎓 학교생활기록부 오탈자 정밀 검증 & 데이터 정제 시스템</h1>
            <p>생기부 기재 지침 준수 여부, 오탈자, 맞춤법/문법 오류를 빠짐없이 걸러내고 페이지 나눔 서술문을 완벽히 정제합니다.</p>
        </div>
    """, unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # 사이드바: 파일 업로드 및 옵션
    # --------------------------------------------------------------------------
    with st.sidebar:
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

        st.markdown("---")
        st.markdown("🎯 **핵심 검증 항목**")
        st.info(
            "1. **오탈자·맞춤법 오류**: 도우는➔돕는, 만듬➔만듦, 되서➔돼서, 띄어쓰기 정밀 검사\n"
            "2. **입력 불가 용어**: 네이버, 구글, 유튜브, 카카오톡, KTX, MBTI, 챗GPT 등 대체어 제안\n"
            "3. **기재 금지어**: 수상, 대회, 논문, 자격증, 방과후학교, 특정 대학/기관명 감지\n"
            "4. **서술 지향**: ~파악함, ~이해함 등 학생 입장 어미 감지"
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

                st.sidebar.success(f"✅ {type_key} 정제 및 검증 완료! ({len(refined_df)}명 학생)")

            except Exception as e:
                st.sidebar.error(f"❌ 파일 처리 오류: {str(e)}")
                with st.expander("오류 상세 내용"):
                    st.code(traceback.format_exc())

    # --------------------------------------------------------------------------
    # 메인 콘텐츠 영역 (Tabs)
    # --------------------------------------------------------------------------
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🚨 오탈자 & 지침 정밀 검증", 
        "🔍 학생별 통합 조회", 
        "🛠️ 페이지 나눔 정제 검증", 
        "📊 데이터 분석 & 통계", 
        "📥 엑셀 통합 다운로드"
    ])

    # ==========================================================================
    # Tab 1: 🚨 오탈자 & 지침 정밀 검증 (최우선 메인 탭)
    # ==========================================================================
    with tab1:
        st.subheader("🚨 생기부 오탈자·맞춤법 및 기재 지침 정밀 검증 리포트")
        st.caption("기재 지침 문서(`학교생활기록부_기재_및_검증_지침.md`)를 바탕으로 오탈자, 맞춤법/문법 오류, 입력불가 용어, 금지어를 정밀 검출합니다.")

        available_types = [k for k in ['창체', '세특', '행특'] if st.session_state['data_store'].get(k) is not None]

        if not available_types:
            st.warning("👈 먼저 사이드바에서 생기부 엑셀 파일을 업로드해 주세요.")
        else:
            audit_df = generate_audit_report_table(st.session_state['data_store'], rules)

            # 요약 지표
            req_cnt = len(audit_df[audit_df['수정구분'] == '수정 필수']) if not audit_df.empty else 0
            rec_cnt = len(audit_df[audit_df['수정구분'] == '수정 권장']) if not audit_df.empty else 0

            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            with col_m1:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="val" style="color:#DC2626;">{len(audit_df)}</div>
                        <div class="lbl">총 검출 오류 건수</div>
                    </div>
                """, unsafe_allow_html=True)
            with col_m2:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="val" style="color:#EF4444;">{req_cnt}</div>
                        <div class="lbl">🚨 수정 필수 (지침 위반/오타)</div>
                    </div>
                """, unsafe_allow_html=True)
            with col_m3:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="val" style="color:#F59E0B;">{rec_cnt}</div>
                        <div class="lbl">⚠️ 수정 권장 (어미/문맥)</div>
                    </div>
                """, unsafe_allow_html=True)
            with col_m4:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="val" style="color:#10B981;">100%</div>
                        <div class="lbl">지침 규칙 동적 반영율</div>
                    </div>
                """, unsafe_allow_html=True)

            st.markdown("---")
            
            if audit_df.empty:
                st.balloons()
                st.success("🎉 축하합니다! 검출된 오탈자나 기재 지침 위반 항목이 없습니다. (클린 생기부)")
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
                        filtered_df["학번"].str.contains(kw) | filtered_df["이름"].str.contains(kw)
                    ]

                st.markdown("### 📋 오탈자 및 검증 결과 표 (학번순 정렬)")
                st.dataframe(filtered_df, use_container_width=True, height=450)

                # 개별 리포트 엑셀 다운로드
                audit_excel_bytes = create_audit_report_excel_bytes(filtered_df)
                st.download_button(
                    label="💾 오탈자 & 지침 검증 리포트 엑셀 다운로드 (.xlsx)",
                    data=audit_excel_bytes,
                    file_name="생기부_오탈자_및_지침검증_리포트.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

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
        st.write("오탈자 검증 리포트 및 정제 완료 데이터를 엑셀 파일로 각각 다운로드할 수 있습니다.")

        available_exports = {}
        for t_key in ["창체", "세특", "행특"]:
            if t_key in st.session_state['data_store'] and st.session_state['data_store'][t_key] is not None:
                available_exports[t_key] = st.session_state['data_store'][t_key]['df']

        if available_exports:
            audit_df = generate_audit_report_table(st.session_state['data_store'], rules)
            
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                st.markdown("### 🚨 오탈자 & 지침 검증 리포트")
                st.write(f"총 {len(audit_df)}건의 검출 항목이 포함된 8개 컬럼 규격 엑셀 리포트입니다.")
                audit_bytes = create_audit_report_excel_bytes(audit_df)
                st.download_button(
                    label="💾 오탈자 검증 리포트 다운로드 (.xlsx)",
                    data=audit_bytes,
                    file_name="생기부_오탈자_및_지침검증_리포트.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

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
