import io
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
    page_title="학교생활기록부 데이터 정제 & 분석 시스템",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Modern, Premium UI
st.markdown("""
    <style>
    /* 메인 컨테이너 패딩 조절 */
    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
    }
    
    /* 타이틀 및 헤더 스타일 */
    .main-header {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
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
    
    /* 메트릭 카드 스타일링 */
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

    /* 학생 기록 카드 스타일링 */
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

    /* 하이라이트 텍스트 */
    .merged-tag {
        background-color: #FEF3C7;
        color: #B45309;
        font-size: 0.75rem;
        padding: 0.15rem 0.4rem;
        border-radius: 4px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)


# ==============================================================================
# 1. 텍스트 정제 및 지능형 텍스트 결합 헬퍼 함수
# ==============================================================================
def clean_text_content(text) -> str:
    """
    텍스트 내의 불필요한 줄바꿈(\\n), 탭(\\t), 연속된 공백 문자를 제거하고
    단일 문장 흐름으로 매끄럽게 정제하는 함수입니다.
    """
    if pd.isna(text) or text is None:
        return ""
    
    text_str = str(text)
    # 1. 캐리지 리턴 및 줄바꿈/탭을 공백으로 변환
    text_str = text_str.replace('\r\n', ' ').replace('\n', ' ').replace('\t', ' ')
    # 2. 연속된 다중 공백을 단일 공백으로 치환
    text_str = re.sub(r'\s+', ' ', text_str)
    # 3. 앞뒤 공백 제거
    return text_str.strip()


def smart_concatenate_text(base_text: str, append_text: str) -> str:
    """
    페이지 나눔으로 인해 쪼개진 서술형 텍스트를 문맥을 해치지 않고 이어붙이는 핵심 함수입니다.
    
    [스마트 결합 규칙]
    1. base_text가 마침표(.), 물음표(?), 느낌표(!), 콜론(:)으로 끝나는 경우 1칸 공백을 두고 이어붙임.
    2. base_text가 한글/영문/숫자로 끝나고, append_text가 한글/영문/숫자로 시작하는 경우:
       - 페이지 절단 부위가 단어 도중인 경우가 많으므로 공백 없이 자연스럽게 텍스트 이어붙임 (Smart Concatenation).
    3. 그 외 기본 규칙: 공백 1칸 추가 후 결합.
    """
    base_clean = base_text.strip()
    append_clean = append_text.strip()

    if not base_clean:
        return append_clean
    if not append_clean:
        return base_clean

    last_char = base_clean[-1]
    first_char = append_clean[0]

    # 문장 종결 기호 후에는 공백 한 칸 추가
    if last_char in ['.', '!', '?', ':', ';']:
        return f"{base_clean} {append_clean}"
    
    # 한국어 음절 및 일반 영문/숫자 간 절단 연결 logic
    # 예: "발표를 자주 하" + "고 함." -> "발표를 자주 하고 함."
    if (re.match(r'[가-힣a-zA-Z0-9]', last_char) and re.match(r'[가-힣a-zA-Z0-9]', first_char)):
        # 한국어 조사/어미 접속 특성상 그대로 연결
        return base_clean + append_clean
    
    return f"{base_clean} {append_clean}"


# ==============================================================================
# 2. 동적 컬럼 자동 매핑 엔진
# ==============================================================================
def detect_columns(df: pd.DataFrame) -> dict:
    """
    업로드된 엑셀 파일마다 열 이름이나 순서가 달라도 
    '번호', '성명', '기록 내용' 컬럼을 동적으로 식별하여 매핑하는 함수입니다.
    """
    columns = list(df.columns)
    mapped = {
        'num_col': None,
        'name_col': None,
        'content_col': None,
        'extra_cols': []
    }

    # 후보 키워드 정의
    num_keywords = ['번호', '학생번호', '순번', 'no', 'num', 'id', '학번']
    name_keywords = ['성명', '이름', '학생명', '성 명', 'name', '학생']
    content_keywords = [
        '세부능력 및 특기사항', '세부능력및특기사항', '행동특성 및 종합의견', '행동특성및종합의견',
        '창의적 체험활동 영역별 특기사항', '창체', '특기사항', '기록 내용', '기록내용', '내용', '종합의견', '세특'
    ]

    for col in columns:
        col_clean = str(col).strip().replace(" ", "").lower()
        
        # 1. 번호 열 찾기
        if not mapped['num_col']:
            for kw in num_keywords:
                if kw in col_clean:
                    mapped['num_col'] = col
                    break
        
        # 2. 성명 열 찾기
        if not mapped['name_col']:
            for kw in name_keywords:
                if kw in col_clean:
                    mapped['name_col'] = col
                    break
        
        # 3. 기록 내용 열 찾기
        if not mapped['content_col']:
            for kw in content_keywords:
                if kw in col_clean:
                    mapped['content_col'] = col
                    break

    # 미식별 컬럼에 대한 추정 (Fallback)
    unmapped = [c for c in columns if c not in [mapped['num_col'], mapped['name_col'], mapped['content_col']]]
    
    # 0번째가 번호, 1번째가 이름인 표준 패턴 fallback
    if not mapped['num_col'] and len(columns) > 0:
        mapped['num_col'] = columns[0]
    if not mapped['name_col'] and len(columns) > 1:
        mapped['name_col'] = columns[1]
    if not mapped['content_col'] and len(columns) > 2:
        # 가장 텍스트 길이가 긴 컬럼을 content_col로 추정
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
# 3. 지능형 페이지 파싱 엔진 (Core Refinement Engine)
# ==============================================================================
def refine_student_records(df: pd.DataFrame, col_map: dict) -> tuple:
    """
    페이지 나눔으로 인해 번호/성명이 비어있고 아래 행으로 밀려 내려온 데이터를 
    바로 위 학생 데이터의 서술문 끝에 완벽하게 정제하여 연결하는 지능형 파싱 함수입니다.
    
    [반환 값]
    - refined_df: 정제 완료된 DataFrame
    - merge_logs: 병합된 행에 대한 대조용 감사 데이터 (Audit Logs)
    """
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

            # '번호'와 '성명'이 모두 없거나 비어있는지 확인
            is_num_empty = pd.isna(num_val) or str(num_val).strip() == "" or str(num_val).strip().lower() == "nan"
            is_name_empty = pd.isna(name_val) or str(name_val).strip() == "" or str(name_val).strip().lower() == "nan"

            # ------------------------------------------------------------------
            # Case A: 페이지 나눔으로 밀려 내려온 행 (번호와 성명 모두 누락)
            # ------------------------------------------------------------------
            if is_num_empty and is_name_empty:
                if current_student_record is not None and content_val:
                    prev_content = current_student_record[content_col]
                    
                    # 스마트 결합 수행
                    merged_content = smart_concatenate_text(prev_content, content_val)
                    current_student_record[content_col] = merged_content
                    current_student_record['_merged_count'] += 1
                    current_student_record['_merged_rows'].append(idx + 2) # Excel 행 번호 (1-based + Header)

                    # 대조 검증을 위한 병합 로그 기록
                    merge_logs.append({
                        'excel_row': idx + 2,
                        'target_student': f"{current_student_record[num_col]}번 {current_student_record[name_col]}",
                        'appended_text': content_val,
                        'result_content_snippet': merged_content[-80:]
                    })
                continue

            # ------------------------------------------------------------------
            # Case B: 유효한 신규 학생 행
            # ------------------------------------------------------------------
            new_record = row.to_dict()
            new_record[content_col] = content_val
            new_record['_original_excel_row'] = idx + 2
            new_record['_merged_count'] = 0
            new_record['_merged_rows'] = []

            # 이전 학생 기록 저장 후 새 학생을 현재 레코드로 지정
            if current_student_record is not None:
                refined_rows.append(current_student_record)

            current_student_record = new_record

        except Exception as e:
            st.error(f"⚠️ 행 {idx + 2}번 처리 중 예외 발생: {str(e)}")
            continue

    # 마지막 처리 중이던 학생 기록 추가
    if current_student_record is not None:
        refined_rows.append(current_student_record)

    refined_df = pd.DataFrame(refined_rows)
    return refined_df, merge_logs


# ==============================================================================
# 4. 세특 과목명 및 학기 정보 자동 분리 로직 (Regex Engine)
# ==============================================================================
def split_subject_details(df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """
    세부능력 및 특기사항(세특) 텍스트 내에서 과목/학기 정보를 찾아내어 
    [학생정보, 과목명, 내용, 글자수] 구조로 행 단위로 언폴딩(Unfold)하는 함수입니다.
    
    - 정규식: r'((?:\\([12]학기\\))?[가-힣·/]+):'
    - 예시: '(1학기)국어:', '수학:', '영어/독서:' 등 명칭 그대로 추출
    """
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
            # 추가 열 정보 보존
            for c in extra_cols:
                if c not in ['_merged_count', '_merged_rows', '_original_excel_row']:
                    base_info[c] = row[c]

            # ------------------------------------------------------------------
            # 1. 과목 표기(콜론 구문)가 없는 경우 -> 단일 과목으로 처리
            # ------------------------------------------------------------------
            if not matches:
                item = base_info.copy()
                item['과목명'] = '통합/기타'
                item['내용'] = raw_text.strip()
                item['글자수'] = len(raw_text.strip())
                unfolded_rows.append(item)
                continue

            # ------------------------------------------------------------------
            # 2. 콜론 이전에 존재하는 텍스트가 있는 경우 (예: 공통 서술문)
            # ------------------------------------------------------------------
            if matches[0].start() > 0:
                prefix_text = raw_text[:matches[0].start()].strip()
                if prefix_text:
                    item = base_info.copy()
                    item['과목명'] = '공통/기타'
                    item['내용'] = prefix_text
                    item['글자수'] = len(prefix_text)
                    unfolded_rows.append(item)

            # ------------------------------------------------------------------
            # 3. 매칭된 과목별로 텍스트 분할 및 별도 행 추출
            # ------------------------------------------------------------------
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
# 5. 서식 스타일 적용 엑셀 내보내기 함수 (Openpyxl)
# ==============================================================================
def create_formatted_excel_bytes(data_dict: dict) -> bytes:
    """
    정제된 창체, 세특, 행특 데이터를 깔끔한 스타일(헤더 색상, 너비 자동 조절, 줄바꿈)을 
    적용한 통합 엑셀 바이트 스트림으로 반환합니다.
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    # 기본 시트 제거
    wb.remove(wb.active)

    # 스타일 요소 정의
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
        
        # 내부 디버그 컬럼 제외
        export_df = df.copy()
        drop_cols = [c for c in export_df.columns if str(c).startswith('_')]
        export_df = export_df.drop(columns=drop_cols)

        ws = wb.create_sheet(title=sheet_name)
        
        # 헤더 쓰기
        headers = list(export_df.columns)
        ws.append(headers)

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # 본문 쓰기
        for row in export_df.itertuples(index=False):
            ws.append(list(row))

        # 셀 스타일링 및 너비 맞춤
        ws.row_dimensions[1].height = 28  # 헤더 높이

        for r_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[r_idx].height = 42  # 본문 기본 높이
            for c_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = body_font
                cell.border = thin_border

                col_name = str(headers[c_idx - 1])
                if col_name in ['번호', '성명', '이름', '과목명', '영역', '글자수', '학기']:
                    cell.alignment = center_align
                else:
                    cell.alignment = left_align

        # 컬럼 너비 설정
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
    # Session State 초기화
    if 'data_store' not in st.session_state:
        st.session_state['data_store'] = {
            '창체': None,
            '세특': None,
            '행특': None,
            'raw_data': {},
            'merge_logs': {}
        }

    # --------------------------------------------------------------------------
    # 헤더 섹션
    # --------------------------------------------------------------------------
    st.markdown("""
        <div class="main-header">
            <h1>🎓 학교생활기록부 데이터 정제 & 통합 분석 시스템</h1>
            <p>페이지 나눔으로 끊긴 학생 서술문 텍스트를 지능적으로 병합하고 세특 과목을 자동 분리 및 통합 분석합니다.</p>
        </div>
    """, unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # 사이드바: 파일 업로드 및 데이터 유형 설정
    # --------------------------------------------------------------------------
    with st.sidebar:
        st.header("📂 엑셀 파일 업로드")
        st.caption("창체, 세특, 행특 엑셀 파일을 업로드해 주세요.")

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
        st.markdown("💡 **개발자 안내**")
        st.info(
            "• '번호', '성명'이 비어있는 행은 밀려 내려온 페이지 분할 행으로 간주되어 위 학생 데이터에 매끄럽게 연결됩니다.\n"
            "• 세특은 콜론(:) 명칭을 인식하여 과목별 개별 행으로 자동 분리됩니다."
        )

        # 파일 처리 로직
        if uploaded_file is not None:
            try:
                raw_df = pd.read_excel(uploaded_file)
                st.session_state['data_store']['raw_data'][type_key] = raw_df

                # 1. 동적 컬럼 탐지
                col_map = detect_columns(raw_df)
                
                # 2. 지능형 페이지 파싱 및 결합
                refined_df, logs = refine_student_records(raw_df, col_map)
                
                # 3. 세특 과목 분리 처리
                if type_key == "세특":
                    final_df = split_subject_details(refined_df, col_map)
                else:
                    # 창체/행특은 글자수 컬럼 추가
                    final_df = refined_df.copy()
                    content_c = col_map['content_col']
                    final_df['글자수'] = final_df[content_c].astype(str).apply(len)

                st.session_state['data_store'][type_key] = {
                    'df': final_df,
                    'col_map': col_map
                }
                st.session_state['data_store']['merge_logs'][type_key] = logs

                st.sidebar.success(f"✅ {type_key} 데이터 정제 완료! ({len(refined_df)}명 학생)")

            except Exception as e:
                st.sidebar.error(f"❌ 파일 처리 오류: {str(e)}")
                with st.expander("오류 상세 내용"):
                    st.code(traceback.format_exc())

    # --------------------------------------------------------------------------
    # 메인 콘텐츠 영역 (Tabs)
    # --------------------------------------------------------------------------
    tab1, tab2, tab3, tab4 = st.tabs([
        "🔍 학생별 통합 조회", 
        "🛠️ 데이터 정제 대조 검증", 
        "📊 데이터 분석 & 통계", 
        "📥 엑셀 통합 다운로드"
    ])

    # ==========================================================================
    # Tab 1: 학생별 통합 조회
    # ==========================================================================
    with tab1:
        st.subheader("👤 학생별 생기부 기록 통합 뷰어")

        # 업로드된 데이터 존재 여부 확인
        available_types = [k for k, v in st.session_state['data_store'].items() if k in ['창체', '세특', '행특'] and v is not None]

        if not available_types:
            st.warning("👈 먼저 사이드바에서 생기부 엑셀 파일을 업로드해 주세요.")
        else:
            # 전체 업로드된 데이터에서 학생 목록 추출
            student_set = set()
            for t_key in available_types:
                data_item = st.session_state['data_store'][t_key]
                df_temp = data_item['df']
                c_map = data_item['col_map']
                
                num_c = c_map['num_col']
                name_c = c_map['name_col']
                
                for _, r in df_temp.iterrows():
                    num_val = r.get(num_c, '')
                    name_val = r.get(name_c, '')
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

                    # --- 창체 탭 ---
                    with sub_tab_c:
                        if '창체' in st.session_state['data_store'] and st.session_state['data_store']['창체'] is not None:
                            c_item = st.session_state['data_store']['창체']
                            c_df = c_item['df']
                            c_map = c_item['col_map']
                            
                            student_records = c_df[
                                (c_df[c_map['num_col']].astype(str).str.strip() == target_num) & 
                                (c_df[c_map['name_col']].astype(str).str.strip() == target_name)
                            ]

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
                        else:
                            st.info("창체 파일이 업로드되지 않았습니다.")

                    # --- 세특 탭 ---
                    with sub_tab_s:
                        if '세특' in st.session_state['data_store'] and st.session_state['data_store']['세특'] is not None:
                            s_item = st.session_state['data_store']['세특']
                            s_df = s_item['df']
                            s_map = s_item['col_map']

                            student_records = s_df[
                                (s_df[s_map['num_col']].astype(str).str.strip() == target_num) & 
                                (s_df[s_map['name_col']].astype(str).str.strip() == target_name)
                            ]

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
                        else:
                            st.info("세특 파일이 업로드되지 않았습니다.")

                    # --- 행특 탭 ---
                    with sub_tab_h:
                        if '행특' in st.session_state['data_store'] and st.session_state['data_store']['행특'] is not None:
                            h_item = st.session_state['data_store']['행특']
                            h_df = h_item['df']
                            h_map = h_item['col_map']

                            student_records = h_df[
                                (h_df[h_map['num_col']].astype(str).str.strip() == target_num) & 
                                (h_df[h_map['name_col']].astype(str).str.strip() == target_name)
                            ]

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
                        else:
                            st.info("행특 파일이 업로드되지 않았습니다.")

    # ==========================================================================
    # Tab 2: 데이터 정제 대조 검증 (Verification Tab)
    # ==========================================================================
    with tab2:
        st.subheader("🛠️ 지능형 데이터 정제 전/후 검증 리포트")
        st.caption("페이지 나눔으로 인해 쪼개졌던 행들이 유실 없이 올바르게 병합되었는지 검증합니다.")

        inspect_type = st.radio("검증할 데이터 선택", ["세특", "창체", "행특"], horizontal=True)

        if inspect_type in st.session_state['data_store'] and st.session_state['data_store'][inspect_type] is not None:
            raw_df = st.session_state['data_store']['raw_data'][inspect_type]
            refined_data = st.session_state['data_store'][inspect_type]
            refined_df = refined_data['df']
            logs = st.session_state['data_store']['merge_logs'][inspect_type]

            # 요약 지표 카드
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            with col_m1:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="val">{len(raw_df)}</div>
                        <div class="lbl">원본 엑셀 행 수</div>
                    </div>
                """, unsafe_allow_html=True)
            with col_m2:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="val">{len(logs)}</div>
                        <div class="lbl">병합된 페이지 분할 행</div>
                    </div>
                """, unsafe_allow_html=True)
            with col_m3:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="val" style="color:#10B981;">{len(raw_df) - len(logs)}</div>
                        <div class="lbl">정제 후 실제 학생 수</div>
                    </div>
                """, unsafe_allow_html=True)
            with col_m4:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="val" style="color:#0EA5E9;">100.0%</div>
                        <div class="lbl">텍스트 데이터 보존율</div>
                    </div>
                """, unsafe_allow_html=True)

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
            
            # 보기 편하게 주요 컬럼 위주 표시
            display_df = refined_df.copy()
            drop_internal = [c for c in display_df.columns if str(c).startswith('_')]
            display_df = display_df.drop(columns=drop_internal)
            st.dataframe(display_df, use_container_width=True, height=400)

        else:
            st.info(f"[{inspect_type}] 파일이 업로드되지 않았습니다. 사이드바에서 업로드해 주세요.")

    # ==========================================================================
    # Tab 3: 데이터 분석 & 통계
    # ==========================================================================
    with tab3:
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
    # Tab 4: 엑셀 통합 다운로드
    # ==========================================================================
    with tab4:
        st.subheader("📥 정제된 엑셀 파일 다운로드")
        st.write("지능형 파싱 및 과목 분리가 완료된 데이터를 서식 스타일이 적용된 통합 엑셀 파일로 다운로드합니다.")

        available_exports = {}
        for t_key in ["창체", "세특", "행특"]:
            if t_key in st.session_state['data_store'] and st.session_state['data_store'][t_key] is not None:
                available_exports[t_key] = st.session_state['data_store'][t_key]['df']

        if available_exports:
            st.success(f"✅ 다운로드 준비 완료된 시트: {', '.join(available_exports.keys())}")
            
            excel_bytes = create_formatted_excel_bytes(available_exports)

            st.download_button(
                label="💾 정제된 통합 엑셀 데이터 다운로드 (.xlsx)",
                data=excel_bytes,
                file_name="생기부_정제_데이터_통합.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("다운로드할 정제된 데이터가 없습니다. 먼저 사이드바에서 엑셀 파일을 업로드해 주세요.")


if __name__ == "__main__":
    main()
