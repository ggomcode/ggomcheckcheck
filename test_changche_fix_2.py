import pandas as pd
import re
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

df_raw = pd.read_excel('data/창체_1반.xlsx')

def is_header_or_footer(row_vals) -> bool:
    combined = "".join([str(v).replace(" ", "") for v in row_vals if pd.notna(v)])
    if '번호' in combined and '성명' in combined:
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

num_idx, name_idx, area_idx, content_idx = 0, 1, 3, 5
num_col, name_col, area_col, content_col = "번호", "성명", "영역", "특기사항"

refined_records = []
current_record = None

for i in range(len(df_raw)):
    row_vals = df_raw.iloc[i].values
    if is_header_or_footer(row_vals):
        continue

    num_val = row_vals[num_idx] if len(row_vals) > num_idx else None
    name_val = row_vals[name_idx] if len(row_vals) > name_idx else None
    area_val = str(row_vals[area_idx]).strip() if len(row_vals) > area_idx and pd.notna(row_vals[area_idx]) else ""
    content_val = str(row_vals[content_idx]).strip() if len(row_vals) > content_idx and pd.notna(row_vals[content_idx]) else ""

    is_num_empty = pd.isna(num_val) or str(num_val).strip() in ['', 'nan', 'NaN', 'None']
    is_name_empty = pd.isna(name_val) or str(name_val).strip() in ['', 'nan', 'NaN', 'None']
    
    num_str = "" if is_num_empty else str(num_val).strip()
    name_str = "" if is_name_empty else str(name_val).strip()

    # Case 1: 번호/성명이 비어 있는 경우 -> 이전 학생 기록에 이어붙이기
    if is_num_empty and is_name_empty:
        if current_record is not None and content_val:
            current_record[content_col] += " " + content_val
        continue

    # Case 2: 페이지 넘김 후 이전 학생의 번호/이름이 동일하게 반복된 연속 서술행인 경우
    # (특기사항 텍스트가 동아리명 (동아리명)(시간) 등으로 새로 시작하지 않고 이전 문장 절단 연결인 경우)
    if (current_record is not None and 
        current_record[num_col] == num_str and 
        current_record[name_col] == name_str and 
        (current_record[area_col] == area_val or not area_val)):
        
        # 새로운 과목/동아리 시작 패턴인지 검사 (예: "(동아리명)(16시간)" or "(1학기)국어:")
        is_new_activity_start = re.match(r'^\([가-힣a-zA-Z0-9\s·/]+\)\s*\(\d+시간\)', content_val) or re.match(r'^\([12]학기\)[가-힣·/]+:', content_val)
        
        if not is_new_activity_start:
            # 이전 기록에 자연스럽게 병합
            current_record[content_col] += " " + content_val
            continue

    # Case 3: 신규 레코드 시작
    if current_record is not None:
        refined_records.append(current_record)

    current_record = {
        num_col: num_str,
        name_col: name_str,
        area_col: area_val,
        content_col: content_val,
        '_excel_row': i + 2
    }

if current_record is not None:
    refined_records.append(current_record)

df_res = pd.DataFrame(refined_records)

print(f"\nFinal Refined Records Count: {len(df_res)}")

unique_students = df_res[[num_col, name_col]].drop_duplicates()
print(f"Unique Students Count: {len(unique_students)}")

print("\nArea Breakdown:")
print(df_res[area_col].value_counts())

assert len(df_res) == 93, f"Expected 93 records, got {len(df_res)}"
assert len(unique_students) == 31, f"Expected 31 unique students, got {len(unique_students)}"

print("\n=== ALL CHANGCHE 1BAN PARSING TESTS PASSED! PERFECT 93 RECORDS FOR 31 STUDENTS! ===")
