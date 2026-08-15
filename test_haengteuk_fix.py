import pandas as pd
import re
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

df_raw = pd.read_excel('data/행특_1빈.xlsx')

def is_header_or_footer(row_vals) -> bool:
    combined = "".join([str(v).replace(" ", "") for v in row_vals if pd.notna(v)])
    if '번호' in combined and '성명' in combined:
        return True
    if '행동특성' in combined or '종합의견' in combined:
        return True
    if '포곡고등학교' in combined or '사용자명' in combined or '페이지' in combined:
        return True
    if re.search(r'\d+/\d+\.?\d*', combined) or re.search(r'\d+학년\d+반', combined):
        return True
    return False

# Detect header row
header_idx = None
for i in range(10):
    row_str = "".join([str(v).replace(" ", "") for v in df_raw.iloc[i].values if pd.notna(v)])
    if '번호' in row_str and '성명' in row_str:
        header_idx = i
        break

df = df_raw.copy()
if header_idx is not None:
    df.columns = [str(v).strip() for v in df.iloc[header_idx].values]
    df = df.iloc[header_idx + 1:].reset_index(drop=True)

# Map columns
num_col, name_col, grade_col, content_col = None, None, None, None
for c in df.columns:
    c_clean = str(c).replace(" ", "").lower()
    if '번호' in c_clean and not num_col:
        num_col = c
    elif ('성명' in c_clean or '이름' in c_clean) and not name_col:
        name_col = c
    elif ('학년' in c_clean) and not grade_col:
        grade_col = c
    elif ('행동' in c_clean or '종합' in c_clean or '내용' in c_clean or '특기' in c_clean) and not content_col:
        content_col = c

print(f"Mapped: num='{num_col}', name='{name_col}', grade='{grade_col}', content='{content_col}'")

refined_records = []
current_student_num = ""
current_student_name = ""
current_record = None

for i, row in df.iterrows():
    row_vals = row.values
    if is_header_or_footer(row_vals):
        continue

    num_val = row.get(num_col)
    name_val = row.get(name_col)
    grade_val = row.get(grade_col) if grade_col else None
    content_val = str(row.get(content_col, '')).strip() if pd.notna(row.get(content_col)) else ""

    is_num_empty = pd.isna(num_val) or str(num_val).strip() in ['', 'nan', 'NaN', 'None']
    is_name_empty = pd.isna(name_val) or str(name_val).strip() in ['', 'nan', 'NaN', 'None']
    is_grade_empty = pd.isna(grade_val) or str(grade_val).strip() in ['', 'nan', 'NaN', 'None']

    num_str = "" if is_num_empty else str(num_val).strip()
    name_str = "" if is_name_empty else str(name_val).strip()
    grade_str = "" if is_grade_empty else str(grade_val).strip().replace(".0", "")

    # Update active student info
    if num_str and name_str:
        current_student_num = num_str
        current_student_name = name_str

    # Case A: 번호/이름이 비어 있고, 학년(Grade) 정보만 존재하는 경우 -> 동일 학생의 새로운 학년 레코드
    if is_num_empty and is_name_empty and grade_str:
        if current_record is not None:
            refined_records.append(current_record)
        current_record = {
            num_col: current_student_num,
            name_col: current_student_name,
            grade_col: grade_str + "학년",
            content_col: content_val
        }
        continue

    # Case B: 번호/이름/학년이 모두 비어 있거나, 페이지 넘김 후 번호/이름만 중복된 연장 서술행인 경우
    if (is_num_empty and is_name_empty and is_grade_empty) or \
       (current_record is not None and num_str == current_student_num and name_str == current_student_name and is_grade_empty):
        if current_record is not None and content_val:
            current_record[content_col] += " " + content_val
        continue

    # Case C: 번호/이름이 새로 시작되는 신규 레코드
    if current_record is not None:
        refined_records.append(current_record)

    current_record = {
        num_col: current_student_num,
        name_col: current_student_name,
        grade_col: (grade_str + "학년") if grade_str else "",
        content_col: content_val
    }

if current_record is not None:
    refined_records.append(current_record)

df_res = pd.DataFrame(refined_records)

print(f"\nFinal Refined Rows: {len(df_res)}")
unique_students = df_res[[num_col, name_col]].drop_duplicates()
print(f"Unique Students Count: {len(unique_students)}")

print("\nGrade Breakdown:")
print(df_res[grade_col].value_counts())

print("\nFirst 10 records:")
for idx, r in df_res.head(10).iterrows():
    print(f"[{r[num_col]}번] {r[name_col]} ({r[grade_col]}) | 내용: {r[content_col][:30]}...")
