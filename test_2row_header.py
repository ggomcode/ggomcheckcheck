import pandas as pd
import re
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

df_raw = pd.read_excel('data/창체_1반.xlsx')
print("Total raw rows:", len(df_raw))

# Header identification
# In this excel, row 2 is ['번 호', '성 명', '학 년', '창의적 체험활동상황']
# and row 3 is [NaN, NaN, NaN, '영 역', '시 간', '특기사항']

# Mapping column indices:
num_idx = 0
name_idx = 1
grade_idx = 2
area_idx = 3
time_idx = 4
content_idx = 5

num_col = "번호"
name_col = "성명"
area_col = "영역"
content_col = "특기사항"

def is_header_or_footer(row_vals) -> bool:
    combined = "".join([str(v).replace(" ", "") for v in row_vals if pd.notna(v)])
    
    # Header check
    if '번호' in combined and '성명' in combined:
        return True
    if '영역' in combined and ('시간' in combined or '특기사항' in combined):
        return True
    if '창의적체험활동상황' in combined:
        return True
        
    # Footer check
    if '포곡고등학교' in combined or '사용자명' in combined or '페이지' in combined:
        return True
    if re.search(r'\d+/\d+\.?\d*', combined) or re.search(r'\d+학년\d+반', combined):
        return True
        
    return False

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

    # Continuous text row from page split
    if is_num_empty and is_name_empty:
        if current_record is not None and content_val:
            current_record[content_col] += " " + content_val
        continue

    # Start new student record
    if current_record is not None:
        refined_records.append(current_record)

    current_record = {
        num_col: str(num_val).strip(),
        name_col: str(name_val).strip(),
        area_col: area_val,
        content_col: content_val,
        '_excel_row': i + 2
    }

if current_record is not None:
    refined_records.append(current_record)

df_res = pd.DataFrame(refined_records)
print(f"\nSuccessfully Refined Records Count: {len(df_res)}")

unique_students = df_res[[num_col, name_col]].drop_duplicates()
print(f"Unique Students Count: {len(unique_students)}")

print("\nArea Breakdown:")
print(df_res[area_col].value_counts())

print("\nSample Output (First 10 records):")
for idx, r in df_res.head(10).iterrows():
    print(f"[{r[num_col]}번] {r[name_col]} | 영역: {r[area_col]} | 글자수: {len(r[content_col])} | 시작부분: {r[content_col][:40]}...")
