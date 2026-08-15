import pandas as pd
import re

df = pd.read_excel('data/창체_1반.xlsx')
print("Raw rows:", len(df))

# Inspect column names
print("Columns:", list(df.columns))

# Find the header row that contains '번호', '성명', '특기사항'
header_idx = None
for idx, row in df.iterrows():
    row_strs = [str(v).strip() for v in row.values if pd.notna(v)]
    if any('번호' in s for s in row_strs) and any('성명' in s for s in row_strs):
        header_idx = idx
        break

print("Detected Header Row Index:", header_idx)

# If header row found, set dataframe columns from that row
if header_idx is not None:
    df.columns = [str(v).strip() for v in df.iloc[header_idx].values]
    df = df.iloc[header_idx + 1:].reset_index(drop=True)

print("Renamed Columns:", list(df.columns))

# Filter function to check if a row is metadata / header / footer
def is_metadata_row(row, num_col, name_col, content_col):
    num_str = str(row.get(num_col, '')).strip()
    name_str = str(row.get(name_col, '')).strip()
    content_str = str(row.get(content_col, '')).strip()
    
    # 1. Header repetition check
    if '번호' in num_str or '성명' in name_str or '특기' in num_str or '특기' in name_str:
        return True
    
    # 2. Page info / School name footer check
    combined = (num_str + " " + name_str + " " + content_str).lower()
    if '포곡고등학교' in combined or '사용자명' in combined or '페이지' in combined or '학교' in combined and len(num_str) > 5:
        return True
    if re.search(r'^\d+\s*/\s*\d+', combined) or re.search(r'^\d+학년\s*\d+반', combined):
        return True
        
    return False

# Detect columns
num_col = None
name_col = None
content_col = None
area_col = None

for c in df.columns:
    c_str = str(c).strip().replace(" ", "")
    if '번호' in c_str and not num_col:
        num_col = c
    elif ('성명' in c_str or '이름' in c_str) and not name_col:
        name_col = c
    elif ('특기' in c_str or '내용' in c_str) and not content_col:
        content_col = c
    elif ('영역' in c_str or '구분' in c_str) and not area_col:
        area_col = c

print(f"Mapped: num_col={num_col}, name_col={name_col}, content_col={content_col}, area_col={area_col}")

# Run parsing test
valid_records = []
current_record = None

for idx, row in df.iterrows():
    if is_metadata_row(row, num_col, name_col, content_col):
        continue
        
    num_val = row.get(num_col)
    name_val = row.get(name_col)
    content_val = str(row.get(content_col, '')).strip() if pd.notna(row.get(content_col)) else ""
    
    is_num_empty = pd.isna(num_val) or str(num_val).strip() in ['', 'nan', 'NaN']
    is_name_empty = pd.isna(name_val) or str(name_val).strip() in ['', 'nan', 'NaN']
    
    # Page break continuation row
    if is_num_empty and is_name_empty:
        if current_record is not None and content_val:
            # Check if text is just footer metadata
            if '포곡고등학교' in content_val or '사용자명' in content_val or re.search(r'^\d+\s*/\s*\d+', content_val):
                continue
            current_record['content'] += " " + content_val
        continue
    
    # Valid student record row
    if current_record:
        valid_records.append(current_record)
        
    current_record = {
        'num': str(num_val).strip(),
        'name': str(name_val).strip(),
        'area': str(row.get(area_col, '')).strip() if area_col else '',
        'content': content_val
    }

if current_record:
    valid_records.append(current_record)

print(f"\nTotal Valid Records Parsed: {len(valid_records)}")

# Unique students count
unique_students = set((r['num'], r['name']) for r in valid_records if r['num'] and r['name'])
print(f"Total Unique Students: {len(unique_students)}")

# Area breakdown
area_counts = {}
for r in valid_records:
    a = r['area']
    area_counts[a] = area_counts.get(a, 0) + 1
print("Records per Area:", area_counts)
