import pandas as pd
import re
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

df_raw = pd.read_excel('data/창체_1반.xlsx')

from test_2row_header import is_header_or_footer

num_idx, name_idx, area_idx, content_idx = 0, 1, 3, 5

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

    if is_num_empty and is_name_empty:
        if current_record is not None and content_val:
            current_record['content'] += " " + content_val
        continue

    if current_record is not None:
        refined_records.append(current_record)

    current_record = {
        'num': str(num_val).strip(),
        'name': str(name_val).strip(),
        'area': area_val,
        'content': content_val,
        'excel_row': i + 2
    }

if current_record is not None:
    refined_records.append(current_record)

df_res = pd.DataFrame(refined_records)

# Find동아리활동 records
dongari = df_res[df_res['area'] == '동아리활동']
print(f"Dongari records count: {len(dongari)}")

# Print all dongari records
for idx, r in dongari.iterrows():
    print(f"Row {r['excel_row']:3d} | Num: {r['num']:<5} | Name: {r['name']:<6} | Content snippet: {r['content'][:30]}")
