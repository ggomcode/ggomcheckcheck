import pandas as pd
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel('data/행특_1빈.xlsx')
print(f"Shape: {df.shape}")

# Row 2 is header: ['번 호', '성  명', '학 년', '행 동 특 성   및   종 합 의 견']
header_idx = None
for i in range(10):
    row_str = "".join([str(v).replace(" ", "") for v in df.iloc[i].values if pd.notna(v)])
    if '번호' in row_str and '성명' in row_str:
        header_idx = i
        break

print("Header Row Index:", header_idx)

# Print first 30 rows of raw data
for i in range(header_idx + 1, header_idx + 35):
    row_vals = df.iloc[i].values
    num = row_vals[0] if pd.notna(row_vals[0]) else ""
    name = row_vals[1] if pd.notna(row_vals[1]) else ""
    grade = row_vals[2] if pd.notna(row_vals[2]) else ""
    content = str(row_vals[3])[:30] if pd.notna(row_vals[3]) else ""
    print(f"Row {i:3d} | Num: {str(num):<4} | Name: {str(name):<5} | Grade: {str(grade):<3} | Content: {content}")
