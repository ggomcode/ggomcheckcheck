import pandas as pd

df = pd.read_excel('data/창체_1반.xlsx')
print("Total rows:", len(df))

# Find row indices where header '번호' appears
for idx, row in df.iterrows():
    row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
    if '번호' in row_str and '성명' in row_str:
        print(f"Header Row at index {idx}: {row_str[:80]}")
    elif '자율' in row_str or '동아리' in row_str or '진로' in row_str:
        print(f"Area Row at index {idx}: {row_str[:80]}")

# Print first 40 rows completely to see structure
for i in range(40):
    vals = [str(v) for v in df.iloc[i].values if pd.notna(v) and str(v).strip() != '']
    if vals:
        print(f"Row {i:3d}: {vals}")
