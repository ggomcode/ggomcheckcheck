import pandas as pd
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel('data/행특_1빈.xlsx')
print("Total rows in 행특_1빈.xlsx:", len(df))

for i in range(len(df)):
    vals = [f"Col{j}:{df.iloc[i, j]}" for j in range(df.shape[1]) if pd.notna(df.iloc[i, j])]
    if vals:
        print(f"Row {i:3d}: {' | '.join(vals)}")
