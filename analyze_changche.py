import pandas as pd

df = pd.read_excel('data/창체_1반.xlsx')

with open('data_analysis.txt', 'w', encoding='utf-8') as out:
    out.write(f"Total shape: {df.shape}\n")
    
    # 1. Print non-null header columns
    out.write(f"Raw Columns: {df.columns.tolist()}\n\n")
    
    # 2. Iterate through all rows and print valid values
    for i in range(len(df)):
        vals = [f"Col{j}:{df.iloc[i, j]}" for j in range(df.shape[1]) if pd.notna(df.iloc[i, j])]
        if vals:
            out.write(f"Row {i:3d}: { ' | '.join(vals) }\n")

print("Analysis written to data_analysis.txt")
