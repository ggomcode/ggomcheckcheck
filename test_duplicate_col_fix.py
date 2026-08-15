import pandas as pd

def deduplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)
    seen = {}
    new_cols = []
    for c in cols:
        c_str = str(c).strip()
        if c_str in seen:
            seen[c_str] += 1
            new_cols.append(f"{c_str}_{seen[c_str]}")
        else:
            seen[c_str] = 0
            new_cols.append(c_str)
    df.columns = new_cols
    return df

# Create dummy dataframe with duplicate columns
df_dup = pd.DataFrame([
    ['1', '고윤', '특기1', '특기2'],
    ['2', '김가온', '특기3', '특기4']
], columns=['번 호', '성 명', '특기사항', '특기사항'])

print("Before deduplication columns:", df_dup.columns.tolist())
# Accessing df_dup['특기사항'] gives DataFrame!
print("Type of df_dup['특기사항']:", type(df_dup['특기사항']))

try:
    df_dup['특기사항'].astype(str).str.strip()
except AttributeError as e:
    print("Caught expected error:", e)

df_clean = deduplicate_columns(df_dup)
print("\nAfter deduplication columns:", df_clean.columns.tolist())
print("Type of df_clean['특기사항']:", type(df_clean['특기사항']))
print("String operation result:", df_clean['특기사항'].astype(str).str.strip())

print("\n=== DEDUPLICATION FIX VERIFIED! ===")
