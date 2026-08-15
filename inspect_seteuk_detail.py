import pandas as pd
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from app import detect_columns, refine_student_records, split_subject_details

raw_df = pd.read_excel('data/세특_1반.xlsx')
print(f"Raw shape: {raw_df.shape}")

df_processed, col_map = detect_columns(raw_df)
print(f"col_map: {col_map}")

refined_df, merge_logs = refine_student_records(df_processed, col_map)
print(f"Refined rows: {len(refined_df)}")

num_c, name_c = col_map['num_col'], col_map['name_col']
students = refined_df[[num_c, name_c]].drop_duplicates()
print(f"Unique students count: {len(students)}")
print(students.to_string())
