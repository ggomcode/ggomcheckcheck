import pandas as pd
import sys
import os

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from app import detect_columns, refine_student_records, split_subject_details

def test_file(filepath: str, type_key: str):
    print(f"\n==========================================")
    print(f"Testing [{type_key}] File: {filepath}")
    print(f"==========================================")
    
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        return

    raw_df = pd.read_excel(filepath)
    print(f"Raw Excel shape: {raw_df.shape}")

    df_processed, col_map = detect_columns(raw_df)
    print(f"Detected col_map: {col_map}")

    refined_df, merge_logs = refine_student_records(df_processed, col_map)
    print(f"Refined Rows Count: {len(refined_df)}")
    print(f"Merge Logs Count (Page break overflows merged): {len(merge_logs)}")

    if type_key == "세특":
        final_df = split_subject_details(refined_df, col_map)
    else:
        final_df = refined_df.copy()

    num_c, name_c = col_map['num_col'], col_map['name_col']
    unique_students = final_df[[num_c, name_c]].drop_duplicates()
    
    print(f"Unique Students Count: {len(unique_students)}")
    print(f"Total Output Final Rows: {len(final_df)}")

    if type_key == "세특" and '과목명' in final_df.columns:
        print("\nSubject Value Counts (Top 10):")
        print(final_df['과목명'].value_counts().head(10))

    print("\nSample Output (First 5 records):")
    print(final_df.head(5))

if __name__ == "__main__":
    test_file("data/세특_1반.xlsx", "세특")
    test_file("data/행특_1빈.xlsx", "행특")
