import pandas as pd
from app import detect_columns, refine_student_records, split_subject_details, create_formatted_excel_bytes

def test_pipeline():
    print("=== Testing Pipeline ===")
    
    # 1. Dummy 세특 Excel data simulating page breaks and colon-separated subjects
    dummy_data = [
        {"번호": 1, "성명": "김철수", "세부능력 및 특기사항": "(1학기)국어: 국어 수업시간에 적극적으로 발표를 수행하"},
        {"번호": None, "성명": None, "세부능력 및 특기사항": "고 독서 활동에 열정을 보임. (2학기)수학: 수학적 사고력이 우수하며 정리 증명을 탐구함."},
        {"번호": 2, "성명": "이영희", "세부능력 및 특기사항": "(1학기)영어: 원문 독해 능력이 뛰어남. (2학기)독서/매체: 시사 이슈 분석 능력이 탁월함."},
        {"번호": 3, "성명": "박민수", "세부능력 및 특기사항": "전체적으로 자기주도 학습 능력이 우수함. (1학기)한국사: 역사의식을 바탕으로 토론에 적극 참여함."},
        {"번호": None, "성명": None, "세부능력 및 특기사항": " (2학기)통합과학: 실험 수행 능력이 정교함."}
    ]
    
    raw_df = pd.DataFrame(dummy_data)
    print("Raw DataFrame:\n", raw_df)
    
    # 2. Detect columns
    col_map = detect_columns(raw_df)
    print("\nDetected Column Mapping:\n", col_map)
    assert col_map['num_col'] == '번호'
    assert col_map['name_col'] == '성명'
    assert col_map['content_col'] == '세부능력 및 특기사항'
    
    # 3. Smart Concatenation (Page break handling)
    refined_df, merge_logs = refine_student_records(raw_df, col_map)
    print("\nRefined DataFrame (Page breaks merged):\n", refined_df[['번호', '성명', '세부능력 및 특기사항']])
    print("\nMerge Logs:\n", merge_logs)
    
    assert len(refined_df) == 3, f"Expected 3 students, got {len(refined_df)}"
    assert len(merge_logs) == 2, f"Expected 2 merged overflow rows, got {len(merge_logs)}"
    
    # Verify Smart concatenation result for 김철수
    chulsoo_text = refined_df.iloc[0]['세부능력 및 특기사항']
    assert "발표를 수행하고" in chulsoo_text, f"Smart concatenation failed: {chulsoo_text}"
    print("\nKim Chulsoo merged text:\n", chulsoo_text)
    
    # 4. Subject Unfolding
    unfolded_df = split_subject_details(refined_df, col_map)
    print("\nUnfolded Subjects DataFrame:\n", unfolded_df[['번호', '성명', '과목명', '내용', '글자수']])
    
    # Check subjects for 김철수
    chulsoo_subjs = unfolded_df[unfolded_df['성명'] == '김철수']['과목명'].tolist()
    print("Kim Chulsoo subjects:", chulsoo_subjs)
    assert '(1학기)국어' in chulsoo_subjs
    assert '(2학기)수학' in chulsoo_subjs
    
    # 5. Test Excel Generation
    excel_bytes = create_formatted_excel_bytes({'세특': unfolded_df})
    assert len(excel_bytes) > 0
    print("\nExcel file created successfully. Byte length:", len(excel_bytes))
    print("\n=== ALL TESTS PASSED! ===")

if __name__ == "__main__":
    test_pipeline()
