import sys
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from app import (
    detect_columns, 
    refine_student_records, 
    split_subject_details, 
    load_guideline_content,
    prepare_records_for_llm,
    create_audit_report_excel_bytes
)

def test_llm_pipeline_prep():
    print("=== Testing LLM Pipeline Data Prep ===")
    
    guideline_text = load_guideline_content()
    assert len(guideline_text) > 0, "Guideline content is empty"
    print("Loaded guideline text length:", len(guideline_text))
    
    dummy_data = [
        {
            "번호": 1, "성명": "김철수", 
            "세부능력 및 특기사항": "(1학기)국어: 수업시간에 열심히 발표를 도우는 학생으로 네이버와 유튜브를 적극 활용함. (2학기)수학: 수학 경시대회에서 수상함."
        },
        {
            "번호": None, "성명": None, 
            "세부능력 및 특기사항": "독서 습관이 돋보이며 스스로 다짐함★"
        }
    ]
    
    raw_df = pd.DataFrame(dummy_data)
    col_map = detect_columns(raw_df)
    refined_df, _ = refine_student_records(raw_df, col_map)
    unfolded_df = split_subject_details(refined_df, col_map)
    
    data_store = {
        '세특': {'df': unfolded_df, 'col_map': col_map}
    }
    
    # Test prepare_records_for_llm
    llm_payload = prepare_records_for_llm(data_store)
    print("\nPrepared LLM Payload sample:\n", llm_payload[:2])
    
    assert len(llm_payload) > 0
    assert "학번" in llm_payload[0]
    assert "이름" in llm_payload[0]
    assert "구분" in llm_payload[0]
    assert "세부" in llm_payload[0]
    assert "기록텍스트" in llm_payload[0]
    
    print("\n=== ALL LLM PIPELINE PREP TESTS PASSED! ===")

if __name__ == "__main__":
    test_llm_pipeline_prep()
