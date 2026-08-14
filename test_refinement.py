import sys
import io
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from app import (
    detect_columns, 
    refine_student_records, 
    split_subject_details, 
    load_guideline_rules,
    inspect_student_record_text,
    generate_audit_report_table,
    create_audit_report_excel_bytes
)

def test_typo_and_guideline_engine():
    print("=== Testing Typo & Guideline Inspection Engine ===")
    
    rules = load_guideline_rules()
    print("Loaded guideline mappings count:", len(rules["mappings"]))
    
    # 1. Dummy data with typos, forbidden words, and guideline issues
    dummy_data = [
        {
            "번호": 1, "성명": "김철수", 
            "세부능력 및 특기사항": "(1학기)국어: 수업시간에 열심히 발표를 도우는 학생으로 네이버와 유튜브를 적극 활용함. (2학기)수학: 수학 경시대회에서 수상함."
        },
        {
            "번호": None, "성명": None, 
            "세부능력 및 특기사항": "독서 습관이 돋보이며 스스로 다짐함★"
        },
        {
            "번호": 2, "성명": "이영희", 
            "세부능력 및 특기사항": "(1학기)영어: KTX를 타고 서울의 롯데타워를 방문하여 영어를 연습함. 내용이 이해함."
        }
    ]
    
    raw_df = pd.DataFrame(dummy_data)
    col_map = detect_columns(raw_df)
    refined_df, _ = refine_student_records(raw_df, col_map)
    unfolded_df = split_subject_details(refined_df, col_map)
    
    data_store = {
        '세특': {'df': unfolded_df, 'col_map': col_map}
    }
    
    # 2. Generate Audit Report Table
    audit_table = generate_audit_report_table(data_store, rules)
    print("\nGenerated Audit Report Table (8 columns):\n")
    print(audit_table.to_string())
    
    # Assertions
    assert "학번" in audit_table.columns
    assert "이름" in audit_table.columns
    assert "구분" in audit_table.columns
    assert "세부" in audit_table.columns
    assert "수정전" in audit_table.columns
    assert "수정 후" in audit_table.columns
    assert "수정해야하는 이유나 근거" in audit_table.columns
    assert "수정구분" in audit_table.columns
    
    # Check for specific detected typos/issues
    detected_raws = audit_table["수정전"].tolist()
    print("\nDetected error items count:", len(detected_raws))
    
    assert any("도우는" in item for item in detected_raws), "Typo '도우는' not detected"
    assert any("네이버" in item for item in detected_raws), "Forbidden term '네이버' not detected"
    assert any("유튜브" in item for item in detected_raws), "Forbidden term '유튜브' not detected"
    assert any("★" in item for item in detected_raws), "Special character '★' not detected"
    assert any("KTX" in item for item in detected_raws), "Forbidden term 'KTX' not detected"
    
    # Test Audit Excel Generation
    audit_excel_bytes = create_audit_report_excel_bytes(audit_table)
    assert len(audit_excel_bytes) > 0
    print("\nAudit Excel Report generated successfully. Byte count:", len(audit_excel_bytes))
    print("\n=== ALL TYPO ENGINE TESTS PASSED! ===")

if __name__ == "__main__":
    test_typo_and_guideline_engine()
