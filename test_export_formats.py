import pandas as pd
import sys
from app import create_audit_report_excel_bytes, create_audit_report_pdf_bytes

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sample_df = pd.DataFrame([
    {
        '학번': '30101',
        '이름': '고윤',
        '구분': '창체',
        '세부': '자율활동',
        '이수시간': '17시간',
        '수정전': '불필요한 이메일을 영구 삭제하는 필수 과제를 제안 함.',
        '수정 후': '불필요한 이메일을 영구 삭제하는 필수 과제를 제안함.',
        '수정해야하는 이유나 근거': "'제안 함' -> '제안함' 띄어쓰기 오류 (동사 파생 접미사 -하다는 붙여 씀)",
        '수정구분': '수정 필수'
    },
    {
        '학번': '30102',
        '이름': '김가온',
        '구분': '세특',
        '세부': '작문',
        '이수시간': '0시간',
        '수정전': '수업 시간에 끝까지 집중하려는 태도를 유지함.',
        '수정 후': '이수시간 0시간 확인 필요 (수업 이수시간이 0시간인데 세특 특기사항이 기록됨)',
        '수정해야하는 이유나 근거': '이수시간 0시간 확인 필요',
        '수정구분': '수정 권장'
    }
])

print("Testing Excel Generator...")
excel_bytes = create_audit_report_excel_bytes(sample_df)
print(f"Excel generation successful! Size: {len(excel_bytes)} bytes")

print("Testing PDF Generator...")
pdf_bytes = create_audit_report_pdf_bytes(sample_df)
print(f"PDF generation successful! Size: {len(pdf_bytes)} bytes")

assert len(excel_bytes) > 0
assert len(pdf_bytes) > 0
print("=== EXPORT FORMATS TEST PASSED! ===")
