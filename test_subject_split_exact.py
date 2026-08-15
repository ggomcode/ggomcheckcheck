import re
import pandas as pd

def parse_setuk_subjects(raw_text: str, row_default_subject: str = "세부능력및특기사항") -> list:
    """
    세특 텍스트에서 과목명을 추출합니다.
    [규칙]
    1. (1학기) 또는 (2학기)로 시작하는 과목명: (1학기)과목명:
    2. 학기 표시가 없는 과목명: 문두 또는 줄바꿈 직후의 과목명:
    3. 과목명(콜론 전 텍스트, 학기표시 포함)의 길이가 20바이트 이하이어야 함.
    4. 매칭되지 않은 경우 임의의 '공통/기타', '통합/기타' 등을 생성하지 않고 원본/기존 과목명 유지.
    """
    text = str(raw_text).strip() if pd.notna(raw_text) else ""
    if not text:
        return []

    # Regex: 
    # Group 1: 학기표시 (optional): (1학기) / (2학기)
    # Group 2: 과목명 (한글, 숫자, 로마자, 공백, ·, /, Ⅰ-Ⅻ 등)
    # 뒤에 콜론(: 또는 ：)이 위치함
    pattern = r'(?:^|\n|\s{2,})((?:\([12]\s*학기\))?\s*([가-힣a-zA-Z0-9\s·/Ⅰ-Ⅻ()\-_]+))\s*[:：]'

    matches = []
    for m in re.finditer(pattern, text):
        full_hdr = m.group(1).strip()
        subj_name = m.group(2).strip() if m.group(2) else full_hdr
        
        # 20바이트 (EUC-KR 기준 20바이트 또는 UTF-8 기준 20바이트) 제한 검사
        # EUC-KR 기준: 한글 1자=2바이트, 영문/숫자=1바이트 -> "(1학기)수학Ⅱ" = 13바이트 <= 20
        # UTF-8 기준 24바이트 또는 EUC-KR 기준 20바이트
        byte_len_euckr = len(full_hdr.encode('euc-kr', errors='ignore'))
        
        # 콜론 앞 헤더의 바이트 수가 20바이트 이하이고, 줄바꿈 또는 문두에 위치한 경우만 인정
        if byte_len_euckr <= 20 and len(subj_name) >= 1:
            # 특수한 단일 단어(예: 작문, 통계, 윤리, 탐구, 건강 등)가 문장 중간에 나온 것은 제외하기 위해
            # 학기표시가 없으면 텍스트 시작(start == 0)이거나 바로 앞이 줄바꿈인 경우로 엄격 적용
            has_semester = bool(re.match(r'^\([12]\s*학기\)', full_hdr))
            is_at_start = (m.start() == 0 or text[max(0, m.start()-1)] == '\n')
            
            if has_semester or is_at_start:
                matches.append({
                    'start': m.start(1),
                    'end': m.end(),
                    'full_hdr': full_hdr,
                    'subj_name': subj_name
                })

    if not matches:
        return [{
            'subject': row_default_subject if row_default_subject and row_default_subject != '과목미지정' else '세부능력및특기사항',
            'content': text
        }]

    results = []
    # 첫 매칭 전 서두 텍스트가 있다면 첫번째 과목에 포함 또는 기본 과목 처리
    for i in range(len(matches)):
        m_curr = matches[i]
        start_content = m_curr['end']
        end_content = matches[i+1]['start'] if i + 1 < len(matches) else len(text)
        
        content_snippet = text[start_content:end_content].strip()
        subj = m_curr['full_hdr']
        
        results.append({
            'subject': subj,
            'content': content_snippet
        })

    return results

# Test cases
test_samples = [
    "(1학기)수학Ⅱ: 이차방정식과 함수의 관계를 깊이 이해하고 분석함.",
    "(1학기)독서: 고전 문학 작품을 감상하고 작문의 원리를 탐구함. (2학기)화학Ⅰ: 원소의 주기적 성질을 실험함.",
    "한국사: 조선 시대 사회 구조를 분석함. 작문: 글쓰기 능력이 뛰어남.", # '작문:' in sentence vs start
    "수행평가 과정에서 윤리: 도덕적 당위성을 설명함.", # Should NOT match 윤리: because inside sentence
    "특별한 과목 구분 없이 작성된 일련의 세특 특기사항 기록 텍스트입니다."
]

print("=== PARSING RESULTS ===")
for sample in test_samples:
    print(f"\nRAW: {sample}")
    parsed = parse_setuk_subjects(sample)
    for p in parsed:
        print(f"  [과목: {p['subject']}] -> {p['content'][:40]}...")
