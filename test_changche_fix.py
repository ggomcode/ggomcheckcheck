import pandas as pd
import re

def is_header_or_footer_row(row_dict: dict, num_col, name_col) -> bool:
    """
    엑셀의 매 페이지마다 반복되는 제목, 헤더('번호', '성명'), 푸터('포곡고등학교', '사용자명', '페이지 번호')를 걸러냅니다.
    """
    vals = [str(v).strip() for v in row_dict.values() if pd.notna(v)]
    combined = "".join(vals).replace(" ", "")

    num_val_str = str(row_dict.get(num_col, '')).replace(" ", "")
    name_val_str = str(row_dict.get(name_col, '')).replace(" ", "")

    # 1. 헤더 행 탐지 ('번호'와 '성명'이 컬럼명/헤더로 들어있는 경우)
    if '번호' in num_val_str or '성명' in name_val_str:
        return True
    if '영역' in combined and ('시간' in combined or '특기사항' in combined):
        return True
    if '창의적체험활동상황' in combined:
        return True

    # 2. 푸터/페이지 정보 탐지
    if '포곡고등학교' in combined or '사용자명' in combined or '페이지' in combined:
        return True
    if re.search(r'\d+/\d+\.?\d*', combined) or re.search(r'\d+학년\d+반', combined):
        return True

    return False

def parse_excel_with_repeated_headers(filepath: str):
    df_raw = pd.read_excel(filepath)
    
    # 헤더 행 찾기 (0~15행 탐색)
    header_idx = None
    for idx in range(min(15, len(df_raw))):
        row_str = "".join([str(v).replace(" ", "") for v in df_raw.iloc[idx].values if pd.notna(v)])
        if '번호' in row_str and '성명' in row_str:
            header_idx = idx
            break

    df = df_raw.copy()
    if header_idx is not None:
        df.columns = [str(v).strip() for v in df.iloc[header_idx].values]
        df = df.iloc[header_idx + 1:].reset_index(drop=True)

    # 동적 컬럼 탐지 (공백 제거 후 검색)
    num_col, name_col, content_col, area_col = None, None, None, None
    for c in df.columns:
        c_clean = str(c).strip().replace(" ", "").lower()
        if '번호' in c_clean and not num_col:
            num_col = c
        elif ('성명' in c_clean or '이름' in c_clean) and not name_col:
            name_col = c
        elif ('특기' in c_clean or '내용' in c_clean) and not content_col:
            content_col = c
        elif ('영역' in c_clean or '구분' in c_clean) and not area_col:
            area_col = c

    print(f"Header Row Index: {header_idx}")
    print(f"Detected columns: num='{num_col}', name='{name_col}', content='{content_col}', area='{area_col}'")

    refined_records = []
    current_record = None

    for idx, row in df.iterrows():
        row_dict = row.to_dict()
        
        # 헤더/푸터 메타데이터 행 건너뛰기
        if is_header_or_footer_row(row_dict, num_col, name_col):
            continue

        num_val = row_dict.get(num_col)
        name_val = row_dict.get(name_col)
        content_val = str(row_dict.get(content_col, '')).strip() if pd.notna(row_dict.get(content_col)) else ""
        area_val = str(row_dict.get(area_col, '')).strip() if area_col and pd.notna(row_dict.get(area_col)) else ""

        is_num_empty = pd.isna(num_val) or str(num_val).strip() in ['', 'nan', 'NaN', 'None']
        is_name_empty = pd.isna(name_val) or str(name_val).strip() in ['', 'nan', 'NaN', 'None']

        # 페이지 나눔으로 밀려 내려온 연속 서술문인 경우
        if is_num_empty and is_name_empty:
            if current_record is not None and content_val:
                current_record[content_col] += " " + content_val
            continue

        # 유효한 학생 데이터 시작
        if current_record is not None:
            refined_records.append(current_record)

        current_record = {
            num_col: str(num_val).strip(),
            name_col: str(name_val).strip(),
            '영역': area_val,
            content_col: content_val
        }

    if current_record is not None:
        refined_records.append(current_record)

    df_refined = pd.DataFrame(refined_records)
    return df_refined, num_col, name_col, content_col

# Run test
df_res, num_col, name_col, content_col = parse_excel_with_repeated_headers('data/창체_1반.xlsx')
print(f"\nRefined Rows Count: {len(df_res)}")

unique_students = df_res[[num_col, name_col]].drop_duplicates()
print(f"Unique Students Count: {len(unique_students)}")

print("\nArea Breakdown:")
print(df_res['영역'].value_counts())

print("\nSample Output (First 5 records):")
print(df_res[[num_col, name_col, '영역', content_col]].head(5))
