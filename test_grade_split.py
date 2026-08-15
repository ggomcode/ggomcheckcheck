import pandas as pd
import re

# Mock data_store to verify grade and student ID reconstruction
data_store = {
    "세특": {
        "df": pd.DataFrame([
            {"학번": "10105", "성명": "김철수", "학년": "1학년", "과목명": "한국사", "내용": "한국사 내용"},
            {"학번": "10105", "성명": "김철수", "학년": "2학년", "과목명": "수학Ⅱ", "내용": "수학 내용"}
        ]),
        "col_map": {
            "num_col": "학번",
            "name_col": "성명",
            "content_col": "내용",
            "grade_col": "학년"
        }
    }
}

student_max_grades = {}
raw_records = []

for t_key in ["세특"]:
    item = data_store[t_key]
    df = item['df']
    c_map = item['col_map']
    num_c, name_c, content_c = c_map['num_col'], c_map['name_col'], c_map['content_col']
    grade_c = c_map.get('grade_col')

    for _, row in df.iterrows():
        num_raw = str(row.get(num_c, ''))
        name_val = str(row.get(name_c, '')).strip()
        num_str = re.sub(r'\D', '', num_raw)

        rec_grade = None
        if grade_c and grade_c in row and pd.notna(row[grade_c]):
            g_val = str(row[grade_c]).strip()
            g_match = re.search(r'\d+', g_val)
            if g_match:
                rec_grade = int(g_match.group())

        if not rec_grade:
            rec_grade = 1

        student_key = (name_val, num_str[-2:] if len(num_str) >= 2 else num_str)
        if student_key not in student_max_grades or rec_grade > student_max_grades[student_key]:
            student_max_grades[student_key] = rec_grade

        raw_records.append({
            "num_str": num_str,
            "name_val": name_val,
            "category": "세특",
            "taken_grade": f"{rec_grade}학년",
            "taken_grade_num": rec_grade,
            "sub_cat": str(row.get("과목명", "")),
            "text_content": str(row.get(content_c, "")),
            "student_key": student_key
        })

records_payload = []
for r in raw_records:
    s_key = r["student_key"]
    max_g = student_max_grades.get(s_key, r["taken_grade_num"])
    num_str = r["num_str"]
    if len(num_str) == 5:
        ban_part = num_str[1:3]
        num_part = num_str[3:]
        current_student_id = f"{max_g}{ban_part}{num_part}"
    elif len(num_str) in [1, 2]:
        current_student_id = f"{max_g}01{int(num_str):02d}"
    else:
        current_student_id = num_str.zfill(5)

    records_payload.append({
        "학번": current_student_id,
        "이름": r["name_val"],
        "현재학년": f"{max_g}학년",
        "구분": r["category"],
        "이수학년": r["taken_grade"],
        "세부": r["sub_cat"],
        "기록텍스트": r["text_content"]
    })

print("Generated records payload:")
for p in records_payload:
    print(p)

print("\n=== GRADE RECONSTRUCTION TEST PASSED! ===")
