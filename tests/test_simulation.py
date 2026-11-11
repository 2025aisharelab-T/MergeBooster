import openpyxl
import os

def generate_test_excel_files():
    # 1. Create Base Sample Excel File
    wb_base = openpyxl.Workbook()
    ws_base = wb_base.active
    ws_base.title = "직원목록"
    
    headers = ["사번", "성명", "부서", "직급", "전화번호", "비고"]
    ws_base.append(headers)
    
    rows_base = [
        [1001, "홍길동", "영업1팀", "대리", "010-1111-2222", "기존"],
        [1002, "김철수", "개발1팀", "사원", "010-2222-3333", "기존"],
        [1003, "이영희", "기획팀", "팀장", "010-3333-4444", "기존"],
        [1004, "박민수", "마케팅팀", "과장", "010-4444-5555", "기존"],
        [1005, "최유리", "인사팀", "대리", "010-5555-6666", "기존"]
    ]
    for r in rows_base:
        ws_base.append(r)
        
    # 1-2. Add Second Sheet "부서정보" to Base
    ws_base2 = wb_base.create_sheet(title="부서정보")
    ws_base2.append(["부서코드", "부서명", "팀장"])
    ws_base2.append(["D01", "영업1팀", "김부장"])
    ws_base2.append(["D02", "개발1팀", "이부장"])
    ws_base2.append(["D03", "기획팀", "박부장"])

    base_file = "sample_base.xlsx"
    wb_base.save(base_file)
    print(f"Sample Base File Created (Multi-Sheet): {os.path.abspath(base_file)}")

    # 2. Create Modified Sample Excel File (Kim Cheolsu department & Park Minsu position changed)
    wb_mod = openpyxl.Workbook()
    ws_mod = wb_mod.active
    ws_mod.title = "직원목록"
    
    ws_mod.append(headers)
    
    rows_mod = [
        [1001, "홍길동", "영업1팀", "대리", "010-1111-2222", "기존"],
        [1002, "김철수", "AI개발팀", "사원", "010-2222-3333", "기존"],  # 부서 변경: 개발1팀 -> AI개발팀
        [1003, "이영희", "기획팀", "팀장", "010-3333-4444", "기존"],
        [1004, "박민수", "마케팅팀", "차장", "010-4444-5555", "승진 반영"], # 직급/비고 변경
        [1005, "최유리", "인사팀", "대리", "010-5555-6666", "기존"]
    ]
    for r in rows_mod:
        ws_mod.append(r)

    # 2-2. Add Second Sheet "부서정보" to Mod (Change Team Leader for D02)
    ws_mod2 = wb_mod.create_sheet(title="부서정보")
    ws_mod2.append(["부서코드", "부서명", "팀장"])
    ws_mod2.append(["D01", "영업1팀", "김부장"])
    ws_mod2.append(["D02", "개발1팀", "최부장"]) # 팀장 변경: 이부장 -> 최부장
    ws_mod2.append(["D03", "기획팀", "박부장"])
        
    mod_file = "sample_mod.xlsx"
    wb_mod.save(mod_file)
    print(f"Sample Modified File Created (Multi-Sheet): {os.path.abspath(mod_file)}")

if __name__ == "__main__":
    generate_test_excel_files()
