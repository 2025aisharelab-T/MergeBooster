import subprocess
import os
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

def test_selective_reset():
    # 1. Base File setup with original green color and original comment
    wb_base = openpyxl.Workbook()
    ws_base = wb_base.active
    ws_base.title = "직원목록"
    
    ws_base.append(["사번", "성명", "부서"])
    ws_base.append([1001, "홍길동", "영업1팀"])
    ws_base.append([1002, "김철수", "개발1팀"])
    
    # Original green fill on header A1
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    ws_base["A1"].fill = green_fill
    
    # Original user comment on B2
    ws_base["B2"].comment = Comment("중요 직원 사원번호", "관리자")
    
    base_file = os.path.abspath("test_base.xlsx")
    wb_base.save(base_file)
    print("Base file created with original green fill on A1 and original comment on B2.")
    
    # 2. Mod File setup (Change Department for Kim Cheolsu)
    wb_mod = openpyxl.Workbook()
    ws_mod = wb_mod.active
    ws_mod.title = "직원목록"
    ws_mod.append(["사번", "성명", "부서"])
    ws_mod.append([1001, "홍길동", "영업1팀"])
    ws_mod.append([1002, "김철수", "AI개발팀"]) # Changed!
    
    mod_file = os.path.abspath("test_mod.xlsx")
    wb_mod.save(mod_file)
    print("Mod file created with department change.")

    print("\n--- Test complete ---")

if __name__ == "__main__":
    test_selective_reset()
