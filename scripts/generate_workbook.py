import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

def create_dashboard():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "대시보드"
    ws.views.sheetView[0].showGridLines = True

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # 1. Title Banner
    ws.merge_cells("B2:F3")
    title_cell = ws["B2"]
    title_cell.value = "📊 엑셀 변경사항 자동 취합기 (Excel Diff Aggregator)"
    title_cell.font = Font(name="맑은 고딕", size=15, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # 2. Section 1: File Selection
    ws["B4"] = "1. 파일 경로 지정"
    ws["B4"].font = Font(name="맑은 고딕", size=11, bold=True, color="1F4E78")

    ws["B5"] = "원본(기준) 파일"
    ws["B5"].font = Font(name="맑은 고딕", size=10, bold=True)
    ws["B5"].fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    ws["B5"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B5"].border = thin_border

    ws["C5"] = ""
    ws["C5"].border = thin_border

    ws["B6"] = "수정본 파일"
    ws["B6"].font = Font(name="맑은 고딕", size=10, bold=True)
    ws["B6"].fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    ws["B6"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B6"].border = thin_border

    ws["C6"] = ""
    ws["C6"].border = thin_border

    # 3. Section 2: Options
    ws["B8"] = "2. 취합 옵션 설정"
    ws["B8"].font = Font(name="맑은 고딕", size=11, bold=True, color="1F4E78")

    ws["B9"] = "비교 시트명"
    ws["B9"].font = Font(name="맑은 고딕", size=10, bold=True)
    ws["B9"].fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    ws["B9"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B9"].border = thin_border

    ws["C9"] = "(첫번째 시트)"
    ws["C9"].border = thin_border

    ws["B10"] = "이전 값 주석 남기기"
    ws["B10"].font = Font(name="맑은 고딕", size=10, bold=True)
    ws["B10"].fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    ws["B10"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B10"].border = thin_border

    ws["C10"] = "예"
    ws["C10"].alignment = Alignment(horizontal="center")
    ws["C10"].border = thin_border

    # 4. Instructions & Buttons Guide Card
    ws.merge_cells("B12:F12")
    ws["B12"] = "3. 사용 방법 안내 (VBA 매크로 버튼 매핑)"
    ws["B12"].font = Font(name="맑은 고딕", size=11, bold=True, color="1F4E78")

    instructions = [
        ("버튼 1", "SelectBaseFile", "원본(기준) 엑셀 파일(.xlsx) 선택 경로 C5셀 자동 입력"),
        ("버튼 2", "SelectModFile", "수정본 엑셀 파일(.xlsx) 선택 경로 C6셀 자동 입력"),
        ("버튼 3", "RunAggregator", "메모리 배열 초고속 비교 후 변경 셀 덮어쓰기 & 노란색 하이라이트"),
        ("버튼 4", "ResetHighlights", "원본 파일에 적용된 하이라이트 배경색 제거 초기화")
    ]

    for idx, (btn_name, macro, desc) in enumerate(instructions, start=13):
        ws.cell(row=idx, column=2, value=btn_name).alignment = Alignment(horizontal="center")
        ws.cell(row=idx, column=2).font = Font(name="맑은 고딕", bold=True)
        ws.cell(row=idx, column=2).fill = PatternFill(start_color="D9E1F2", fill_type="solid")
        ws.cell(row=idx, column=2).border = thin_border

        ws.cell(row=idx, column=3, value=macro).font = Font(name="Consolas", color="1F4E78")
        ws.cell(row=idx, column=3).border = thin_border

        ws.merge_cells(start_row=idx, start_column=4, end_row=idx, end_column=6)
        ws.cell(row=idx, column=4, value=desc).font = Font(name="맑은 고딕", size=9)
        ws.cell(row=idx, column=4).border = thin_border

    # Adjust widths
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 35
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 15

    output_path = "ExcelDiffAggregator.xlsx"
    wb.save(output_path)
    print(f"Generated dashboard template: {os.path.abspath(output_path)}")

if __name__ == "__main__":
    create_dashboard()
