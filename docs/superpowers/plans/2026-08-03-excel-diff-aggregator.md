# 엑셀 변경사항 자동 취합기 (Excel Diff Aggregator) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 필터 및 숨겨진 행의 영향 없이, 원본 대비 수정본 엑셀 문서에서 변경된 셀만 추출하여 원본에 자동 덮어쓰기 및 노란색 하이라이트를 적용하는 전용 VBA 취합기(`ExcelDiffAggregator.xlsm`) 및 빌더 환경 구축

**Architecture:** 대시보드 시트 컨트롤 UI와 4개의 분리된 VBA 모듈(`mod_Main`, `mod_DiffEngine`, `mod_FileDialog`, `mod_Utils`)로 구성. 메모리(Variant Array) 상에서 데이터를 비동기 비교하여 속도 및 필터/숨김 행 무관 1:1 절대좌표 매칭 보장.

**Tech Stack:** VBA (Visual Basic for Applications), Python (`openpyxl` / `xlsxwriter` for automated test data generation & workbook creation).

## Global Constraints
- Target File: `ExcelDiffAggregator.xlsm`
- Highlight Color: Light Yellow RGB(255, 242, 204) / Hex `#FFF2CC`
- Fast Array Diff: Memory Variant Array comparison (`OriginalVal <> ModifiedVal`)
- Safe Cleanup: Restore `ScreenUpdating = True` and `Calculation = xlCalculationAutomatic` on any exit/error

---

### Task 1: VBA 모듈 작성 (`mod_Utils.bas` & `mod_FileDialog.bas`)

**Files:**
- Create: `src/vba/mod_Utils.bas`
- Create: `src/vba/mod_FileDialog.bas`

**Interfaces:**
- Consumes: Native VBA `Application` methods and `FileDialog(msoFileDialogFilePicker)`
- Produces: 
  - `mod_Utils.EnableFastMode()`: Toggles ScreenUpdating/Calculation off
  - `mod_Utils.DisableFastMode()`: Restores settings
  - `mod_FileDialog.GetFilePath(Title As String) As String`: Single file selector
  - `mod_FileDialog.GetMultipleFilePaths(Title As String) As Collection`: Multi-file selector

- [ ] **Step 1: Write `src/vba/mod_Utils.bas`**

```vba
Attribute VB_Name = "mod_Utils"
Option Explicit

Private SavedCalcMode As Long

Public Sub EnableFastMode()
    On Error Resume Next
    SavedCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub DisableFastMode()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```

- [ ] **Step 2: Write `src/vba/mod_FileDialog.bas`**

```vba
Attribute VB_Name = "mod_FileDialog"
Option Explicit

Public Function SelectExcelFile(ByVal DialogTitle As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = DialogTitle
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SelectExcelFile = .SelectedItems(1)
        Else
            SelectExcelFile = ""
        End If
    End With
End Function
```

- [ ] **Step 3: Commit Task 1**

```bash
git add src/vba/mod_Utils.bas src/vba/mod_FileDialog.bas
git commit -m "feat(vba): add mod_Utils and mod_FileDialog modules"
```

---

### Task 2: 코어 비교 엔진 모듈 작성 (`mod_DiffEngine.bas`)

**Files:**
- Create: `src/vba/mod_DiffEngine.bas`

**Interfaces:**
- Consumes: Target worksheet objects, file paths, options (AddComment, HighlightColor)
- Produces: `mod_DiffEngine.ProcessDiffAndMerge(BaseWbPath, ModWbPath, TargetSheetName, AddComment)`

- [ ] **Step 1: Write `src/vba/mod_DiffEngine.bas`**

```vba
Attribute VB_Name = "mod_DiffEngine"
Option Explicit

Public Function ProcessDiffAndMerge( _
    ByVal BasePath As String, _
    ByVal ModPath As String, _
    ByVal SheetName As String, _
    ByVal AddCommentOption As Boolean _
) As Long

    Dim wbBase As Workbook, wbMod As Workbook
    Dim wsBase As Worksheet, wsMod As Worksheet
    Dim rMax As Long, cMax As Long
    Dim r As Long, c As Long
    Dim baseArr As Variant, modArr As Variant
    Dim diffCount As Long
    
    diffCount = 0
    
    Set wbBase = Workbooks.Open(BasePath, ReadOnly:=False)
    Set wbMod = Workbooks.Open(ModPath, ReadOnly:=True)
    
    If SheetName = "" Or SheetName = "(첫번째 시트)" Then
        Set wsBase = wbBase.Worksheets(1)
        Set wsMod = wbMod.Worksheets(1)
    Else
        Set wsBase = wbBase.Worksheets(SheetName)
        Set wsMod = wbMod.Worksheets(SheetName)
    End If
    
    rMax = wsBase.UsedRange.Rows.Count
    cMax = wsBase.UsedRange.Columns.Count
    
    If rMax < 1 Or cMax < 1 Then Exit Function
    
    baseArr = wsBase.Range(wsBase.Cells(1, 1), wsBase.Cells(rMax, cMax)).Value
    modArr = wsMod.Range(wsMod.Cells(1, 1), wsMod.Cells(rMax, cMax)).Value
    
    For r = 1 To rMax
        For c = 1 To cMax
            Dim vBase As String, vMod As String
            vBase = CStr(baseArr(r, c))
            vMod = CStr(modArr(r, c))
            
            If vMod <> "" And vMod <> vBase Then
                wsBase.Cells(r, c).Value = modArr(r, c)
                wsBase.Cells(r, c).Interior.Color = RGB(255, 242, 204)
                
                If AddCommentOption Then
                    On Error Resume Next
                    wsBase.Cells(r, c).ClearComments
                    wsBase.Cells(r, c).AddComment "이전 값: " & vBase & vbCrLf & "수정일시: " & Format(Now, "yyyy-mm-dd hh:nn")
                    On Error GoTo 0
                End If
                diffCount = diffCount + 1
            End If
        Next c
    Next r
    
    wbMod.Close SaveChanges:=False
    wbBase.Save
    
    ProcessDiffAndMerge = diffCount
End Function
```

- [ ] **Step 2: Commit Task 2**

```bash
git add src/vba/mod_DiffEngine.bas
git commit -m "feat(vba): add mod_DiffEngine for high-speed array diffing and highlighting"
```

---

### Task 3: 대시보드 진입점 및 컨트롤러 (`mod_Main.bas`)

**Files:**
- Create: `src/vba/mod_Main.bas`

**Interfaces:**
- Consumes: Dashboard cell values (`B4`, `B5`, `B8`, etc.), `mod_FileDialog`, `mod_DiffEngine`, `mod_Utils`
- Produces: Macro buttons handlers (`SelectBaseFile`, `SelectModFile`, `RunAggregator`, `ResetHighlights`)

- [ ] **Step 1: Write `src/vba/mod_Main.bas`**

```vba
Attribute VB_Name = "mod_Main"
Option Explicit

Sub SelectBaseFile()
    Dim p As String
    p = mod_FileDialog.SelectExcelFile("원본(기준) 엑셀 파일 선택")
    If p <> "" Then
        ThisWorkbook.Sheets("대시보드").Range("C5").Value = p
    End If
End Sub

Sub SelectModFile()
    Dim p As String
    p = mod_FileDialog.SelectExcelFile("수정본 엑셀 파일 선택")
    If p <> "" Then
        ThisWorkbook.Sheets("대시보드").Range("C6").Value = p
    End If
End Sub

Sub RunAggregator()
    Dim wsDash As Worksheet
    Dim basePath As String, modPath As String
    Dim sheetName As String
    Dim addComment As Boolean
    Dim count As Long
    
    Set wsDash = ThisWorkbook.Sheets("대시보드")
    basePath = wsDash.Range("C5").Value
    modPath = wsDash.Range("C6").Value
    sheetName = wsDash.Range("C8").Value
    addComment = (UCase(wsDash.Range("C9").Value) = "예")
    
    If basePath = "" Or Dir(basePath) = "" Then
        MsgBox "올바른 원본 파일 경로를 선택해 주세요.", vbExclamation, "오류"
        Exit Sub
    End If
    
    If modPath = "" Or Dir(modPath) = "" Then
        MsgBox "올바른 수정본 파일 경로를 선택해 주세요.", vbExclamation, "오류"
        Exit Sub
    End If
    
    mod_Utils.EnableFastMode
    
    On Error GoTo ErrHandler
    count = mod_DiffEngine.ProcessDiffAndMerge(basePath, modPath, sheetName, addComment)
    
    mod_Utils.DisableFastMode
    MsgBox "취합이 성공적으로 완료되었습니다!" & vbCrLf & "총 " & count & "개 셀의 변경사항이 반영되었습니다.", vbInformation, "완료"
    Exit Sub

ErrHandler:
    mod_Utils.DisableFastMode
    MsgBox "취합 중 오류가 발생했습니다: " & Err.Description, vbCritical, "오류"
End Sub

Sub ResetHighlights()
    Dim p As String
    Dim wb As Workbook, ws As Worksheet
    p = ThisWorkbook.Sheets("대시보드").Range("C5").Value
    If p = "" Or Dir(p) = "" Then
        MsgBox "원본 파일 경로를 먼저 선택해주세요.", vbExclamation, "오류"
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(p)
    Set ws = wb.Worksheets(1)
    ws.UsedRange.Interior.ColorIndex = xlNone
    wb.Save
    wb.Close
    MsgBox "하이라이트 배경색이 초기화되었습니다.", vbInformation, "완료"
End Sub
```

- [ ] **Step 2: Commit Task 3**

```bash
git add src/vba/mod_Main.bas
git commit -m "feat(vba): add mod_Main dashboard control handlers"
```

---

### Task 4: 대시보드 워크북 자동 생성 스크립트 및 통합 시뮬레이션 테스트

**Files:**
- Create: `scripts/generate_workbook.py`
- Create: `tests/test_simulation.py`

**Interfaces:**
- Consumes: Python `openpyxl`
- Produces: `ExcelDiffAggregator.xlsx` / `.xlsm` file template and automated diff test suite

- [ ] **Step 1: Write `scripts/generate_workbook.py`**

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_dashboard():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "대시보드"
    ws.views.sheetView[0].showGridLines = True

    # Title Banner
    ws.merge_cells("B2:F3")
    title_cell = ws["B2"]
    title_cell.value = "📊 엑셀 변경사항 자동 취합기 (Excel Diff Aggregator)"
    title_cell.font = Font(name="맑은 고딕", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Section 1: File Paths
    ws["B4"] = "1. 파일 경로 지정"
    ws["B4"].font = Font(name="맑은 고딕", size=11, bold=True, color="1F4E78")

    ws["B5"] = "원본(기준) 파일"
    ws["B5"].font = Font(name="맑은 고딕", bold=True)
    ws["C5"] = ""

    ws["B6"] = "수정본 파일"
    ws["B6"].font = Font(name="맑은 고딕", bold=True)
    ws["C6"] = ""

    # Section 2: Options
    ws["B7"] = "2. 취합 옵션 설정"
    ws["B7"].font = Font(name="맑은 고딕", size=11, bold=True, color="1F4E78")

    ws["B8"] = "비교 시트명"
    ws["C8"] = "(첫번째 시트)"

    ws["B9"] = "이전 값 주석 남기기"
    ws["C9"] = "예"

    # Buttons guide
    ws["B11"] = "3. 매크로 실행 안내"
    ws["B11"].font = Font(name="맑은 고딕", size=11, bold=True, color="1F4E78")
    
    ws["B12"] = "• [원본 파일 선택] 매크로 실행: SelectBaseFile"
    ws["B13"] = "• [수정본 파일 선택] 매크로 실행: SelectModFile"
    ws["B14"] = "• [취합 실행] 매크로 실행: RunAggregator"
    ws["B15"] = "• [색상 초기화] 매크로 실행: ResetHighlights"

    # Adjust widths
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 50

    wb.save("ExcelDiffAggregator.xlsx")
    print("Dashboard Excel file generated successfully: ExcelDiffAggregator.xlsx")

if __name__ == "__main__":
    create_dashboard()
```

- [ ] **Step 2: Run Python Workbook Generator**

Run: `python scripts/generate_workbook.py`
Expected: Success message "Dashboard Excel file generated successfully"

- [ ] **Step 3: Write Test Simulation Script `tests/test_simulation.py`**

```python
import openpyxl

def test_excel_diff():
    # 1. Create sample Base file
    wb_base = openpyxl.Workbook()
    ws = wb_base.active
    ws.append(["ID", "이름", "부서", "직급", "연락처"])
    ws.append([101, "홍길동", "영업팀", "대리", "010-1111-2222"])
    ws.append([102, "김철수", "개발팀", "사원", "010-3333-4444"])
    ws.append([103, "이영희", "기획팀", "팀장", "010-5555-6666"])
    wb_base.save("sample_base.xlsx")

    # 2. Create sample Modified file (Kim Cheolsu team changed to 'AI팀', Lee Younghee position changed to '이사')
    wb_mod = openpyxl.Workbook()
    ws_m = wb_mod.active
    ws_m.append(["ID", "이름", "부서", "직급", "연락처"])
    ws_m.append([101, "홍길동", "영업팀", "대리", "010-1111-2222"])
    ws_m.append([102, "김철수", "AI팀", "사원", "010-3333-4444"])
    ws_m.append([103, "이영희", "기획팀", "이사", "010-5555-6666"])
    wb_mod.save("sample_mod.xlsx")

    print("Sample base and modified Excel files generated for testing.")

if __name__ == "__main__":
    test_excel_diff()
```

- [ ] **Step 4: Run test simulation data generator**

Run: `python tests/test_simulation.py`
Expected: Creates `sample_base.xlsx` and `sample_mod.xlsx`

- [ ] **Step 5: Commit Task 4**

```bash
git add scripts/generate_workbook.py tests/test_simulation.py
git commit -m "feat: add python automation scripts and sample test data generators"
```
