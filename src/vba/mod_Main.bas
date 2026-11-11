Attribute VB_Name = "mod_Main"
Option Explicit

' 1. 원본 파일 선택 버튼 매크로
Sub SelectBaseFile()
    Dim filePath As String
    filePath = mod_FileDialog.SelectExcelFile("원본(기준) 엑셀 파일 선택")
    If filePath <> "" Then
        ThisWorkbook.Sheets("대시보드").Range("C5").Value = filePath
    End If
End Sub

' 2. 수정본 파일 선택 버튼 매크로
Sub SelectModFile()
    Dim filePath As String
    filePath = mod_FileDialog.SelectExcelFile("수정본 엑셀 파일 선택")
    If filePath <> "" Then
        ThisWorkbook.Sheets("대시보드").Range("C6").Value = filePath
    End If
End Sub

' 3. 취합 및 하이라이트 실행 버튼 매크로
Sub RunAggregator()
    Dim wsDash As Worksheet
    Dim basePath As String, modPath As String
    Dim sheetName As String
    Dim addComment As Boolean
    Dim changedCount As Long
    
    Set wsDash = ThisWorkbook.Sheets("대시보드")
    basePath = Trim(wsDash.Range("C5").Value)
    modPath = Trim(wsDash.Range("C6").Value)
    sheetName = Trim(wsDash.Range("C8").Value)
    addComment = (UCase(Trim(wsDash.Range("C9").Value)) = "예")
    
    ' 경로 검증 (FSO 기반 안전 검증)
    If basePath = "" Or Not mod_Utils.FileExists(basePath) Then
        MsgBox "올바른 원본(기준) 파일 경로를 선택해 주세요." & vbCrLf & "입력 경로: " & basePath, vbExclamation, "경로 오류"
        Exit Sub
    End If
    
    If modPath = "" Or Not mod_Utils.FileExists(modPath) Then
        MsgBox "올바른 수정본 파일 경로를 선택해 주세요." & vbCrLf & "입력 경로: " & modPath, vbExclamation, "경로 오류"
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    changedCount = mod_DiffEngine.ProcessDiffAndMerge(basePath, modPath, sheetName, addComment)
    
    ' 속도 최적화 해제
    mod_Utils.DisableFastMode
    
    MsgBox "취합이 성공적으로 완료되었습니다!" & vbCrLf & _
           "총 " & changedCount & "개 셀의 변경사항이 원본에 반영 및 하이라이트 처리되었습니다.", _
           vbInformation, "취합 완료"
    Exit Sub

ErrHandler:
    mod_Utils.DisableFastMode
    MsgBox "취합 처리 중 오류가 발생했습니다:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "실행 오류"
End Sub

' 4. 하이라이트 배경색 초기화 버튼 매크로
Sub ResetHighlights()
    Dim basePath As String
    Dim wb As Workbook, ws As Worksheet
    
    basePath = Trim(ThisWorkbook.Sheets("대시보드").Range("C5").Value)
    If basePath = "" Or Not mod_Utils.FileExists(basePath) Then
        MsgBox "원본 파일 경로를 먼저 선택해주세요.", vbExclamation, "경로 오류"
        Exit Sub
    End If
    
    On Error GoTo ErrReset
    Set wb = mod_Utils.GetOrOpenWorkbook(basePath, IsReadOnly:=False)
    
    ' 워크북 오픈 후 속도 최적화 활성화
    mod_Utils.EnableFastMode
    
    Set ws = wb.Worksheets(1)
    ws.UsedRange.Interior.ColorIndex = xlNone
    wb.Save
    If Not (wb Is ThisWorkbook) Then wb.Close
    
    mod_Utils.DisableFastMode
    MsgBox "원본 파일의 하이라이트 배경색이 성공적으로 초기화되었습니다.", vbInformation, "초기화 완료"
    Exit Sub

ErrReset:
    mod_Utils.DisableFastMode
    MsgBox "초기화 중 오류가 발생했습니다:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "오류"
End Sub
