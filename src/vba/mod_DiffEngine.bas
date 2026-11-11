Attribute VB_Name = "mod_DiffEngine"
Option Explicit

' 원본과 수정본 워크북을 로드하여 메모리 배열 상에서 1:1 비교 후 변경 셀 반영 및 하이라이트
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
    
    ' 1. 워크북 안전 열기 (DisplayAlerts=True 상태 보장)
    Set wbBase = mod_Utils.GetOrOpenWorkbook(BasePath, IsReadOnly:=False)
    Set wbMod = mod_Utils.GetOrOpenWorkbook(ModPath, IsReadOnly:=True)
    
    ' 2. 워크북 열기 완료 후 초고속 메모리 비교를 위해 속도 최적화 활성화
    mod_Utils.EnableFastMode
    
    Dim isAutoAllSheets As Boolean, isFirstSheetOnly As Boolean
    isAutoAllSheets = (SheetName = "" Or SheetName = "(전체 시트 자동 인식)" Or SheetName = "(전체 시트)" Or SheetName = "(모든 시트)")
    isFirstSheetOnly = (SheetName = "(첫번째 시트)")
    
    Dim wsItem As Worksheet
    Dim sheetList As Collection
    Set sheetList = New Collection
    
    If isAutoAllSheets Then
        For Each wsItem In wbBase.Worksheets
            sheetList.Add wsItem.Name
        Next wsItem
    ElseIf isFirstSheetOnly Then
        sheetList.Add wbBase.Worksheets(1).Name
    Else
        sheetList.Add SheetName
    End If
    
    Dim sIndex As Long
    For sIndex = 1 To sheetList.Count
        Dim currentSheetName As String
        currentSheetName = sheetList(sIndex)
        
        Set wsBase = Nothing
        Set wsMod = Nothing
        
        On Error Resume Next
        Set wsBase = wbBase.Worksheets(currentSheetName)
        Set wsMod = wbMod.Worksheets(currentSheetName)
        On Error GoTo 0
        
        If Not wsBase Is Nothing And Not wsMod Is Nothing Then
            rMax = wsBase.UsedRange.Rows.Count
            cMax = wsBase.UsedRange.Columns.Count
            
            If rMax >= 1 And cMax >= 1 Then
                baseArr = wsBase.Range(wsBase.Cells(1, 1), wsBase.Cells(rMax, cMax)).Value
                modArr = wsMod.Range(wsMod.Cells(1, 1), wsMod.Cells(rMax, cMax)).Value
                
                For r = 1 To rMax
                    For c = 1 To cMax
                        Dim vBase As String, vMod As String
                        
                        If IsError(baseArr(r, c)) Then
                            vBase = "#ERROR"
                        Else
                            vBase = CStr(baseArr(r, c))
                        End If
                        
                        If IsError(modArr(r, c)) Then
                            vMod = "#ERROR"
                        Else
                            vMod = CStr(modArr(r, c))
                        End If
                        
                        If Trim(vMod) <> "" And vMod <> vBase Then
                            wsBase.Cells(r, c).Value = modArr(r, c)
                            wsBase.Cells(r, c).Interior.Color = RGB(255, 242, 204)
                            
                            If AddCommentOption Then
                                On Error Resume Next
                                wsBase.Cells(r, c).ClearComments
                                wsBase.Cells(r, c).AddComment "기존 값: " & vBase & vbCrLf & "수정일시: " & Format(Now, "yyyy-mm-dd hh:nn")
                                On Error GoTo 0
                            End If
                            
                            diffCount = diffCount + 1
                        End If
                    Next c
                Next r
            End If
        End If
    Next sIndex
    
    ' 파일 저장 및 닫기
    wbMod.Close SaveChanges:=False
    wbBase.Save
    If Not (wbBase Is ThisWorkbook) Then wbBase.Close SaveChanges:=False
    
    ProcessDiffAndMerge = diffCount
End Function
