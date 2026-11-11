Attribute VB_Name = "mod_Utils"
Option Explicit

Private SavedCalcMode As Long

' 엑셀 화면 갱신 및 자동 계산 일시 정지 (초고속 속도 최적화)
Public Sub EnableFastMode()
    On Error Resume Next
    SavedCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

' 엑셀 설정 화면 갱신 및 자동 계산 복구
Public Sub DisableFastMode()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' 파일 존재 여부 안전 검증 (UNC 경로 및 특수문자 대응)
Public Function FileExists(ByVal FilePath As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(FilePath)
    On Error GoTo 0
End Function

' 절대 경로 변환 및 열려 있는 워크북 재사용 처리 래퍼 함수 (Run-time Error 1004 완전 방지)
Public Function GetOrOpenWorkbook(ByVal FilePath As String, ByVal IsReadOnly As Boolean) As Workbook
    Dim fso As Object
    Dim absPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    
    If Trim(FilePath) = "" Then Exit Function
    
    ' 셀 편집 모드 탈출 시도
    On Error Resume Next
    DoEvents
    On Error GoTo 0
    
    ' 1. FileSystemObject를 사용해 상대 경로를 절대 경로로 안전하게 변환
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    absPath = fso.GetAbsolutePathName(FilePath)
    fileName = fso.GetFileName(absPath)
    On Error GoTo 0
    
    If absPath = "" Then absPath = FilePath
    If fileName = "" Then fileName = FilePath
    
    ' 2. 파일 존재 여부 최종 확인
    If Not FileExists(absPath) Then
        Err.Raise vbObjectError + 1003, "GetOrOpenWorkbook", "지정한 엑셀 파일을 찾을 수 없습니다:" & vbCrLf & absPath
        Exit Function
    End If
    
    ' 3. 이미 Excel 내에 열려 있는 워크북인지 먼저 확인
    On Error Resume Next
    Set wb = Application.Workbooks(fileName)
    On Error GoTo 0
    
    ' 4. 열려 있지 않은 경우 Open 호출
    ' ★ 핵심: Open 시에는 DisplayAlerts=True 상태여야 보안 경고/제한된 보기/파일 변환 알림으로 인한 1004 에러를 방지함
    If wb Is Nothing Then
        prevDisplayAlerts = Application.DisplayAlerts
        prevEnableEvents = Application.EnableEvents
        
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        
        On Error Resume Next
        Set wb = Application.Workbooks.Open( _
            FileName:=absPath, _
            UpdateLinks:=0, _
            ReadOnly:=IsReadOnly, _
            IgnoreReadOnlyRecommended:=True _
        )
        If wb Is Nothing Then
            Set wb = Application.Workbooks.Open(absPath)
        End If
        Err.Clear
        On Error GoTo 0
        
        ' 이전 알림 설정 복원
        Application.DisplayAlerts = prevDisplayAlerts
        Application.EnableEvents = prevEnableEvents
    End If
    
    ' 5. 여전히 열지 못한 경우 (파일 잠김, 보호된 시트, 암호 등)
    If wb Is Nothing Then
        Err.Raise vbObjectError + 1002, "GetOrOpenWorkbook", _
            "엑셀 파일(" & fileName & ")을 열 수 없습니다." & vbCrLf & vbCrLf & _
            "■ 원인 및 해결 방법:" & vbCrLf & _
            "1. 파일이 다른 사용자나 다른 프로그램에서 사용 중인지 확인하세요." & vbCrLf & _
            "2. 탐색기에서 파일 속성 -> [차단 해제] 체크 여부를 확인하세요." & vbCrLf & _
            "3. 엑셀 셀 편집 중(커서가 셀 안에 있음)이라면 Enter를 눌러 편집을 마친 후 실행하세요."
    End If
    
    Set GetOrOpenWorkbook = wb
End Function
