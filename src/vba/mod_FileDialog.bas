Attribute VB_Name = "mod_FileDialog"
Option Explicit

' 파일 선택 창(FileDialog)을 띄워 사용자로부터 엑셀 파일 경로 선택 받아오기
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
