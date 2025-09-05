Attribute VB_Name = "AutomateAddons"
Option Explicit

' Convert the "Year" column on the "Date" sheet from text to number
Public Sub Convert_YearColumn_ToNumber()
    Convert_Column_ToNumber_InTable "Date", "Year"
End Sub

' Generic: convert any named column in the first ListObject of a sheet to Number (robust, no short-circuit issues)
Public Sub Convert_Column_ToNumber_InTable(ByVal sheetName As String, ByVal colName As String)
    On Error GoTo Fail
    Dim ws As Worksheet, lo As ListObject, lc As ListColumn, rng As Range

    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws.ListObjects.Count = 0 Then
        MsgBox "Sheet '" & sheetName & "' không có Table.", vbExclamation
        Exit Sub
    End If

    Set lo = ws.ListObjects(1)

    ' Tìm cột theo tên
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0

    If lc Is Nothing Then
        MsgBox "Không tìm thấy cột '" & colName & "' trong Table ở sheet '" & sheetName & "'.", vbExclamation
        Exit Sub
    End If

    ' Nếu table chưa có dòng dữ liệu
    If lo.DataBodyRange Is Nothing Then
        MsgBox "Table ở sheet '" & sheetName & "' chưa có dữ liệu để chuyển.", vbExclamation
        Exit Sub
    End If

    ' Lấy phần giao giữa cột và phần thân dữ liệu (tránh header)
    Set rng = Intersect(lc.Range, lo.DataBodyRange)
    If rng Is Nothing Then
        MsgBox "Không có ô dữ liệu cho cột '" & colName & "'.", vbExclamation
        Exit Sub
    End If

    ' Chuẩn hoá định dạng và chuyển text -> number hàng loạt
    rng.NumberFormat = "0"
    Dim addr As String
    addr = rng.Address(External:=True)
    rng.Value = Evaluate("IF(" & addr & "="""",""""," & addr & "*1)")

    ' Clear green triangles nếu còn
    On Error Resume Next
    rng.Errors(xlNumberAsText).Ignore = True
    rng.Errors(xlNumberAsText).Ignore = False
    On Error GoTo 0

    MsgBox "Đã chuyển cột '" & colName & "' trên sheet '" & sheetName & "' sang Number.", vbInformation
    Exit Sub

Fail:
    MsgBox "Lỗi chuyển Number: " & Err.Description, vbExclamation
End Sub
