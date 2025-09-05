Attribute VB_Name = "AutomateAddons"
Option Explicit

' Convert the "Year" column on the "Date" sheet from text to number
Public Sub Convert_YearColumn_ToNumber()
    Convert_Column_ToNumber_InTable "Date", "Year"
End Sub

' Generic: convert any named column in the first ListObject of a sheet to Number
Public Sub Convert_Column_ToNumber_InTable(ByVal sheetName As String, ByVal colName As String)
    On Error GoTo Errh
    Dim ws As Worksheet, lo As ListObject, lc As ListColumn, rng As Range
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws.ListObjects.Count = 0 Then
        MsgBox "Sheet '" & sheetName & "' không có Table.", vbExclamation
        Exit Sub
    End If
    
    Set lo = ws.ListObjects(1)
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0
    If lc Is Nothing Or lc.DataBodyRange Is Nothing Then
        MsgBox "Không tìm thấy cột '" & colName & "' trong Table ở sheet '" & sheetName & "'.", vbExclamation
        Exit Sub
    End If
    
    Set rng = lc.DataBodyRange
    
    ' Chuẩn hoá định dạng số
    rng.NumberFormat = "0"
    
    ' Chuyển toàn bộ text số -> số (nhanh, không vòng lặp cell-by-cell)
    Dim f As String
    f = rng.Address(External:=True)
    ' Evaluate trả mảng cùng kích thước, với giá trị số = *1
    rng.Value = Evaluate("IF(" & f & "="""",""""," & f & "*1)")
    
    ' Gỡ cờ lỗi số dạng text (nếu còn)
    On Error Resume Next
    rng.Errors(xlNumberAsText).Ignore = True
    rng.Errors(xlNumberAsText).Ignore = False
    On Error GoTo 0
    
    MsgBox "Đã chuyển cột '" & colName & "' trên sheet '" & sheetName & "' sang Number.", vbInformation
    Exit Sub
Errh:
    MsgBox "Lỗi chuyển Number: " & Err.Description, vbExclamation
End Sub
