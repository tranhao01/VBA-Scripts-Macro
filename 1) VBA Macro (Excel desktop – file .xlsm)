Option Explicit

Private Function EnsureColumn(lo As ListObject, colName As String) As ListColumn
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0
    If lc Is Nothing Then
        Set EnsureColumn = lo.ListColumns.Add
        EnsureColumn.Name = colName
    Else
        Set EnsureColumn = lc
    End If
End Function

Public Sub BuildIDs()
    Dim sep As String: sep = Application.International(xlListSeparator) ' "," hoặc ";"
    Dim ws As Worksheet, lo As ListObject

    '=== FACT_FINANCE: thêm 3 cột ID ===
    Set ws = ThisWorkbook.Worksheets("Fact_Finance")
    If ws.ListObjects.Count = 0 Then
        MsgBox "Sheet 'Fact_Finance' chưa có Table.", vbCritical: Exit Sub
    End If
    Set lo = ws.ListObjects(1)

    Dim colSeg As ListColumn, colCty As ListColumn, colProd As ListColumn

    Set colSeg = EnsureColumn(lo, "SegmentID")
    Set colCty = EnsureColumn(lo, "CountryID")
    Set colProd = EnsureColumn(lo, "ProductID")

    ' Công thức: dùng structured refs & UPPER(LEFT(...))
    colSeg.DataBodyRange.Formula = "=UPPER(LEFT([@Segment]" & sep & "3))"
    colCty.DataBodyRange.Formula = "=UPPER(LEFT([@Country]" & sep & "2))"
    colProd.DataBodyRange.Formula = "=UPPER(LEFT([@Product]" & sep & "2))"

    '=== PRODUCT: tạo ProductID nếu có bảng ===
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Product")
    On Error GoTo 0
    If Not ws Is Nothing Then
        If ws.ListObjects.Count > 0 Then
            Set lo = ws.ListObjects(1)
            Dim colPID As ListColumn
            Set colPID = EnsureColumn(lo, "ProductID")
            colPID.DataBodyRange.Formula = "=UPPER(LEFT([@Product]" & sep & "2))"
        End If
    End If

    MsgBox "Đã tạo/cập nhật SegmentID, CountryID, ProductID.", vbInformation
End Sub
