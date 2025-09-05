Public Sub TransformFactFinance()
    Dim sep As String: sep = Application.International(xlListSeparator)
    Dim ws As Worksheet, lo As ListObject
    
    Set ws = ThisWorkbook.Worksheets("Fact_Finance")
    If ws.ListObjects.Count = 0 Then
        MsgBox "Sheet 'Fact_Finance' chưa có Table.", vbCritical: Exit Sub
    End If
    Set lo = ws.ListObjects(1)
    
    '--- Tạo ID cột ---
    Dim colSeg As ListColumn, colCty As ListColumn
    Dim colProd As ListColumn, colDis As ListColumn, colUnit As ListColumn
    
    Set colSeg = EnsureColumn(lo, "SegmentID")
    colSeg.DataBodyRange.Formula = "=UPPER(LEFT([@Segment]" & sep & "3))"
    
    Set colCty = EnsureColumn(lo, "CountryID")
    colCty.DataBodyRange.Formula = "=UPPER(LEFT([@Country]" & sep & "2))"
    
    Set colProd = EnsureColumn(lo, "ProductID")
    colProd.DataBodyRange.Formula = "=UPPER(LEFT([@Product]" & sep & "2))"
    
    Set colDis = EnsureColumn(lo, "DiscountID")
    colDis.DataBodyRange.Formula = "=UPPER(LEFT([@[Discount Band]]" & sep & "2))"
    
    Set colUnit = EnsureColumn(lo, "UnitsID")
    colUnit.DataBodyRange.Formula = "=UPPER(LEFT(TEXT([@[Units Sold]],""0"")" & sep & "3))"
    
    '--- Chuyển công thức thành giá trị ---
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If lc.Name Like "*ID" Then
            lc.DataBodyRange.Value = lc.DataBodyRange.Value
        End If
    Next lc
    
    '--- Xoá cột gốc ---
    On Error Resume Next
    lo.ListColumns("Segment").Delete
    lo.ListColumns("Country").Delete
    lo.ListColumns("Product").Delete
    lo.ListColumns("Discount Band").Delete
    lo.ListColumns("Units Sold").Delete
    On Error GoTo 0
    
    MsgBox "Đã tạo ID và xoá cột gốc, giữ lại cột ID.", vbInformation
End Sub

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
