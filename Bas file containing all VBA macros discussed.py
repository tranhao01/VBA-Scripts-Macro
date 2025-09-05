# Create a .bas file containing all VBA macros discussed
content = r'''Attribute VB_Name = "AutomateAll"
Option Explicit

' ================= MAIN (IDs for Product/Discount/Units + Dimension build) =================
Public Sub BuildIDs_AllInOne_Styled()
    Dim fact As ListObject
    Set fact = Fact_Process_3Cols()                 ' overwrite 3 cột gốc bằng ID + đổi header
    If fact Is Nothing Then Exit Sub

    ' Lấy TableStyle của Fact để áp cho Dimension
    Dim styleName As String
    On Error Resume Next
    styleName = fact.TableStyle
    On Error GoTo 0
    If Len(styleName) = 0 Then styleName = "TableStyleMedium9" ' fallback

    ' Tạo/đổ dữ liệu đầy đủ cho toàn bộ dimension (distinct, có giá trị ở 2 cột)
    Dims_Process_All fact, styleName

    MsgBox "Done ✓ Fact_Finance đã đổi sang ProductID/DiscountID/UnitsID và toàn bộ Dimension (Segment, Country, Product, Discount, Units) đã được tạo/đổ dữ liệu distinct, style đồng bộ.", vbInformation
End Sub

' ============= FACT: ghi đè 3 cột gốc bằng ID + đổi header =============
Private Function Fact_Process_3Cols() As ListObject
    Dim ws As Worksheet, lo As ListObject
    Set ws = ThisWorkbook.Worksheets("Fact_Finance")
    If ws.ListObjects.Count = 0 Then Exit Function
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function

    ReplaceColWithID lo, "Product", "ProductID", 2, True      ' Text UPPER
    ReplaceColWithID lo, "Discount Band", "DiscountID", 2, True
    ReplaceColWithID lo, "Units Sold", "UnitsID", 3, False    ' Number (không warning)

    Set Fact_Process_3Cols = lo
End Function

Private Sub ReplaceColWithID(lo As ListObject, origName As String, newHeader As String, takeN As Long, asText As Boolean)
    Dim lc As ListColumn, i As Long, r As Long, n As Long
    Set lc = ColByName(lo, origName)
    If lc Is Nothing Then Exit Sub

    n = lc.DataBodyRange.Rows.Count
    If asText Then
        lc.DataBodyRange.NumberFormat = "@"
        For r = 1 To n
            lc.DataBodyRange(r, 1).Value = UCase$(Left$(CStr(lc.DataBodyRange(r, 1).Value), takeN))
        Next r
    Else
        lc.DataBodyRange.NumberFormat = "0"
        Dim s As String, digits As String, cutN As String
        For r = 1 To n
            s = CStr(lc.DataBodyRange(r, 1).Value)
            digits = DigitsOnly(s)
            cutN = Left$(digits, takeN)
            If cutN <> "" Then
                lc.DataBodyRange(r, 1).Value = CLng(cutN)   ' để thành số
            Else
                lc.DataBodyRange(r, 1).ClearContents
            End If
        Next r
    End If

    ' Đổi header
    lc.Name = newHeader

    ' Xóa cột ID trùng tên khác (nếu có)
    For i = lo.ListColumns.Count To 1 Step -1
        If lo.ListColumns(i).Name = newHeader And lo.ListColumns(i).Index <> lc.Index Then
            lo.ListColumns(i).Delete
        End If
    Next i
End Sub

' ============= DIMENSIONS: tạo/đổ dữ liệu DISTINCT 2 cột =============
Private Sub Dims_Process_All(fact As ListObject, styleName As String)
    ' Distinct từ Fact theo các ID (nếu cột nào không có thì bỏ qua)
    Dim idsSeg As Variant, idsCty As Variant, idsProd As Variant, idsDis As Variant, idsUnits As Variant
    idsSeg = DistinctFromFact(fact, "SegmentID")
    idsCty = DistinctFromFact(fact, "CountryID")
    idsProd = DistinctFromFact(fact, "ProductID")
    idsDis  = DistinctFromFact(fact, "DiscountID")
    idsUnits = DistinctFromFact(fact, "UnitsID")

    ' Build lookup tên hiện có (giữ tên cũ bạn đã nhập)
    Dim lkSeg As Object, lkCty As Object, lkProd As Object, lkDis As Object, lkUnits As Object
    Set lkSeg = ExistingLookup("Segment", 1, 2)
    Set lkCty = ExistingLookup("Country", 1, 2)
    Set lkProd = ExistingLookup("Product", 1, 2)
    Set lkDis = ExistingLookup("Discount", 1, 2)
    Set lkUnits = ExistingLookup("Units", 1, 2) ' với Units sẽ bị override bởi số

    ' Tạo/đảm bảo sheet + table và set style
    Dim loSeg As ListObject, loCty As ListObject, loProd As ListObject, loDis As ListObject, loU As ListObject
    Set loSeg = EnsureDimTable("Segment", "SegmentID", "Segment", styleName)
    Set loCty = EnsureDimTable("Country", "CountryID", "Country", styleName)
    Set loProd = EnsureDimTable("Product", "ProductID", "Product", styleName)
    Set loDis = EnsureDimTable("Discount", "DiscountID", "Discount", styleName)
    Set loU   = EnsureDimTable("Units",   "UnitsID",   "Units Sold", styleName)

    ' Đổ dữ liệu: Text = True, Number = False
    If Not IsEmpty(idsSeg) Then FillDimWithLookup loSeg, idsSeg, True, lkSeg
    If Not IsEmpty(idsCty) Then FillDimWithLookup loCty, idsCty, True, lkCty
    If Not IsEmpty(idsProd) Then FillDimWithLookup loProd, idsProd, True, lkProd
    If Not IsEmpty(idsDis) Then FillDimWithLookup loDis, idsDis, True, lkDis
    If Not IsEmpty(idsUnits) Then FillDimUnits loU, idsUnits       ' cả 2 cột đều là số
End Sub

' ---- Fill dimension (Text) giữ tên cũ nếu có; nếu trống → đặt = ID ----
Private Sub FillDimWithLookup(lo As ListObject, ids As Variant, asText As Boolean, nameLookup As Object)
    Dim r As Long, n As Long, idKey As String, nm As Variant

    ' Xóa data cũ (giữ header), resize theo số dòng mới
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    n = UBound(ids) - LBound(ids) + 1
    lo.Resize lo.Parent.Range("A1").Resize(n + 1, 2)

    lo.ListColumns(1).DataBodyRange.NumberFormat = "@"
    lo.ListColumns(2).DataBodyRange.NumberFormat = "@"   ' tên là text

    For r = 1 To n
        idKey = CStr(ids(r - 1))
        lo.DataBodyRange(r, 1).Value = idKey                 ' cột ID

        nm = vbNullString
        If Not nameLookup Is Nothing Then
            If nameLookup.Exists(idKey) Then nm = nameLookup(idKey)
        End If
        If Len(CStr(nm)) = 0 Then nm = idKey                 ' nếu chưa có tên, đặt = ID
        lo.DataBodyRange(r, 2).Value = nm                    ' cột Tên
    Next r

    ' Bỏ trùng theo cột ID nếu có duplicate
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.RemoveDuplicates Columns:=1, Header:=xlNo
End Sub

' ---- Fill Units: cả hai cột là số, distinct và có giá trị ----
Private Sub FillDimUnits(lo As ListObject, ids As Variant)
    Dim r As Long, n As Long
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    n = UBound(ids) - LBound(ids) + 1
    lo.Resize lo.Parent.Range("A1").Resize(n + 1, 2)

    lo.ListColumns(1).DataBodyRange.NumberFormat = "0"
    lo.ListColumns(2).DataBodyRange.NumberFormat = "0"

    For r = 1 To n
        lo.DataBodyRange(r, 1).Value = CLng(ids(r - 1))   ' UnitsID
        lo.DataBodyRange(r, 2).Value = CLng(ids(r - 1))   ' Units Sold (giá trị gốc)
    Next r

    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.RemoveDuplicates Columns:=1, Header:=xlNo
End Sub

' ---- Sync lại IDs từ cột tên cho toàn bộ dimension (nếu bạn chỉnh Name thủ công) ----
Public Sub RebuildDimIDs_FromNames_All()
    RebuildOneDim_Text "Segment", 3
    RebuildOneDim_Text "Country", 2
    RebuildOneDim_Text "Product", 2
    RebuildOneDim_Text "Discount", 2
    RebuildUnitsDim_Number "Units"
    MsgBox "Đã đồng bộ lại ID từ cột tên và loại trùng cho tất cả Dimension.", vbInformation
End Sub

Private Sub RebuildOneDim_Text(sheetName As String, takeN As Long)
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Or ws.ListObjects.Count = 0 Then Exit Sub

    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim n As Long, r As Long, nm As String
    n = lo.DataBodyRange.Rows.Count

    lo.ListColumns(1).DataBodyRange.NumberFormat = "@"
    For r = 1 To n
        nm = Trim$(CStr(lo.DataBodyRange(r, 2).Value))
        If Len(nm) > 0 Then lo.DataBodyRange(r, 1).Value = UCase$(Left$(nm, takeN))
    Next r
    ' Xoá hàng ID trống
    For r = n To 1 Step -1
        If Len(Trim$(CStr(lo.DataBodyRange(r, 1).Value))) = 0 Then lo.ListRows(r).Delete
    Next r
    ' Loại trùng
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.RemoveDuplicates Columns:=1, Header:=xlNo
End Sub

Private Sub RebuildUnitsDim_Number(sheetName As String)
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Or ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim n As Long, r As Long, v As Variant
    n = lo.DataBodyRange.Rows.Count

    lo.ListColumns(1).DataBodyRange.NumberFormat = "0"
    lo.ListColumns(2).DataBodyRange.NumberFormat = "0"

    For r = n To 1 Step -1
        If Len(Trim$(CStr(lo.DataBodyRange(r, 1).Value))) = 0 And Len(Trim$(CStr(lo.DataBodyRange(r, 2).Value))) = 0 Then
            lo.ListRows(r).Delete
        Else
            If Len(Trim$(CStr(lo.DataBodyRange(r, 1).Value))) > 0 Then
                v = CLng(lo.DataBodyRange(r, 1).Value)
                lo.DataBodyRange(r, 1).Value = v
                If Len(Trim$(CStr(lo.DataBodyRange(r, 2).Value))) = 0 Then
                    lo.DataBodyRange(r, 2).Value = v
                End If
            End If
        End If
    Next r

    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.RemoveDuplicates Columns:=1, Header:=xlNo
End Sub

' ---- Macro tiện dụng: sinh DiscountID từ Discount Band ở sheet Discount (rồi chuyển sang Value) ----
Public Sub DiscountID_FromName()
    Dim lo As ListObject, lc As ListColumn
    Set lo = Worksheets("Discount").ListObjects(1)
    Set lc = lo.ListColumns("DiscountID")
    lc.DataBodyRange.NumberFormat = "General"
    lc.DataBodyRange.Formula = "=UPPER(LEFT([@[Discount]],2))" ' nếu header là "Discount"
    On Error Resume Next
    lc.DataBodyRange.SpecialCells(xlCellTypeFormulas).Value = lc.DataBodyRange.SpecialCells(xlCellTypeFormulas).Value
    On Error GoTo 0
End Sub

' ============= DATE: build DateID + Date dimension =============
Public Sub BuildDate_ForFactAndDim()
    Dim fact As ListObject, styleName As String
    Set fact = GetFactTable("Fact_Finance")
    If fact Is Nothing Then
        MsgBox "Không tìm thấy Table ở sheet 'Fact_Finance'.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    styleName = fact.TableStyle
    On Error GoTo 0
    If Len(styleName) = 0 Then styleName = "TableStyleMedium9"

    Dim datesArr() As Date
    datesArr = DistinctDatesFromFact(fact)

    ReplaceDateWithDateID fact

    BuildDateDimension datesArr, styleName

    MsgBox "Done ✓ Đã thay Date→DateID trong Fact và dựng Dimension 'Date' đầy đủ (distinct).", vbInformation
End Sub

Private Function DistinctDatesFromFact(fact As ListObject) As Date()
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim iDateCol&, iDateIDCol&
    iDateCol = ColIndexByName(fact, "Date")
    iDateIDCol = ColIndexByName(fact, "DateID")

    Dim r As Long, n As Long
    n = fact.DataBodyRange.Rows.Count

    If iDateCol > 0 Then
        For r = 1 To n
            If IsDate(fact.DataBodyRange(r, iDateCol).Value) Then
                Dim d As Date: d = CDate(fact.DataBodyRange(r, iDateCol).Value)
                If Not dict.Exists(CLng(d)) Then dict.Add CLng(d), d
            End If
        Next r
    ElseIf iDateIDCol > 0 Then
        For r = 1 To n
            Dim id As Variant: id = fact.DataBodyRange(r, iDateIDCol).Value
            If IsNumeric(id) Then
                Dim y&, m&, d2&, dt As Date
                y = CLng(id) \ 10000
                m = (CLng(id) \ 100) Mod 100
                d2 = CLng(id) Mod 100
                If IsDate(DateSerial(y, m, d2)) Then
                    dt = DateSerial(y, m, d2)
                    If Not dict.Exists(CLng(dt)) Then dict.Add CLng(dt), dt
                End If
            End If
        Next r
    End If

    Dim arr() As Date, i As Long
    If dict.Count > 0 Then
        ReDim arr(0 To dict.Count - 1)
        For i = 0 To dict.Count - 1
            arr(i) = dict.Items()(i)
        Next i
        QuickSortDates arr, LBound(arr), UBound(arr)
    End If
    DistinctDatesFromFact = arr
End Function

Private Sub ReplaceDateWithDateID(fact As ListObject)
    Dim iDateCol&, r As Long, n As Long
    iDateCol = ColIndexByName(fact, "Date")
    If iDateCol = 0 Then Exit Sub

    n = fact.DataBodyRange.Rows.Count
    fact.ListColumns(iDateCol).DataBodyRange.NumberFormat = "0"  ' numeric
    For r = 1 To n
        If IsDate(fact.DataBodyRange(r, iDateCol).Value) Then
            Dim d As Date: d = fact.DataBodyRange(r, iDateCol).Value
            fact.DataBodyRange(r, iDateCol).Value = CDateID(d)
        Else
            fact.DataBodyRange(r, iDateCol).ClearContents
        End If
    Next r

    fact.ListColumns(iDateCol).Name = "DateID"
End Sub

Private Function CDateID(d As Date) As Long
    CDateID = Year(d) * 10000& + Month(d) * 100& + Day(d)
End Function

Private Sub BuildDateDimension(datesArr() As Date, styleName As String)
    Dim lo As ListObject: Set lo = EnsureDateTable(styleName)

    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    Dim n As Long, r As Long
    If (Not Not datesArr) = 0 Then Exit Sub
    n = UBound(datesArr) - LBound(datesArr) + 1

    lo.Resize lo.Parent.Range("A1").Resize(n + 1, 9)

    lo.ListColumns("DateID").DataBodyRange.NumberFormat = "0"
    lo.ListColumns("Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    lo.ListColumns("Year").DataBodyRange.NumberFormat = "0"
    lo.ListColumns("Quarter").DataBodyRange.NumberFormat = "General"
    lo.ListColumns("Month").DataBodyRange.NumberFormat = "0"
    lo.ListColumns("MonthName").DataBodyRange.NumberFormat = "@"
    lo.ListColumns("Day").DataBodyRange.NumberFormat = "0"
    lo.ListColumns("DayName").DataBodyRange.NumberFormat = "@"
    lo.ListColumns("WeekNum").DataBodyRange.NumberFormat = "0"

    For r = 1 To n
        Dim d As Date: d = datesArr(r - 1)
        lo.DataBodyRange(r, 1).Value = CDateID(d)
        lo.DataBodyRange(r, 2).Value = d
        lo.DataBodyRange(r, 3).Value = Year(d)
        lo.DataBodyRange(r, 4).Value = "Q" & ((Month(d) - 1) \ 3 + 1)
        lo.DataBodyRange(r, 5).Value = Month(d)
        lo.DataBodyRange(r, 6).Value = Format$(d, "mmmm")
        lo.DataBodyRange(r, 7).Value = Day(d)
        lo.DataBodyRange(r, 8).Value = Format$(d, "dddd")
        lo.DataBodyRange(r, 9).Value = WorksheetFunction.WeekNum(d, 2)
    Next r
End Sub

' ============= Units dimension theo range (bin) =============
Public Sub Rebuild_UnitsDim_FromFact()
    Dim fact As ListObject
    Set fact = GetFactTable("Fact_Finance")
    If fact Is Nothing Then
        MsgBox "Không tìm thấy Table ở 'Fact_Finance'.", vbCritical: Exit Sub
    End If

    Dim styleName As String
    On Error Resume Next
    styleName = fact.TableStyle
    On Error GoTo 0
    If Len(styleName) = 0 Then styleName = "TableStyleMedium9"

    Dim ids As Variant: ids = DistinctFromFact(fact, "UnitsID")
    If IsEmpty(ids) Then
        MsgBox "Không lấy được UnitsID từ Fact.", vbExclamation: Exit Sub
    End If

    Dim lo As ListObject: Set lo = EnsureUnitsTable(styleName)

    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    Dim n As Long: n = UBound(ids) - LBound(ids) + 1
    lo.Resize lo.Parent.Range("A1").Resize(n + 1, 2)

    lo.ListColumns(1).DataBodyRange.NumberFormat = "0"
    lo.ListColumns(2).DataBodyRange.NumberFormat = "@"

    Dim r As Long, id As Long, BIN As Long: BIN = 100
    For r = 1 To n
        id = CLng(ids(r - 1))
        lo.DataBodyRange(r, 1).Value = id
        lo.DataBodyRange(r, 2).Value = CStr(id) & "–" & CStr(id + (BIN - 1))
    Next r

    ' Sort theo UnitsID (data range, không header)
    lo.DataBodyRange.Sort Key1:=lo.ListColumns(1).DataBodyRange.Cells(1, 1), _
        Order1:=xlAscending, Header:=xlNo

    MsgBox "Đã dựng lại Units dimension: UnitsID | Units Range.", vbInformation
End Sub

' ============= Shared helpers =============
Private Function EnsureDateTable(styleName As String) As ListObject
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("Date")
    Dim lo As ListObject

    If ws.ListObjects.Count = 0 Then
        ws.Range("A1:I1").Value = Array("DateID", "Date", "Year", "Quarter", "Month", "MonthName", "Day", "DayName", "WeekNum")
        ws.Range("A2:I2").ClearContents
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:I2"), , xlYes)
        lo.Name = "Date_Table"
    Else
        Set lo = ws.ListObjects(1)
        On Error Resume Next
        lo.ListColumns(1).Name = "DateID"
        lo.ListColumns(2).Name = "Date"
        lo.ListColumns(3).Name = "Year"
        lo.ListColumns(4).Name = "Quarter"
        lo.ListColumns(5).Name = "Month"
        lo.ListColumns(6).Name = "MonthName"
        lo.ListColumns(7).Name = "Day"
        lo.ListColumns(8).Name = "DayName"
        lo.ListColumns(9).Name = "WeekNum"
        On Error GoTo 0
    End If

    On Error Resume Next
    lo.TableStyle = styleName
    On Error GoTo 0

    Set EnsureDateTable = lo
End Function

Private Function EnsureUnitsTable(styleName As String) As ListObject
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("Units")
    Dim lo As ListObject

    If ws.ListObjects.Count = 0 Then
        ws.Range("A1").Value = "UnitsID"
        ws.Range("B1").Value = "Units Range"
        ws.Range("A2:B2").ClearContents
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes)
        lo.Name = "Units_Table"
    Else
        Set lo = ws.ListObjects(1)
        On Error Resume Next
        lo.ListColumns(1).Name = "UnitsID"
        lo.ListColumns(2).Name = "Units Range"
        On Error GoTo 0
    End If

    On Error Resume Next
    lo.TableStyle = styleName
    On Error GoTo 0

    Set EnsureUnitsTable = lo
End Function

Private Function ExistingLookup(sheetName As String, idCol As Long, nameCol As Long) As Object
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Or ws.ListObjects.Count = 0 Then Exit Function

    Dim lo As ListObject: Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, idv As String, nm
    For r = 1 To lo.DataBodyRange.Rows.Count
        idv = CStr(lo.DataBodyRange(r, idCol).Value)
        nm = lo.DataBodyRange(r, nameCol).Value
        If Len(Trim$(idv)) > 0 And Len(Trim$(CStr(nm))) > 0 Then
            If Not dict.Exists(idv) Then dict.Add idv, nm
        End If
    Next r
    Set ExistingLookup = dict
End Function

Private Function DistinctFromFact(fact As ListObject, colName As String) As Variant
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = fact.ListColumns(colName)
    On Error GoTo 0
    If lc Is Nothing Then Exit Function
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, v
    For r = 1 To lc.DataBodyRange.Rows.Count
        v = lc.DataBodyRange(r, 1).Value
        If Len(Trim$(CStr(v))) > 0 Then
            If Not dict.Exists(CStr(v)) Then dict.Add CStr(v), v
        End If
    Next r
    If dict.Count = 0 Then Exit Function
    Dim arr() As Variant, i As Long
    ReDim arr(0 To dict.Count - 1)
    For i = 0 To dict.Count - 1
        arr(i) = dict.Items()(i)
    Next i
    DistinctFromFact = arr
End Function

Private Function GetFactTable(sheetName As String) As ListObject
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Or ws.ListObjects.Count = 0 Then Exit Function
    Set GetFactTable = ws.ListObjects(1)
End Function

Private Function GetOrCreateSheet(name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = name
    End If
End Function

Private Function ColByName(lo As ListObject, name As String) As ListColumn
    On Error Resume Next
    Set ColByName = lo.ListColumns(name)
    On Error GoTo 0
End Function

Private Function ColIndexByName(lo As ListObject, colName As String) As Long
    On Error Resume Next
    ColIndexByName = lo.ListColumns(colName).Index
    If Err.Number <> 0 Then ColIndexByName = 0
    Err.Clear
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

Private Sub QuickSortDates(a() As Date, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Date, t As Date
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortDates a, lo, j
    If i < hi Then QuickSortDates a, i, hi
End Sub
'''
path = '/mnt/data/Excel_Automate_All.bas'
with open(path, 'w', encoding='utf-8') as f:
    f.write(content)
print(path)
