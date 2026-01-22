Attribute VB_Name = "Aldos_Module"
Sub Aldo_Sort()
    '
    ' Aldo_Sort Macro
    ' Move below cutoff entries to bottom of list.
    '

    Dim ws As Worksheet
    Dim lastRow As Long

    ' Set the worksheet
    Set ws = ActiveWorkbook.Worksheets("Aldos")

    ' Find the last row with data in column A (or any column that you know will always have data)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Clear previous sort fields
    ws.Sort.SortFields.Clear

    ' Add sort fields with dynamic last row
    ws.Sort.SortFields.Add Key:=Range("D2:D" & lastRow), _
        SortOn:=xlSortOnCellColor, Order:=xlDescending, DataOption:=xlSortNormal
    ws.Sort.SortFields(ws.Sort.SortFields.Count).SortOnValue.Color = RGB(255, 199, 206)
    
    ws.Sort.SortFields.Add2 Key:=Range("C2:C" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add2 Key:=Range("B2:B" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add2 Key:=Range("D2:D" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add2 Key:=Range("E2:E" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    ' Apply the sort with dynamic range
    With ws.Sort
        .SetRange Range("A1:G" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Original_Sort()
'
' Original_Sort Macro
' Keep below cutoff values with there topics.
'

'
    ' Set the worksheet
    Set ws = ActiveWorkbook.Worksheets("Aldos")
    
    ' Find the last row with data in column A (or any column that you know will always have data)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 Key:=Range("C2:C" & lastRow) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add2 Key:=Range("B2:B" & lastRow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add(Range("D2:D" & lastRow), _
        xlSortOnCellColor, xlDescending, , xlSortNormal).SortOnValue.Color = RGB(255, _
        199, 206)
    ws.Sort.SortFields.Add2 Key:=Range("D2:D" & lastRow) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add2 Key:=Range("E2:E" & lastRow) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range("A1:G" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


