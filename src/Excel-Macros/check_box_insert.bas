Attribute VB_Name = "Module1"
'check_box_insert

Sub InsertCheckBoxes()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim chkBox As CheckBox
    Dim cellC As Range
    Dim cellD As Range
    Dim lastRow As Long
    
    ' Determinar la última fila con datos en la columna C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Eliminar checkboxes existentes
    For Each chkBox In ws.CheckBoxes
        chkBox.Delete
    Next chkBox
    
    ' Agregar checkboxes en la columna D para cada valor en la columna C
    For i = 2 To lastRow
        Set cellC = ws.Cells(i, "C")
        Set cellD = ws.Cells(i, "D")
        If Not IsEmpty(cellC.Value) Then
            Set chkBox = ws.CheckBoxes.Add(cellD.Left, cellD.Top, cellD.Width, cellD.Height)
            chkBox.Caption = "Excel"
            chkBox.LinkedCell = cellC.Address  ' Vincular el checkbox a la celda correspondiente en la columna C
        End If
    Next i
End Sub

