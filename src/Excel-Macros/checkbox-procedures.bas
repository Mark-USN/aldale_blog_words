Attribute VB_Name = "Module1"
'check_box_procedures

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
            ' 20240503 MMH added checkbox title
            chkBox.Caption = "Excelude"
            chkBox.LinkedCell = cellC.Address  ' Vincular el checkbox a la celda correspondiente en la columna C
        End If
    Next i
End Sub

Sub RemoveCheckBoxesFromD()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim chkBox As CheckBox
    Dim cell As Range
    Dim targetColumn As Integer
    targetColumn = 4  ' Columna D es la 4ta columna (A=1, B=2, C=3, D=4)
    
    ' Recorrer todos los checkboxes en la hoja
    For Each chkBox In ws.CheckBoxes
        ' Determinar si el checkbox está en la columna D
        Set cell = ws.Cells(chkBox.TopLeftCell.Row, targetColumn)
        If Not Intersect(chkBox.TopLeftCell, cell) Is Nothing Then
            chkBox.Delete
        End If
    Next chkBox
End Sub


