'check_box_remove

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

