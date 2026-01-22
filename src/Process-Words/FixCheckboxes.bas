Attribute VB_Name = "Module2"
Sub AdjustCheckBoxes()
    Dim rng As Range
    Dim cell As Range
    Dim cb As CheckBox
    
    ' Define the range where your checkboxes are located
    Set rng = Range("D2:D91") ' Change this range as needed
    
    For Each cell In rng
        ' Check if the cell contains a checkbox
        For Each cb In ActiveSheet.CheckBoxes
            If Not Intersect(cb.TopLeftCell, cell) Is Nothing Then
                ' Adjust the cell link property based on the checkbox's position
                cb.LinkedCell = cell.Offset(0, -1).Address
            End If
        Next cb
    Next cell
End Sub

