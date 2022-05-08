Attribute VB_Name = "Module6"
Sub Calcs()

'Calculate Yearly Changes

Dim c As Integer

For c = 2 To 3002

Cells(c, 12).Value = Cells(c, 11).Value - Cells(c + 1, 10).Value

If Cells(c, 12).Value > 0# Then
    Cells(c, 12).Interior.ColorIndex = 4
    
Else
    Cells(c, 12).Interior.ColorIndex = 3
    
End If

Next c


End Sub
