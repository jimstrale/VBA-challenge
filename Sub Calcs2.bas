Attribute VB_Name = "Module5"
Sub Calcs2()

'Calculate Percent Change


Dim e As Integer

Columns("M").NumberFormat = "0.00%"

For e = 2 To 3002

Cells(e, 13).Value = (Cells(e, 12).Value / Cells(e + 1, 10).Value)


Next e


End Sub
