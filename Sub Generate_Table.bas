Attribute VB_Name = "Module2"
Sub Generate_Table()

'Add column titles

Range("I1").Value = "Ticker"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Percent Change"
Range("N1").Value = "Total Stock Volume"

Columns("J:K").Hidden = True


End Sub

