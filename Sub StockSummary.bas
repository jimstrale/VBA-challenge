Attribute VB_Name = "Module4"
Sub StockSummary()

'Define variables
Dim Ticker As String
Dim Stock_Date As Date
Dim Stock_High As Double
Dim Stock_Low As Double
Dim Stock_Volume As Integer

'Set initial variable for holding the total volume per ticker

Dim Ticker_Total As Double
Ticker_Total = 0

'Set initial variable for holding the count per ticker

Dim Ticker_Count As Double
Ticker_Counter = 0

'Set location for each variable in new summary table

Dim Summary_Table_Volume As Integer
Summary_Table_Volume = 2

Dim Stock_Open As Double
Dim Stock_Close As Double
Dim Stock_Change As Double
Dim Summary_Table_Open As Double
Dim Summary_Table_Close As Double
Dim Summary_Table_Change As Double

'Determine the Last Row

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all the ticker names

For i = 2 To lastrow

'Check if we are still within the same Ticker name, if it is not...

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'Set ticker name
    
    Ticker = Cells(i, 1).Value
    
    'Add to the Total Stock Volume
    
    Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
    'Print Stock Open
    
    Range("J" & Summary_Table_Volume).Value = Stock_Open

    'Print the Ticker Name in the Table
    
    Range("I" & Summary_Table_Volume).Value = Ticker
    
    'Print the Volume in the Summary Table
    
    Range("N" & Summary_Table_Volume).Value = Ticker_Total
    
    'Add one to the table row
    
    Summary_Table_Volume = Summary_Table_Volume + 1
    
    'Set Open Value

    Stock_Open = Cells(i - Ticker_Count, 3).Value

    'Print Open Value

    Range("J" & Summary_Table_Volume).Value = Stock_Open
    
    'Reset Ticker Total
    
    Ticker_Total = 0
    Ticker_Count = 0
    
    'If the cell immediately following a row is the same ticker name...
    Else
    
'Add to the Ticker Total

Ticker_Total = Ticker_Total + Cells(i, 7).Value

'Add to Count Total

Ticker_Count = Ticker_Count + 1

'Set Close Value

Stock_Close = Cells(i + 1, 6).Value

'Print close value

Range("K" & Summary_Table_Volume).Value = Stock_Close


End If
 
Next i

End Sub


