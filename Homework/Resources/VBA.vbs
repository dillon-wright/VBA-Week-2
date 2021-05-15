Sub stockmarket():

Dim CompanyName As String
Dim CompanyFirstRow As Long
Dim CompanyLastRow As Long
Dim CompanyFirstValueOpen As Double
Dim CompanyLastValueClose As Double
Dim TickerRow As Long
Dim StockVol As Long

'Determine the first row

FirstRow = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
TickerRow = 2

'MsgBox (FirstRow)
'MsgBox (LastRow)



'Determine the number of tickers
    For Counter1 = FirstRow To LastRow

    StockVol = StockVol + Cells(Counter1, "G")

    'First row of a company stuff
    If Cells(Counter1, "A") <> CompanyName Then
        
        'set the new compnay name
        CompanyName = Cells(Counter1, "A")
        
        'set the new first Row and value
        CompanyFirstRow = Counter1
        CompanyFirstValueOpen = Cells(Counter1, "C")
        
        'print company name for testing
        'Cells(Counter1, "H") = CompanyName
        
        'print the value for testing
        'Cells(Counter1, "I") = CompanyFirstValueOpen
        
        End If
    
    
    
    
    'last row of a compnay stuff
    If Cells(Counter1, "A") <> Cells(Counter1 + 1, "A") Then
        
        'set the new last Row and value
        CompanyLastRow = Counter1
        CompanyLastValueClose = Cells(Counter1, "F")
        
        'print the value for testing
        'Cells(Counter1, "J") = CompanyLastValueClose
        
        'set the company ticker
        Cells(TickerRow, "I") = CompanyName
        
        
        'set the difference
        Cells(TickerRow, "J") = CompanyLastValueClose - CompanyFirstValueOpen
        'set the difference percentage
        Cells(TickerRow, "K") = (CompanyLastValueClose - CompanyFirstValueOpen) / CompanyFirstValueOpen
        'set the stock vol
        Cells(TickerRow, "L") = StockVol
        'reset the stock vol
        StockVol = 0
        'increment the ticker
         TickerRow = TickerRow + 1
        
         End If
        
      
    Next Counter1

'MsgBox (Company_Count)



'grab the first row
'grab the last row


End Sub



