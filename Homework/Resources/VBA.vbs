Sub stockmarket():

Dim CompanyName As String
Dim CompanyFirstRow As Long
Dim CompanyLastRow As Long
Dim CompanyFirstValueOpen As Double
Dim CompanyLastValueClose As Double
Dim TickerRow As Long
Dim StockVol As Variant
Dim Active_Sheet As String
Dim LastRow As Long

'Determine the first row

Active_Sheet = "test"
For SheetCount = 1 To 3

If SheetCount = 1 Then Active_Sheet = "2014"
If SheetCount = 2 Then Active_Sheet = "2015"
If SheetCount = 3 Then Active_Sheet = "2016"
FirstRow = 2
TickerRow = 2
LastRow = Worksheets(Active_Sheet).Cells(Rows.Count, 1).End(xlUp).Row

'MsgBox (FirstRow)
'MsgBox (LastRow)


'create headers
Worksheets(Active_Sheet).Cells(1, "I") = "Ticker"
Worksheets(Active_Sheet).Cells(1, "J") = "Difference"
Worksheets(Active_Sheet).Cells(1, "K") = "% Difference"
Worksheets(Active_Sheet).Cells(1, "L") = "Total Volume"



'Determine the number of tickers
    For Counter1 = FirstRow To LastRow

    StockVol = StockVol + Worksheets(Active_Sheet).Cells(Counter1, "G")

    'First row of a company stuff
    If Worksheets(Active_Sheet).Cells(Counter1, "A") <> CompanyName Then
        
        'set the new compnay name
        CompanyName = Worksheets(Active_Sheet).Cells(Counter1, "A")
        
        'set the new first Row and value
        CompanyFirstRow = Counter1
        CompanyFirstValueOpen = Worksheets(Active_Sheet).Cells(Counter1, "C")
        
        'print company name for testing
        'Cells(Counter1, "H") = CompanyName
        
        'print the value for testing
        'Cells(Counter1, "I") = CompanyFirstValueOpen
        
        End If
    
    
    
    
    'last row of a compnay stuff
    If Worksheets(Active_Sheet).Cells(Counter1, "A") <> Worksheets(Active_Sheet).Cells(Counter1 + 1, "A") Then
        
        'set the new last Row and value
        CompanyLastRow = Counter1
        CompanyLastValueClose = Worksheets(Active_Sheet).Cells(Counter1, "F")
        
        'print the value for testing
        'Cells(Counter1, "J") = CompanyLastValueClose
        
        'set the company ticker
        Worksheets(Active_Sheet).Cells(TickerRow, "I") = CompanyName
        
        
        'set the difference and conditinal formatting
        Worksheets(Active_Sheet).Cells(TickerRow, "J") = CompanyLastValueClose - CompanyFirstValueOpen
        If CompanyLastValueClose - CompanyFirstValueOpen < 0 Then Worksheets(Active_Sheet).Cells(TickerRow, "J").Interior.ColorIndex = 3
        If CompanyLastValueClose - CompanyFirstValueOpen > 0 Then Worksheets(Active_Sheet).Cells(TickerRow, "J").Interior.ColorIndex = 4
        'set the difference percentage (added if because there are some 0 values
        If CompanyLastValueClose - CompanyFirstValueOpen <> 0 And CompanyFirstValueOpen <> 0 Then Worksheets(Active_Sheet).Cells(TickerRow, "K") = (CompanyLastValueClose - CompanyFirstValueOpen) / CompanyFirstValueOpen
        Worksheets(Active_Sheet).Cells(TickerRow, "K").NumberFormat = "0.00%"
        'set the stock vol
        Worksheets(Active_Sheet).Cells(TickerRow, "L") = StockVol
        'reset the stock vol
        StockVol = 0
        'increment the ticker
         TickerRow = TickerRow + 1
        
         End If
        
      
    Next Counter1
Next SheetCount
'MsgBox (Company_Count)



'grab the first row
'grab the last row


End Sub



