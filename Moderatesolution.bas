Attribute VB_Name = "Module1"
Sub tickercalculation():
'
'Easy
'
'Form table of stock values on each worksheet
'Compiling total values for fiscal year
'Restarting the count for each new symbol
'Sum of stock volume from first day to last day = total volume
'
    Dim Col  As Double
    Dim Total_Volume As Double
    Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"
    Col = 2
    Cells(Col, 9).Value = Cells(Col, 1).Value
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For Row = 2 To LastRow
    If Cells(Row, 1).Value = Cells(Col, 9) Then
    
     Total_Volume = Total_Volume + Cells(Row, 7).Value
     Else
     Cells(Col, 10).Value = Total_Volume
     Total_Volume = Cells(Row, 7).Value
     Col = Col + 1
     Cells(Col, 9).Value = Cells(Row, 1).Value
     End If
     
     Next Row
     
     Cells(Col, 10).Value = Total_Volume
     
     Next WS
'Medium
'
'Dollar change subtract open first day of year from close last day of year
'Close last day less open first day of year
'Percentage change equals dollar change divided by open first day of year
'Close last day less open first day of year
'
'Hard
'
'Return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:
'Run on every worksheet just by running the VBA script once.
'


End Sub
