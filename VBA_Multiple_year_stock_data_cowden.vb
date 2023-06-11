Sub stock()

'code for across all sheets in workbook
For Each ws In Worksheets
Dim WorksheetName As String
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
WorksheetName = ws.Name

'set column headings
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"

ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"


  'Declare variables
  Dim Ticker As String
  Dim yearly_change As Double
  yearly_change = 0
  Dim my_Total As Double
  my_Total = 0
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  Dim OpenValue As Double
  Dim CloseValue As Double
  Dim StockDifference As Double
  Dim PercentChange As Double
  Dim MaxpercentValue As Double
  Dim MinpercentValue As Double
  Dim MaxTotalVolume As Double
  
  
  For i = 2 To lastRow
   
    
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
     OpenValue = ws.Cells(i, 3).Value
     End If
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    CloseValue = ws.Cells(i, 6).Value
    
    StockDifference = CloseValue - OpenValue
    
    PercentChange = StockDifference / OpenValue
     
     
      
      Ticker = ws.Cells(i, 1).Value
      my_Total = my_Total + ws.Cells(i, 7).Value
      
      
      'print to summary table
      ws.Range("J" & Summary_Table_Row).Value = Ticker
      ws.Range("K" & Summary_Table_Row).Value = StockDifference
      ws.Range("L" & Summary_Table_Row).Value = PercentChange
      ws.Range("M" & Summary_Table_Row).Value = my_Total
      ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
      
      
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume Total
      my_Total = 0


    ' If the cell immediately following a row is the same stock...
    Else

      ' Add to the Total
     my_Total = my_Total + ws.Cells(i, 7).Value
    
    End If


  Next i
  
  
Next ws



'formatting summary table
For Each ws In Worksheets

 For i = 2 To lastRow
 If ws.Cells(i, 11).Value > 0 Then
 ws.Cells(i, 11).Interior.ColorIndex = 4
 ElseIf Cells(i, 11).Value < 0 Then
 ws.Cells(i, 11).Interior.ColorIndex = 3
 Else
 ws.Cells(i, 11).Interior.ColorIndex = 3
 End If
 Next i
'formatting summary table
For i = 2 To lastRow
 If ws.Cells(i, 12).Value > 0 Then
 ws.Cells(i, 12).Interior.ColorIndex = 4
 ElseIf Cells(i, 12).Value < 0 Then
 ws.Cells(i, 12).Interior.ColorIndex = 3
 Else
 ws.Cells(i, 12).Interior.ColorIndex = 3
 End If
 Next i
 
'Max percentage, min percentage, max total volume
MaxpercentValue = Application.WorksheetFunction.Max(ws.Range("L:L"))
ws.Cells(2, 18) = MaxpercentValue
ws.Cells(2, 18).NumberFormat = "0.00%"
MinpercentValue = Application.WorksheetFunction.Min(ws.Range("L:L"))
ws.Cells(3, 18) = MinpercentValue
ws.Cells(3, 18).NumberFormat = "0.00%"
MaxTotalVolume = Application.WorksheetFunction.Max(ws.Range("M:M"))
ws.Cells(4, 18) = MaxTotalVolume


'retrieves the ticker name based on cell values
For i = 2 To lastRow
'ticker name for greatest total
If ws.Cells(i, 13).Value = ws.Cells(4, 18).Value Then
ws.Cells(4, 17).Value = ws.Cells(i, 10).Value

End If
'ticker name for greatest decrease
If ws.Cells(i, 12).Value = ws.Cells(3, 18).Value Then
ws.Cells(3, 17).Value = ws.Cells(i, 10).Value

End If
'ticker name for greatest increase
If ws.Cells(i, 12).Value = ws.Cells(2, 18).Value Then
ws.Cells(2, 17).Value = ws.Cells(i, 10).Value

End If



Next i

Next ws











End Sub