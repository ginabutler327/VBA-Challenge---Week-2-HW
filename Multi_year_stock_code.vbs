Sub stocks_per_qtr()

'Stocks are presorted

   Dim ticker As String
   Dim total As Long
   Dim i As Long
   Dim j As Integer
   Dim qtr_change As Double
   Dim percent_chg As Double
   Dim dailyChange As Double
   Dim AvgChange As Double
   Dim total_stock_vol As Double
   Dim next_vol As Long
   Dim start As Long
   Dim RowCount As Long
   Dim GreatestIncrease As Double
   Dim GreatestDecrease As Double
   Dim GreatestIncreaseTicker As String
   Dim GreatestDecreaseTicker As String
   Dim GreatestTotalVolume As Double
   Dim GreatestVolumeTicker As String
   Dim ws As Worksheet
   
For Each ws In ThisWorkbook.Worksheets
'set title row
ws.Range("I1").Value = "Ticker"
ws.Range("j1").Value = "Quarterly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volue"
ws.Range("s1").Value = "Value"
ws.Range("q2").Value = "Greatest % Increase"
ws.Range("Q3").Value = "Greatest % Decrease"
ws.Range("Q4").Value = "Greatest Total Volume"
ws.Range("r1").Value = "Ticker"

'set initial values
    total_stock_vol = 0
    GreatestIncrease = -1
    GreatestDecrease = 1
    GreatestTotalVolume = 0
'get row number of last row with data
  RowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
 
'set initial value
  j = 0
  total_stock_vol = 0
      
  For i = 2 To RowCount
     'Loop through column A, starting at row 2
       ticker = ws.Cells(i, 1).Value
       total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
              
   'if ticker changes then print results, use range!
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      'Calculate quarterly change and percent change
       qtr_change = ws.Cells(i, 3).Value - ws.Cells(i, 6).Value
       If ws.Cells(i, 3).Value <> 0 Then
          percent_chg = (qtr_change / ws.Cells(i, 3).Value) * 100
       Else
          percent_chg = 0
       End If
       If percent_chg > GreatestIncrease Then
       GreatestIncrease = percent_chg
       GreatestIncreaseTicker = ticker
       End If
       If percent_chg < GreatestDecrease Then
       GreatestDecrease = percent_chg
       GreatestDecreaseTicker = ticker
       End If
       If total_stock_vol > GreatestTotalVolume Then
       GreatestTotalVolume = total_stock_vol
       GreatestVolumeTicker = ticker
       End If
       'Output results
       ws.Cells(j + 2, 9).Value = ticker
       ws.Cells(j + 2, 10).Value = qtr_change
       ws.Cells(j + 2, 11).Value = FormatPercent(percent_chg)
       ws.Cells(j + 2, 12).Value = total_stock_vol
       ws.Cells(j + 2, 18).Value = ticker
       ws.Cells(j + 2, 19).Value = Value
       ws.Range("r2").Value = GreatestIncreaseTicker
       ws.Range("r3").Value = GreatestDecreaseTicker
       ws.Range("r4").Value = GreatestVolumeTicker
       ws.Range("s2").Value = FormatPercent(GreatestIncrease)
       ws.Range("s3").Value = FormatPercent(GreatestDecrease)
       ws.Range("s4").Value = GreatestTotalVolume
       
       'conditional format for positive and negative qtr_change
       If qtr_change > 0 Then
       ws.Cells(j + 2, 10).Interior.ColorIndex = 4
       ElseIf qtr_change < 0 Then
       ws.Cells(j + 2, 10).Interior.ColorIndex = 3
       Else
       End If
       
       
       'If next ticker is different, group finished
       total_stock_vol = 0
       j = j + 1
       
        End If
    Next i
    
 Next ws
End Sub