Attribute VB_Name = "Stock_Data"
Sub Stock_Data():

Dim stock_name As String
Dim total_vol As Double
Dim OUT_ROW_COUNT As Integer



For Each ws In Worksheets

' set Headers

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Vol"


OUT_ROW_COUNT = 2
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
total_vol = 0
open_market = ws.Cells(2, 3).Value

  For i = 2 To lastrow
 
    total_vol = total_vol + ws.Cells(i, 7).Value
 
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
     'calculate Stock name and Volume
     stock_name = ws.Cells(i, 1).Value
     ws.Range("I" & OUT_ROW_COUNT).Value = stock_name
     ws.Range("L" & OUT_ROW_COUNT).Value = total_vol
     
     
     'calculate Yearly Change
     close_market = ws.Cells(i, 6).Value
     yearly_change = close_market - open_market
     ws.Range("J" & OUT_ROW_COUNT).Value = yearly_change
     
     ' calculate Percent Change
     percent_change_market = ((close_market * 100) / open_market) - 100
     percent_change_market = (Round(percent_change_market, 2)) / 100
     ws.Range("K" & OUT_ROW_COUNT).Value = FormatPercent(percent_change_market)
     
      
     open_market = ws.Cells(i + 1, 3).Value
     
     
     
     OUT_ROW_COUNT = OUT_ROW_COUNT + 1
     
     
     total_vol = 0
     
     
     
     End If
 
  Next i
  Next ws
  
  Sheet1.Columns("A:P").AutoFit
  Sheet2.Columns("A:P").AutoFit
  Sheet3.Columns("A:P").AutoFit
  
  
  
  
  
  
  
  

End Sub
