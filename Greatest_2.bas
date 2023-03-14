Attribute VB_Name = "Greatest_2"
Sub Greatest():


For Each ws In Worksheets


Dim r As Range
    Dim m As Double
    Set r = ws.Range("K:K")
    m = Application.WorksheetFunction.Max(r)
    
    min_ = Application.WorksheetFunction.Min(r)
    
     ws.Range("P3").Value = "Greatest % Increase"
     ws.Range("P4").Value = "Greatest % Decrease"
     ws.Range("P5").Value = "Greatest Total Volume"
     ws.Range("Q2") = "Ticker"
     ws.Range("R2") = "Value"
     
     lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     
     
    ' Greatest % increase
     For i = 2 To lastrow
    If ws.Cells(i, 11).Value = m Then
    max_value = ws.Cells(i, 11).Value
    stock_max = ws.Cells(i, 9).Value
    ws.Range("R3").Value = FormatPercent(max_value)
    ws.Range("Q3").Value = stock_max
    max_value = 0
    
    End If
    
    Next i
    
    
    'Greatest % Decrease
    
      For i = 2 To lastrow
    If ws.Cells(i, 11).Value = min_ Then
    min_value = ws.Cells(i, 11).Value
    stock_min = ws.Cells(i, 9).Value
    ws.Range("R4").Value = FormatPercent(min_value)
    ws.Range("Q4").Value = stock_min
    min_value = 0
    
    End If
    
    Next i
    
    
    

    
    ' Greatest Total Volume
    'Dim r2 As Range
    Dim vol As Double
    Set r2 = ws.Range("L:L")
     vol = Application.WorksheetFunction.Max(r2)
     lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
    For i = 2 To lastrow
    If ws.Cells(i, 12).Value = vol Then
    total_volume = ws.Cells(i, 12).Value
    stock_total = ws.Cells(i, 9).Value
    ws.Range("R5").Value = FormatPercent(total_volume)
    ws.Range("Q5").Value = stock_total
    total_volume = 0
    
    End If

    
    Next i
    
    Next ws
    
    
  
    
  'Sheet1.Columns("A:P").AutoFit
  'Sheet2.Columns("A:P").AutoFit
  'Sheet3.Columns("A:P").AutoFit

End Sub
