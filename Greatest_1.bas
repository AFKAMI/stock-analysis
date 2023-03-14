Attribute VB_Name = "Greatest1"
Sub Greatest():


Dim j As Integer
Dim ws_num As Integer

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
ws_num = ThisWorkbook.Worksheets.Count

For j = 1 To ws_num
    ThisWorkbook.Worksheets(j).Activate
    'do whatever you need

Dim r As Range
    Dim m As Double
    Set r = Range("K:K")
    m = Application.WorksheetFunction.Max(r)
    
    min_ = Application.WorksheetFunction.Min(r)
    
     Range("P3").Value = "Greatest % Increase"
     Range("P4").Value = "Greatest % Decrease"
     Range("P5").Value = "Greatest Total Volume"
     Range("Q2") = "Ticker"
     Range("R2") = "Value"
     
     lastrow = Cells(Rows.Count, 1).End(xlUp).Row
     
     
     
    ' Greatest % increase
     For i = 2 To lastrow
    If Cells(i, 11).Value = m Then
    max_value = Cells(i, 11).Value
    stock_max = Cells(i, 9).Value
    Range("R3").Value = FormatPercent(max_value)
    Range("Q3").Value = stock_max
    max_value = 0
    
    End If
    
    Next i
    
    
    
    'Greatest % Decrease
    
      For i = 2 To lastrow
    If Cells(i, 11).Value = min_ Then
    min_value = Cells(i, 11).Value
    stock_min = Cells(i, 9).Value
    Range("R4").Value = FormatPercent(min_value)
    Range("Q4").Value = stock_min
    min_value = 0
    
    End If
    
    Next i
    
    
    ' Greatest Total Volume
    'Dim r2 As Range
    Dim vol As Double
    Set r2 = Range("L:L")
     vol = Application.WorksheetFunction.Max(r2)
     lastrow = Cells(Rows.Count, 1).End(xlUp).Row
   
    For i = 2 To lastrow
    If Cells(i, 12).Value = vol Then
    total_volume = Cells(i, 12).Value
    stock_total = Cells(i, 9).Value
    Range("R5").Value = FormatPercent(total_volume)
    Range("Q5").Value = stock_total
    total_volume = 0
    
    End If

    
    Next i
    
    
        ThisWorkbook.Worksheets(j).Cells(1, 1) = 1  'this sets cell A1 of each sheet to "1"
Next

starting_ws.Activate 'activate the worksheet that was originally active

    
  Sheet1.Columns("A:R").AutoFit
  Sheet2.Columns("A:R").AutoFit
  Sheet3.Columns("A:R").AutoFit

End Sub
