Attribute VB_Name = "Coloring"
Sub color():

GreenColor = RGB(0, 128, 0)

RedColor = RGB(255, 0, 0)

 

'Get number of rows in the specified column

RowsCount = Range("J2", Range("J2").End(xlDown)).Rows.Count

 

'Select cell

Range("J2").Select

 

'Loop the cells

For i = 1 To RowsCount

    If ((ActiveCell.Value) > 0) Then

        ActiveCell.Interior.color = GreenColor

    Else

       ActiveCell.Interior.color = RedColor

  

    End If

 

    ActiveCell.Offset(1, 0).Select
    

 

Next i


End Sub
