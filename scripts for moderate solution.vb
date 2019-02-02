Sub Moderatesolution():

For Each ws In Worksheets

ws.Range("J1").EntireColumn.Insert
ws.Range("K1").EntireColumn.Insert

ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"

'define variables, numT=number with the same Ticker
Dim ticker As String
Dim i As Long
Dim numT As Long
Dim lastrow As Long
'Dim lastcolumn As Long
Dim openv As Double
Dim closev As Double

'set up value for last row and column
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

'Set up initial Ticker value
numT = 0
ticker = ws.Cells(2, 1).Value
openv = ws.Cells(2, 3).Value

'create for loop for total set of data
For i = 2 To lastrow

        If ws.Cells(i, 1).Value = ticker Then
            
            numT = numT + 1

        Else
            closev = ws.Cells(i - 1, 6).Value
            
            ws.Cells(i - numT, 10).Value = Format((closev - openv), "0.000000000")
            
            If openv = 0 Then
            ws.Cells(i - numT, 11).Value = "N/A"
            
            Else
            ws.Cells(i - numT, 11).Value = Format((closev / openv - 1), "Percent")
            
            End If
            
            openv = ws.Cells(i, 3).Value
            ticker = ws.Cells(i, 1).Value
            
            
        End If
      
                        
Next i

closev = ws.Cells(lastrow, 6).Value
ws.Cells(lastrow - numT + 1, 10).Value = Format((closev - openv), "0.000000000")
ws.Cells(lastrow - numT + 1, 11).Value = Format((closev / openv - 1), "Percent")

ws.Range("I:L").Columns.AutoFit


Dim resultrows As Long
resultrows = ws.Cells(1, 10).End(xlDown).Row

For i = 2 To resultrows
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    
Next i

Next ws


End Sub




