Sub Combinedsolutions():

'First: easy solution part
'loop through each worksheet
For Each ws In Worksheets

'set up title
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

'define variables, numT=number with the same Ticker
Dim ticker As String
Dim i As Long
Dim numT As Long
Dim lastrow As Long
'Dim lastcolumn As Long

'set up value for last row and column
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

'Set up initial Ticker value
numT = 0
ticker = ws.Cells(2, 1).Value
ws.Cells(2, 9).Value = ticker



'create for loop for total set of data
For i = 2 To lastrow

        If ticker = ws.Cells(i, 1).Value Then
            ws.Cells(i - numT, 10).Value = ws.Cells(i - numT, 10).Value + ws.Cells(i, 7).Value
            numT = numT + 1

        Else
            ticker = ws.Cells(i, 1).Value
            ws.Cells(i - numT + 1, 9).Value = ticker
            ws.Cells(i - numT + 1, 10).Value = ws.Cells(i, 7).Value
            
        End If
      
Next i



'Second: Moderate solution part
ws.Range("J1").EntireColumn.Insert
ws.Range("K1").EntireColumn.Insert

ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"

'define variables,
Dim openv As Double
Dim closev As Double

'set up value for last row and column
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

'Set up initial values
numT = 0
ticker = ws.Cells(2, 1).Value
openv = ws.Cells(2, 3).Value

'create for loop for total set of data
For i = 2 To lastrow

        If ws.Cells(i, 1).Value = ticker Then
            
            numT = numT + 1

        Else
            closev = ws.Cells(i - 1, 6).Value
            
            ws.Cells(i - numT, 10).Value = closev - openv
            ws.cells(i - numT, 10).numberformat="0.000000000"
            
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
ws.Cells(lastrow - numT + 1, 10).Value = closev - openv
ws.Cells(lastrow - numT + 1, 10).numberformat="0.000000000"
ws.Cells(lastrow - numT + 1, 11).Value = Format((closev / openv - 1), "Percent")


'Third: hard solution part

Dim resultrows As Long
resultrows = ws.Cells(1, 10).End(xlDown).Row

For i = 2 To resultrows
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    
Next i


Dim increasep As Double
Dim decreasep As Double
Dim tvol As Double
Dim tickerI As String
Dim tickerD As String
Dim tickerV As String


tvol = ws.Range("L2").Value
increasep = ws.Range("K2").Value
decreasep = ws.Range("K2").Value
tickerI = ws.Range("J2").Value
tickerD = ws.Range("J2").Value
tickerV = ws.Range("J2").Value


For i = 3 To resultrows
    If ws.Cells(i, 11).Value >= increasep And ws.Cells(i, 11).Value <> "N/A" Then
        increasep = ws.Cells(i, 11).Value
        tickerI = ws.Cells(i, 9).Value
    
    ElseIf ws.Cells(i, 11).Value <= decreasep And ws.Cells(i, 11).Value <> "N/A" Then
        decreasep = ws.Cells(i, 11).Value
        tickerD = ws.Cells(i, 9).Value

    End If
    
    If ws.Cells(i, 12).Value >= tvol Then
        tvol = ws.Cells(i, 12).Value
        tickerT = ws.Cells(i, 9).Value
    End If

Next i

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volumn"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("Q2").Value = Format(increasep, "Percent")
ws.Range("Q3").Value = Format(decreasep, "Percent")
ws.Range("Q4").Value = tvol
ws.Range("P2").Value = tickerI
ws.Range("P3").Value = tickerD
ws.Range("P4").Value = tickerT



ws.Range("I:Q").Columns.AutoFit
Next ws

End Sub






