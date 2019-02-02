Sub Hardsolution()

For Each ws In Worksheets

Dim increasep As Double
Dim decreasep As Double
Dim tvol As Double
Dim tickerI As String
Dim tickerD As String
Dim tickerV As String

Dim resultrows As Long
resultrows = ws.Cells(1, 10).End(xlDown).Row

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

ws.Range("O:Q").Columns.AutoFit


Next ws



End Sub

