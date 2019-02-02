Sub Easysolution():

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

ws.Range("A:J").Columns.AutoFit

Next ws


End Sub






