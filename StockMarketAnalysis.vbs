Sub stockMktScript():

Dim ws As Worksheet
Dim i As Double
Dim ticker As String
Dim first As Double
Dim last As Double
Dim volume As Double
Dim RowCnt As Double
Dim p As Double
Dim grtInc As Double
Dim grtDec As Double
Dim grtVol As Double

'looping through each worksheet

For Each ws In ThisWorkbook.Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

ws.Range("I1:L1").Font.Bold = True

i = 2
p = 2


RowCnt = ws.UsedRange.Rows.Count


Do While (i <= RowCnt)

first = ws.Cells(i, 3)
ticker = ws.Cells(i, 1)
volume = 0


Do While (i <= RowCnt)

If (ticker = ws.Cells(i, 1)) Then

volume = volume + ws.Cells(i, 7)

last = ws.Cells(i, 6)

i = i + 1

Else

Exit Do

End If

Loop

ws.Cells(p, 9) = ticker

ws.Cells(p, 10) = last - first

If (first = 0) Then
ws.Cells(p, 11) = 0
Else

ws.Cells(p, 11) = Format(((last - first) / first) * 100, "0.00") + "%"

End If


ws.Cells(p, 12) = volume

'Conditional formatting

If last - first > 0 Then
ws.Cells(p, 10).Interior.Color = RGB(0, 255, 0)
ws.Cells(p, 11).Interior.Color = RGB(0, 255, 0)



ElseIf last - first < 0 Then

ws.Cells(p, 10).Interior.Color = RGB(255, 0, 0)
ws.Cells(p, 11).Interior.Color = RGB(255, 0, 0)

Else

ws.Cells(p, 10).Interior.Color = RGB(255, 255, 255)
ws.Cells(p, 10).Interior.Color = RGB(255, 255, 255)

End If

p = p + 1


Loop

'Calculating Greatest Increase
grtInc = WorksheetFunction.Max(ws.Range("K2:K" & (p - 1)).Value)
ws.Cells(2, 16).Value = WorksheetFunction.Index(ws.Range("I2:I" & (p - 1)).Value, WorksheetFunction.Match(grtInc, ws.Range("K2:K" & (p - 1)).Value, 0))
ws.Cells(2, 17).Value = Format(grtInc * 100, "0.00") + "%"

'Calculating Greatest Decrease
grtDec = WorksheetFunction.Min(ws.Range("K2:K" & (p - 1)).Value)
ws.Cells(3, 16).Value = WorksheetFunction.Index(ws.Range("I2:I" & (p - 1)).Value, WorksheetFunction.Match(grtDec, ws.Range("K2:K" & (p - 1)).Value, 0))
ws.Cells(3, 17).Value = Format(grtDec * 100, "0.00") + "%"

'Calculating greatest Volume

grtVol = WorksheetFunction.Max(ws.Range("L2:L" & (p - 1)).Value)
ws.Cells(4, 16).Value = WorksheetFunction.Index(ws.Range("I2:I" & (p - 1)).Value, WorksheetFunction.Match(grtVol, ws.Range("L2:L" & (p - 1)).Value, 0))
ws.Cells(4, 17).Value = grtVol

Next ws

End Sub

