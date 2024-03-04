Attribute VB_Name = "Module2"
Sub AnalyzeStockData()
Dim ws As Worksheet
Dim ticker As String
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalVolume As Double
Dim startPrice As Double
Dim row As Long
Dim summaryRow As Integer
Dim lastRow As Long

' Variables for tracking the greatest % increase, % decrease, and total volume
Dim greatestInc As Double
Dim greatestDec As Double
Dim greatestVol As Double
Dim greatestIncTicker As String
Dim greatestDecTicker As String
Dim greatestVolTicker As String

For Each ws In ThisWorkbook.Sheets
row = 2
summaryRow = 2
totalVolume = 0
startPrice = ws.Cells(2, 3).Value
greatestInc = 0
greatestDec = 0
greatestVol = 0

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Volume"
' Headers for greatest % increase, decrease, and total volume
ws.Range("O1").Value = "Metric"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

For i = 2 To lastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
totalVolume = totalVolume + ws.Cells(i, 7).Value
yearlyChange = ws.Cells(i, 6).Value - startPrice
If startPrice = 0 Then
percentChange = 0
Else
percentChange = yearlyChange / startPrice
End If
ws.Cells(summaryRow, 9).Value = ticker
ws.Cells(summaryRow, 10).Value = yearlyChange
ws.Cells(summaryRow, 11).Value = percentChange
ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
ws.Cells(summaryRow, 12).Value = totalVolume

' Check and update greatest increase, decrease, and volume
If percentChange > greatestInc Then
greatestInc = percentChange
greatestIncTicker = ticker
End If
If percentChange < greatestDec Then
greatestDec = percentChange
greatestDecTicker = ticker
End If
If totalVolume > greatestVol Then
greatestVol = totalVolume
greatestVolTicker = ticker
End If

' Reset for next stock
summaryRow = summaryRow + 1
totalVolume = 0
If ws.Cells(i + 1, 1).Value <> "" Then
startPrice = ws.Cells(i + 1, 3).Value
End If
Else
totalVolume = totalVolume + ws.Cells(i, 7).Value
End If
Next i

' Output the greatest % increase, decrease, and volume
ws.Cells(2, 16).Value = greatestIncTicker
ws.Cells(2, 17).Value = Format(greatestInc, "Percent")
ws.Cells(3, 16).Value = greatestDecTicker
ws.Cells(3, 17).Value = Format(greatestDec, "Percent")
ws.Cells(4, 16).Value = greatestVolTicker
ws.Cells(4, 17).Value = greatestVol

' Conditional Formatting
For j = 2 To summaryRow
If ws.Cells(j, 10).Value > 0 Then
ws.Cells(j, 10).Interior.Color = vbGreen
Else
ws.Cells(j, 10).Interior.Color = vbRed
End If
Next j

Next ws

MsgBox "Stock Data Analysis Complete!"

End Sub
