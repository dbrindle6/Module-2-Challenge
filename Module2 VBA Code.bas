Attribute VB_Name = "Module1"
Sub QuarterlyStockDataCalculations()

'Establishing the Variables
Dim ws As Worksheet
Dim ticker As String
Dim lastRow As Long
Dim startPrice As Double
Dim endPrice As Double
Dim volumeTotal As Double
Dim currentQuarter As Long
Dim previousQuarter As Long
Dim currentTicker As String
Dim i As Long
Dim outputRow As Long

'To Loop through all worksheets
For Each ws In ThisWorkbook.Worksheets
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'Output Header Columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'To ensure outputs start on the second row
outputRow = 2

'To Loop through all rows of stock data
For i = 2 To lastRow
    
    currentTicker = ws.Cells(i, 1).Value    'To get current ticker
    
'To check if the ticker or quarter has changed
If i = 2 Or currentTicker <> ticker Or currentQuarter <> previousQuarter Then
    If i > 2 Then
'Output data for the previous quarter
ws.Cells(outputRow, 9).Value = ticker
ws.Cells(outputRow, 10).Value = endPrice - startPrice   'The Quarterly change

'To color-code the Quarterly change
If ws.Cells(outputRow, 10).Value > 0 Then
    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 225, 0) 'Green for +
ElseIf ws.Cells(outputRow, 10).Value < 0 Then
    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) 'Red for -
End If

'To Avoid dividing by 0
If startPrice <> 0 Then
    ws.Cells(outputRow, 11).Value = (endPrice - startPrice) / startPrice    '% Change
Else
    ws.Cells(outputRow, 11).Value = 0
End If
    
ws.Cells(outputRow, 12).Value = volumeTotal
outputRow = outputRow + 1   'To move down for the next row of output on the worksheet
End If

'To Reset Tracking Variables
ticker = currentTicker
previousQuarter = currentQuarter
startPrice = ws.Cells(i, 3).Value   'Set the opening price for the new quarter
volumeTotal = 0     'To reset volume total
End If

'To Accumulate the volume and track the end price
volumeTotal = volumeTotal + ws.Cells(i, 7).Value 'Volume column
endPrice = ws.Cells(i, 6).Value 'Closing price
Next i

'Output the final quarter data
ws.Cells(outputRow, 9).Value = ticker
ws.Cells(outputRow, 10).Value = endPrice - startPrice

'Color code the Quarterly Change
If ws.Cells(outputRow, 10).Value > 0 Then
    ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) 'Green for +
ElseIf ws.Cells(outputRow, 10).Value < 0 Then
    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) 'Red for -
End If

'To Avoid dividing by 0
If startPrice <> 0 Then
    ws.Cells(outputRow, 11).Value = (endPrice - startPrice) / startPrice    '% change
Else
    ws.Cells(outputRow, 11).Value = 0
End If

ws.Cells(outputRow, 12).Value = volumeTotal 'Total volume
Next ws

MsgBox "Quarterly stock data summary has been created."

End Sub


