VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker_volume()
For Each ws In Worksheets
Dim WorksheetName As String
Dim stock_total As Double
stock_total = 0
Dim ticker_row As Long
ticker_row = 2
Dim ticker As String
Dim old_I As Double
old_I = 2


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
WorksheetName = ws.Name

'Add Column Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Great % increase"
ws.Cells(3, 15).Value = "Great % decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Check if we are still in the same ticker
For I = 2 To LastRow
If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
yearly_change = ws.Range("F" & I).Value - ws.Range("C" & old_I).Value
percent_change = (yearly_change / ws.Range("C" & old_I).Value)
stock_total = stock_total + ws.Cells(I, 7).Value
ticker = ws.Cells(I, 1).Value
ws.Range("I" & ticker_row).Value = ticker
ws.Range("L" & ticker_row).Value = stock_total
ws.Range("J" & ticker_row).Value = yearly_change
ws.Range("K" & ticker_row).Value = percent_change
ws.Range("K" & ticker_row).NumberFormat = "0.00%"
    If ws.Range("K" & ticker_row).Value < 0 Then
    ws.Range("K" & ticker_row).Interior.ColorIndex = 3
    Else
    ws.Range("K" & ticker_row).Interior.ColorIndex = 4
    End If
    
old_I = I + 1
stock_total = 0
ticker_row = ticker_row + 1
Else
stock_total = stock_total + ws.Cells(I, 7).Value
End If
Next I

Dim maxValue_percent As Double
Dim minValue_percent As Double
Dim maxValue_volume As Double
Dim x As Double
Dim ticker1 As String
LastRow1 = ws.Cells(Rows.Count, 10).End(xlUp).Row

For x = 2 To LastRow1
    If ws.Range("K" & x).Value >= maxValue_percent Then
    maxValue_percent = ws.Range("K" & x).Value
    ticker1 = ws.Range("I" & x).Value
    End If
    Next x
ws.Cells(2, 17).Value = maxValue_percent
ws.Cells(2, 16).Value = ticker1
For x = 2 To LastRow1
    If ws.Range("K" & x).Value <= minValue_percent Then
    minValue_percent = ws.Range("K" & x).Value
    ticker1 = ws.Range("I" & x).Value
    End If
    Next x
ws.Cells(3, 17).Value = minValue_percent
ws.Cells(3, 16).Value = ticker1
For x = 2 To LastRow1
    If ws.Range("L" & x).Value >= maxValue_volume Then
    maxValue_volume = ws.Range("L" & x).Value
    ticker1 = ws.Range("I" & x).Value
    End If
    Next x
ws.Cells(4, 17).Value = maxValue_volume
ws.Cells(4, 16).Value = ticker1
ws.Range("Q" & 2).NumberFormat = "0.00%"
ws.Range("Q" & 3).NumberFormat = "0.00%"
Next ws

End Sub
