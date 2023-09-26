Sub Stock_Analysis()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Dim Ticker As String

Dim YearlyChange As Double
YearlyChange = 0

Dim PercentChange As Double
PercentChange = 0

Dim TotalVolume As Double
TotalVolume = 0

Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim SummaryRow As Long
SummaryRow = 2

Dim OpeningPrice As Double
OpeningPrice = 0

Dim ClosingPrice As Double
ClosingPrice = 0


For i = 2 To LastRow

If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

OpeningPrice = ws.Cells(i, 3).Value

End If

TotalVolume = TotalVolume + ws.Cells(i, 7).Value

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

ws.Cells(SummaryRow, 9).Value = ws.Cells(i, 1).Value

ws.Cells(SummaryRow, 12).Value = TotalVolume

ClosingPrice = ws.Cells(i, 6).Value

YearlyChange = ClosingPrice - OpeningPrice
ws.Cells(SummaryRow, 10).Value = YearlyChange

If YearlyChange >= 0 Then
ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
ElseIf YearlyChange < 0 Then
ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
End If


If OpeningPrice = 0 And ClosingPrice = 0 Then
YearlyChange = 0
ws.Cells(SummaryRow, 11).Value = PercentChange
ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"

ElseIf OpeningPrice = 0 Then
PercentChange_NA = "New Stock"
ws.Cells(SummaryRow, 11).Value = PercentChange

Else
PercentChange = YearlyChange / OpeningPrice
ws.Cells(SummaryRow, 11).Value = PercentChange
ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"


End If

SummaryRow = SummaryRow + 1

TotalVolume = 0
OpeningPrice = 0
ClosingPrice = 0
YearlyChange = 0
PercentChange = 0

End If

Next i

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

Dim IncreasedStock As String
Dim IncreasedValue As Double

IncreasedValue = ws.Cells(2, 11).Value

Dim DecreasedStock As String
Dim DecreasedValue As Double

DecreasedValue = ws.Cells(2, 11).Value

Dim TotalVolumeStock As String
Dim TotalVolumeValue As Double

TotalVolumeStock = ws.Cells(2, 12).Value

For j = 2 To LastRow

If ws.Cells(j, 11).Value > IncreasedValue Then
IncreasedValue = ws.Cells(j, 11).Value
IncreasedStock = ws.Cells(j, 9).Value

End If

If ws.Cells(j, 11).Value < DecreasedValue Then
DecreasedValue = ws.Cells(j, 11).Value
DecreasedStock = ws.Cells(j, 9).Value
End If

If ws.Cells(j, 12).Value > TotalVolumeValue Then
TotalVolumeValue = ws.Cells(j, 12).Value
TotalVolumeStock = ws.Cells(j, 9).Value
End If

Next j

ws.Cells(2, 16).Value = IncreasedStock
ws.Cells(2, 17).Value = IncreasedValue
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = DecreasedStock
ws.Cells(3, 17).Value = DecreasedValue
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 16).Value = TotalVolumeStock
ws.Cells(4, 17).Value = TotalVolumeValue
ws.Cells(4, 17).NumberFormat = "0.00E+00"


ws.Columns("I:L").EntireColumn.AutoFit
ws.Columns("O:Q").EntireColumn.AutoFit



Next ws

End Sub


