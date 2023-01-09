Attribute VB_Name = "Module1"
Sub stock_data()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

Dim WorksheetName As String
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
WorksheetName = ws.Name
Dim ticker As String
ticker = " "
Dim total_vol As LongLong
total_vol = 0
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
Dim year_open As Double
year_open = 0
Dim year_close As Double
year_close = 0
Dim yearly_change As Double
yearly_change = 0
Dim yearly_percent_change As Double
yearly_percent_change = 0
Dim max_ticker As String
max_ticker = " "
Dim min_ticker As String
min_ticker = " "
Dim max_percent As Double
max_percent = 0
Dim min_percent As Double
min_percent = 0
Dim max_vol_ticker As String
max_vol_ticker = " "
Dim max_vol As LongLong
max_vol = 0
year_open = ws.Cells(2, 3).Value

ws.Cells(1, 9).Value = "ticker"
ws.Cells(1, 10).Value = "Yearly_change"
ws.Cells(1, 11).Value = "Yearly_percentage"
ws.Cells(1, 12).Value = "Total_Vol"
For i = 2 To lastrow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
year_close = ws.Cells(i, 6).Value
yearly_change = year_close - year_open
If year_open <> 0 Then
yearly_percent_change = (yearly_change / year_open) * 100
End If
total_vol = total_vol + ws.Cells(i, 7).Value

ws.Range("I" & Summary_Table_Row).Value = ticker
ws.Range("J" & Summary_Table_Row).Value = yearly_change
ws.Range("K" & Summary_Table_Row).Value = yearly_percent_change
ws.Range("L" & Summary_Table_Row).Value = total_vol
Summary_Table_Row = Summary_Table_Row + 1
year_open = ws.Cells(i + 1, 3).Value

If (yearly_percent_change > max_percent) Then
max_percent = yearly_percent_change
max_ticker = ticker
ElseIf (yearly_percent_change < min_percent) Then
min_percent = yearly_percent_change
min_ticker = ticker
End If

If (total_vol > max_vol) Then
max_vol = total_vol
max_vol_ticker = ticker
ElseIf (total_vol < min_vol) Then
min_vol = total_vol
min_vol_ticker = ticker
End If

yearly_percent_change = 0
total_vol = 0
Else

total_vol = total_vol + ws.Cells(i, 7).Value


End If

If ws.Cells(i, 10).Value >= 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
Else: ws.Cells(i, 10).Interior.ColorIndex = 3
End If
ws.Range("K" & Summary_Table_Row).Value = (CStr(yearly_percent_change) & "%")


Next i

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P2").Value = max_ticker
ws.Range("P3").Value = min_ticker
ws.Range("Q2").Value = (CStr(max_percent) & "%")
ws.Range("Q3").Value = (CStr(min_percent) & "%")
ws.Range("Q4").Value = max_vol

Next ws
MsgBox ("Complete")
End Sub
