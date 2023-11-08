
Public Sub Main_Macro()

'This is the main macro and runs all subs. Always start the macro from this sub'

Call tickersymbols
Call color_code
Call max_value
Call ticker_max

End Sub

Sub tickersymbols()

'loops through ticker symbols to populate needed values'

For Each ws In Worksheets

Dim ticker_symbol As String
Dim year_end_close As Double
Dim Total_Volume As Double
Dim year_open As Double
Dim start As Double
Dim LastRow As Double
Total_Volume = 0
year_end_close = 0
yearly_price_change = 0
year_open = 0

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

start = 2

For i = 2 To LastRow

'Loops through each ticker to find the start of a new ticker, as well as first open (on the first day of year) and last close(on last day of year)'

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ticker_symbol = ws.Cells(i, 1).Value

year_end_close = ws.Cells(i, 6).Value

year_open = ws.Cells(start, 3)

start = i + 1




Total_Volume = (Total_Volume + Cells(i, 7))


ws.Range("I" & Summary_Table_Row).Value = ticker_symbol

'Range("J" & Summary_Table_Row).Value = year_end_close ----this was used in calculations but does not need to display on the worksheet'

'Range("K" & Summary_Table_Row).Value = year_open ----this was used in calculations but does not need to display on the worksheet'

ws.Range("L" & Summary_Table_Row).Value = Total_Volume

ws.Range("J" & Summary_Table_Row).Value = year_end_close - year_open

ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"

ws.Range("K" & Summary_Table_Row).Value = ((year_end_close - year_open) / year_open)

ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

'start of column titles for worksheet'

ws.Cells(1, 9).Value = "Ticker"

ws.Cells(1, 10).Value = "Yearly Change"

ws.Cells(1, 11).Value = "Percent Change"

ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(2, 15).Value = "Greatest % Increase"

ws.Cells(3, 15).Value = "Greatest % Decrease"

ws.Cells(4, 15).Value = "Greatest Total Volume"

ws.Range("P1").Value = "Ticker"

ws.Range("Q1").Value = "Value"

'End of column titles'

Summary_Table_Row = Summary_Table_Row + 1


Total_Volume = 0

Else

Total_Volume = (Total_Volume + Cells(i, 7))

End If

Next i

Next ws

End Sub

Sub color_code()
'Formats the correct colors in Column J for every worksheet'

For Each ws In Worksheets

Dim LastRow As Double

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

If ws.Cells(i, 10).Value > 0 Then

ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)

ElseIf ws.Cells(i, 10).Value < 0 Then

ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)




End If

Next i

Next ws

End Sub


Sub max_value()

'Finds the max values from columns K and L'

For Each ws In Worksheets

ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K:K"))
ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K:K"))
ws.Range("Q3").NumberFormat = "0.00%"

ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))

Next ws

End Sub

Sub ticker_max()

'Populates the ticker symbol into Column P that corresponds to each max value found in the max_value sub'

For Each ws In Worksheets

Dim ticker As String
Dim ticker2 As String
Dim ticker3 As String
Dim LastRow As Double
Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

If ws.Cells(i, 11).Value = ws.Range("Q2").Value Then

ticker = ws.Cells(i, 9).Value

ws.Range("P" & Summary_Table_Row).Value = ticker

Summary_Table_Row = Summary_Table_Row + 1



End If

If ws.Cells(i, 11).Value = ws.Range("Q3").Value Then

ticker2 = ws.Cells(i, 9).Value

ws.Range("P3").Value = ticker2




End If

If ws.Cells(i, 12).Value = ws.Range("Q4").Value Then

ticker3 = ws.Cells(i, 9).Value

ws.Range("P4").Value = ticker3




End If


Next i

Next ws

End Sub