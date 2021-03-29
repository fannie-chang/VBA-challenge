Sub stock_analysis()

Dim ticker As String

Dim open_price As Double
Dim high_price As Double
Dim low_price As Double
Dim close_price As Double

Dim total_volume As Double
Dim yearly_change As Double
Dim percent_change As Double


Dim Summary_Table_Row As Integer


Dim ws As Worksheet

For Each ws In Worksheets
Summary_Table_Row = 2

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To LastRow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ticker = ws.Cells(i, 1).Value

total_volume = total_volume + ws.Cells(i, 7).Value
ws.Range("I" & Summary_Table_Row).Value = ticker
ws.Range("L" & Summary_Table_Row).Value = total_volume

open_price = ws.Range("C" & Summary_Table_Row).Value

close_price = ws.Cells(i, 6).Value

yearly_change = close_price - open_price

ws.Range("J" & Summary_Table_Row).Value = yearly_change


    If yearly_change > 0 Then

        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

    Else

        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

    End If

If open_price = 0 Then

    percent_change = 0

Else

    percent_change = (yearly_change / open_price)

End If

ws.Range("K" & Summary_Table_Row).Value = Format(percent_change, "Percent")

ws.Range("K" & Summary_Table_Row).Value = percent_change



If percent_change > 0 Then

    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4

Else

    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3


End If
Summary_Table_Row = Summary_Table_Row + 1

total_volume = 0


Else
total_volume = total_volume + ws.Cells(i, 7).Value


End If

Next i


Next ws

End Sub
