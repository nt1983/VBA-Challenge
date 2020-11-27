Sub stock()

For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    Total_Volume = 0
    Table_Row = 2
    close_yearly = 0
    open_yearly = ws.Cells(Table_Row, 3).Value
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker_Name = ws.Cells(i, 1).Value
            ws.Range("I" & Table_Row).Value = Ticker_Name
            close_yearly = ws.Cells(i, 6).Value
            yearly_change = close_yearly - open_yearly
            ws.Range("J" & Table_Row).Value = yearly_change
            If yearly_change < 0 Then
                ws.Range("J" & Table_Row).Interior.Color = RGB(255, 0, 0)
            ElseIf yearly_change > 0 Then
                ws.Range("J" & Table_Row).Interior.Color = RGB(0, 255, 0)
            End If
            If ((yearly_change <> 0) And (open_yearly <> 0)) Then
                percent_change = yearly_change / open_yearly
            Else
                percent_change = 0
            End If
            ws.Range("K" & Table_Row).Value = percent_change
            ws.Range("K" & Table_Row).NumberFormat = "0.00%"
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            ws.Range("L" & Table_Row).Value = Total_Volume
            Table_Row = Table_Row + 1
            next_ticker = i + 1
            open_yearly = ws.Cells(next_ticker, 3).Value
            Total_Volume = 0
                        
        Else
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        End If
    Next i

'Bonus Homework-----------------------------------------------------
    increase_ticker = Application.WorksheetFunction.Max(ws.Range("K:K"))
    decrease_ticker = Application.WorksheetFunction.Min(ws.Range("K:K"))
    Greatest_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
    Row = ws.Range("K" & Rows.Count).End(xlUp).Row
    For j = 2 To Row
        If ws.Cells(j, 11).Value = increase_ticker Then
            ws.Range("P2").Value = ws.Cells(j, 9).Value
            ws.Range("Q2").Value = increase_ticker
            ws.Range("Q2").NumberFormat = "0.00%"
        ElseIf ws.Cells(j, 11).Value = decrease_ticker Then
            ws.Range("P3").Value = ws.Cells(j, 9).Value
            ws.Range("Q3").Value = decrease_ticker
            ws.Range("Q3").NumberFormat = "0.00%"
        End If
        If ws.Cells(j, 12).Value = Greatest_Volume Then
            ws.Range("P4").Value = ws.Cells(j, 9).Value
            ws.Range("Q4").Value = Greatest_Volume
        End If
    Next j

Next ws
End Sub
