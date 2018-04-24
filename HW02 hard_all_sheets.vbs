Sub hard_all_sheets()
For Each ws In Worksheets

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    TickerRow = 1

    For i = 2 To LastRow
        'Caculate the total volume of a certain ticker one row at a time.
        TotalVolume = TotalVolume + ws.Cells(i, "G")
        If ws.Cells(i, "A") = ws.Cells(i + 1, "A") Then
            'Log the number of rows a ticker expands to.
            RowNum = RowNum + 1
        Else
            'Decide which row to output the results of each ticker.
            TickerRow = TickerRow + 1
            'Output ticker name.
            ws.Cells(TickerRow, "I") = ws.Cells(i, "A")
            'Output yearly price change.
            OpenPrice = ws.Cells(i - RowNum, "C")
            ClosePrice = ws.Cells(i, "F")
            ws.Cells(TickerRow, "J") = ClosePrice - OpenPrice
            'Output percentage of price change. Set to 0 if the open price is 0.
            If OpenPrice <> 0 Then
                ws.Cells(TickerRow, "K") = (ClosePrice - OpenPrice) / OpenPrice
            Else
                ws.Cells(TickerRow, "K") = 0
            End If
            'Output the total volume of a ticker.
            ws.Cells(TickerRow, "L") = TotalVolume
            'Reset the total volume and the number of rows before moving on to the next ticker.
            TotalVolume = 0
            RowNum = 0
        End If
    Next i

    'Add headers and formats.
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("K2:K" & TickerRow).NumberFormat = "0.00%"
    ws.Range("L1") = "Total Stock Volume"
    ws.Columns("J:L").AutoFit

    'Conditionally format "Yearly Change" column.
    For i = 2 To TickerRow
        If ws.Cells(i, "J") > 0 Then
            ws.Cells(i, "J").Interior.ColorIndex = 4
        ElseIf ws.Cells(i, "J") < 0 Then
            ws.Cells(i, "J").Interior.ColorIndex = 3
        End If
    Next i

    'Add headers.
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"

    'Loop to compare each row to decide max/min value.
    'Here ">" and "<" are used to output the first cell if threre are more than one max/min values.
    'Use ">=" and "<=" to output the last cell instead in the above senario.
    For i = 2 To TickerRow
        If MaxIn < ws.Cells(i, "K") Then
            MaxInTicker = ws.Cells(i, "I")
            MaxIn = ws.Cells(i, "K")
        End If

        If MaxDe > ws.Cells(i, "K") Then
            MaxDeTicker = ws.Cells(i, "I")
            MaxDe = ws.Cells(i, "K")
        End If

        If MaxVol < ws.Cells(i, "L") Then
            MaxVolTicker = ws.Cells(i, "I")
            MaxVol = ws.Cells(i, "L")
        End If
    Next i

    'Output the results calculated above.
    ws.Range("P2") = MaxInTicker
    ws.Range("P3") = MaxDeTicker
    ws.Range("P4") = MaxVolTicker
    ws.Range("Q2") = MaxIn
    ws.Range("Q3") = MaxDe
    ws.Range("Q4") = MaxVol

    'Add format.
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Columns("O:Q").AutoFit

    'Reset the values before moving on to the next worksheet.
    MaxIn = 0
    MaxDe = 0
    MaxVol = 0
Next ws
End Sub
