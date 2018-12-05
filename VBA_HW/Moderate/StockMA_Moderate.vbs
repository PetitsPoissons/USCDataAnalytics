Sub StockMA_Moderate()

    For Each ws in Worksheets

        Dim ticker As String
        Dim volume, opening_price, closing_price, yearly_change As Double
        Dim row, last_row, table_row, year_start, year_end, count_row As Long

        'Write headers for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Initialize some variables
        last_row = ws.Cells(Rows.Count,1).End(xlUp).Row
        table_row = 2
        year_start = 2
        ticker = ws.Range("A2").Value
        volume = 0
        count_row = -1

        For row = 2 To (last_row + 1)

            'Evaluate if ticker changed
            If ticker = ws.Cells(row,1).Value Then
                volume = volume + ws.Cells(row,7).Value
                count_row = count_row + 1

            Else
                'Calculate Yearly Change
                opening_price = ws.Cells(year_start,3).Value
                year_end = year_start + count_row
                closing_price = ws.Cells(year_end,6).Value
                yearly_change = closing_price - opening_price

                'Fill in the summary table for that ticker
                ws.Cells(table_row,9).Value = ticker
                ws.Cells(table_row,10).Value = yearly_change
                If opening_price = 0 Then
                    ws.Cells(table_row,11).Value = "N/A"
                Else
                    ws.Cells(table_row,11).Value = yearly_change / opening_price
                End If
                ws.Cells(table_row,12).Value = volume

                'Apply conditional formatting on Yearly Change
                If yearly_change < 0 Then
                    ws.Cells(table_row,10).Interior.ColorIndex = 3
                Else
                    ws.Cells(table_row,10).Interior.ColorIndex = 4
                End If

                'Reset variables for the next ticker
                ticker = ws.Cells(row,1).Value
                year_start = year_start + count_row + 1
                count_row = 0
                volume = ws.Cells(row,7).Value
                table_row = table_row + 1

            End If

        Next row

        'Format the summary table some more...
        ws.Columns(11).EntireColumn.NumberFormat = "0.00%"
        For j = 9 To 12
            ws.Columns(j).EntireColumn.AutoFit
        Next j
        
    Next ws

End Sub