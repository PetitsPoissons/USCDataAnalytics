Sub StockMA_Hard()

    For Each ws in Worksheets

        Dim ticker As String
        Dim volume, opening_price, closing_price, yearly_change As Double
        Dim row, last_row, table_row, year_start, year_end, count_row, end_row As Long
        Dim num1, num2, num3 As Double
        Dim ticker1, ticker2, ticker3 As Long

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

            'Evaluate if ticker changes
            If ticker = ws.Cells(row,1).Value Then
                volume = volume + ws.Cells(row,7).Value
                count_row = count_row + 1

            Else
                'Calculate Yearly Change
                opening_price = ws.Cells(year_start,3).Value
                year_end = year_start + count_row
                closing_price = ws.Cells(year_end,6).Value
                yearly_change = closing_price - opening_price

                'Fill in summary table for current ticker
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

                'Reset variables for next ticker
                ticker = ws.Cells(row,1).Value
                year_start = year_start + count_row + 1
                count_row = 0
                volume = ws.Cells(row,7).Value
                table_row = table_row + 1

            End If

        Next row

        'Format summary table some more...
        ws.Columns(11).EntireColumn.NumberFormat = "0.00%"
        For j = 9 To 12
            ws.Columns(j).EntireColumn.AutoFit
        Next j

        'PART III -----------------------

        'Calculate last row number of the summary table
        end_row = table_row - 1
        
        'Find ticker with greatest % increase
        num1 = WorksheetFunction.Max(ws.Range("K2:K" & end_row))
        ticker1 = WorksheetFunction.Match(num1,ws.Range("K2:K" & end_row),0)
        ws.Range("P2").Value = ws.Cells(ticker1 + 1, 9).Value
        ws.Range("Q2").Value = num1

        'Find ticker with greatest % decrease
        num2 = WorksheetFunction.Min(ws.Range("K2:K" & end_row))
        ticker2 = WorksheetFunction.Match(num2,ws.Range("K2:K" & end_row),0)
        ws.Range("P3").Value = ws.Cells(ticker2 + 1, 9).Value
        ws.Range("Q3").Value = num2

        'Find ticker with greatest total volume
        num3 = WorksheetFunction.Max(ws.Range("L2:L" & end_row))
        ticker3 = WorksheetFunction.Match(num3,ws.Range("L2:L" & end_row),0)
        ws.Range("P4").Value = ws.Cells(ticker3 + 1, 9).Value
        ws.Range("Q4").Value = num3

        'Add labels and format
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        For x = 15 To 17
            ws.Columns(x).EntireColumn.AutoFit
        Next x

    Next ws

End Sub