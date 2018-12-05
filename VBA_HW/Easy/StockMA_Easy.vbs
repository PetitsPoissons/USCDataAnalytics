Sub StockMA_Easy()

    For Each ws in Worksheets

        Dim ticker As String
        Dim volume As Double
        Dim last_row As Long
        Dim table_row As Long
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        last_row = ws.Cells(Rows.Count,1).End(xlUp).Row
        table_row = 2
        ticker = ws.Range("A2").Value
        volume = 0

        For row = 2 To (last_row + 1)

            If ticker = ws.Cells(row,1).Value Then
                volume = volume + ws.Cells(row,7).Value
            Else
                ws.Cells(table_row,9).Value = ticker
                ws.Cells(table_row,10).Value = volume
                table_row = table_row + 1
                ticker = ws.Cells(row,1).Value
                volume = ws.Cells(row,7).Value
            End If

        Next row

    Next ws

End Sub