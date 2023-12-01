Attribute VB_Name = "Module5"
Sub FindGreatestMetrics()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Stock_Data") ' Replace with the actual sheet name

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxTickerIncrease As String
    Dim maxTickerDecrease As String
    Dim maxTickerTotalVolume As String

    ' Initialize with a small negative value to ensure any positive percentage will be greater
    maxIncrease = -1
    ' Initialize with a small positive value to ensure any negative percentage will be greater
    maxDecrease = 1
    ' Initialize with a small negative value to ensure any positive total volume will be greater
    maxTotalVolume = -1

    Dim i As Long
    For i = 2 To lastRow
        Dim percentChange As Double
        percentChange = ws.Cells(i, 14).Value ' Assuming percent change is in column N (adjust as needed)

        Dim totalVolume As Double
        totalVolume = ws.Cells(i, 15).Value ' Assuming total volume is in column O (adjust as needed)

        If percentChange > maxIncrease Then
            maxIncrease = percentChange
            maxTickerIncrease = ws.Cells(i, 10).Value ' Assuming ticker symbol is in column J (adjust as needed)
        End If

        If percentChange < maxDecrease Then
            maxDecrease = percentChange
            maxTickerDecrease = ws.Cells(i, 10).Value ' Assuming ticker symbol is in column J (adjust as needed)
        End If

        If totalVolume > maxTotalVolume Then
            maxTotalVolume = totalVolume
            maxTickerTotalVolume = ws.Cells(i, 10).Value ' Assuming ticker symbol is in column J (adjust as needed)
        End If
    Next i

    ' Store the results in specific cells
    ws.Range("Q2").Value = "Greatest Percent Increase:"
    ws.Range("R1").Value = "Ticker"
    ws.Range("S1").Value = "Value"

    ws.Range("R2").Value = maxTickerIncrease
    ws.Range("S2").Value = maxIncrease & "%"

    ws.Range("Q4").Value = "Greatest Percent Decrease:"
    ws.Range("R4").Value = maxTickerDecrease
    ws.Range("S4").Value = maxDecrease & "%"

    ws.Range("Q6").Value = "Greatest Total Volume:"
    ws.Range("R6").Value = maxTickerTotalVolume
    ws.Range("S6").Value = maxTotalVolume

End Sub

