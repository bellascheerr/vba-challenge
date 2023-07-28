Sub RunStockInsightsOnAllSheets()
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name <> "Summary" Then ' Exclude the Summary sheet itself
            ' Run StockInsights on the current worksheet (year)
            StockInsights ws
        End If
    Next ws
End Sub

Sub StockInsights(ws As Worksheet)

    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double

    ' Find the last row in the current worksheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    totalVolume = 0

    ' Find the last column in the sheet and add 1 to get the first available column for writing data
    Dim summaryColumn As Long
    summaryColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column + 1

    ' Set up Summary headers
    ws.Cells(1, summaryColumn).Value = "Ticker Symbol"
    ws.Cells(1, summaryColumn + 1).Value = "Yearly Change"
    ws.Cells(1, summaryColumn + 2).Value = "Percent Change"
    ws.Cells(1, summaryColumn + 3).Value = "Total Stock Volume"

    ' Loop through all rows in the current worksheet
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        openingPrice = ws.Cells(i, 3).Value
        closingPrice = ws.Cells(i, 6).Value

        ' Calculate yearly change and percent change
        yearlyChange = closingPrice - openingPrice

        ' Check if openingPrice is close to zero and avoid potential overflow
        If Abs(openingPrice) > 1E-06 Then    ' You can adjust the threshold value as needed
            percentChange = yearlyChange / openingPrice
        Else
            percentChange = 0
        End If

        ' Write the data to the current sheet
        ws.Cells(i, summaryColumn).Value = ticker
        ws.Cells(i, summaryColumn + 1).Value = yearlyChange
        ws.Cells(i, summaryColumn + 2).Value = percentChange
        ws.Cells(i, summaryColumn + 3).Value = totalVolume

        ' Apply conditional formatting to yearly change
        If yearlyChange >= 0 Then
            ws.Cells(i, summaryColumn + 1).Interior.Color = RGB(0, 255, 0) ' Green
        Else
            ws.Cells(i, summaryColumn + 1).Interior.Color = RGB(255, 0, 0) ' Red
        End If
    Next i

    ' Find the row with the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseRow As Long
    Dim maxPercentDecreaseRow As Long
    Dim maxTotalVolumeRow As Long

    maxPercentIncrease = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, summaryColumn + 2), ws.Cells(lastRow, summaryColumn + 2)))
    maxPercentDecrease = Application.WorksheetFunction.Min(ws.Range(ws.Cells(2, summaryColumn + 2), ws.Cells(lastRow, summaryColumn + 2)))
    maxTotalVolume = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, summaryColumn + 3), ws.Cells(lastRow, summaryColumn + 3)))

    ' Find the row numbers associated with the max values
    For i = 2 To lastRow
        If ws.Cells(i, summaryColumn + 2).Value = maxPercentIncrease Then
            maxPercentIncreaseRow = i
        End If
        If ws.Cells(i, summaryColumn + 2).Value = maxPercentDecrease Then
            maxPercentDecreaseRow = i
        End If
        If ws.Cells(i, summaryColumn + 3).Value = maxTotalVolume Then
            maxTotalVolumeRow = i
        End If
    Next i

    ' Write the "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume" on the current sheet
    ws.Cells(2, summaryColumn + 5).Value = "Greatest % Increase"
    ws.Cells(2, summaryColumn + 6).Value = "Greatest % Decrease"
    ws.Cells(2, summaryColumn + 7).Value = "Greatest Total Volume"

    ws.Cells(3, summaryColumn + 5).Value = ws.Cells(maxPercentIncreaseRow, summaryColumn).Value
    ws.Cells(3, summaryColumn + 5).NumberFormat = "0.00%"
    ws.Cells(3, summaryColumn + 6).Value = ws.Cells(maxPercentDecreaseRow, summaryColumn).Value
    ws.Cells(3, summaryColumn + 6).NumberFormat = "0.00%"
    ws.Cells(3, summaryColumn + 7).Value = ws.Cells(maxTotalVolumeRow, summaryColumn).Value

End Sub

