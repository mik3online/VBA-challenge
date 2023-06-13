Attribute VB_Name = "Module1"
Sub CalculateYearlyData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Set column titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Find the last row of data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Initialize summary variables
        summaryRow = 2
        ticker = ws.Cells(2, 1).Value
        openingPrice = ws.Cells(2, 3).Value
        totalVolume = 0

        ' Loop through each row of data
        For i = 2 To lastRow
            ' Check if ticker symbol has changed
            If ws.Cells(i + 1, 1).Value <> ticker Then
                ' Get the closing price
                closingPrice = ws.Cells(i, 6).Value

                ' Calculate yearly change and percent change
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice

                ' Output the results in respective columns
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume

                ' Format percent change as percentage
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"

                ' Apply conditional formatting to highlight positive and negative changes
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If

                ' Reset summary variables for the next ticker symbol
                summaryRow = summaryRow + 1
                ticker = ws.Cells(i + 1, 1).Value
                openingPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
            End If

            ' Add the volume to the total
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        Next i

        ' Autofit columns for better visibility
        ws.Columns.AutoFit
    Next ws
End Sub

