Attribute VB_Name = "Module2"
Sub FindGreatestValues2()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row of data
        lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Initialize variables
        maxIncrease = ws.Cells(2, 11).Value
        maxDecrease = ws.Cells(2, 11).Value
        maxVolume = ws.Cells(2, 12).Value
        tickerIncrease = ws.Cells(2, 9).Value
        tickerDecrease = ws.Cells(2, 9).Value
        tickerVolume = ws.Cells(2, 9).Value
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Check for greatest % increase
            If ws.Cells(i, 11).Value > maxIncrease Then
                maxIncrease = ws.Cells(i, 11).Value
                tickerIncrease = ws.Cells(i, 9).Value
            End If
            
            ' Check for greatest % decrease
            If ws.Cells(i, 11).Value < maxDecrease Then
                maxDecrease = ws.Cells(i, 11).Value
                tickerDecrease = ws.Cells(i, 9).Value
            End If
            
            ' Check for greatest total volume
            If ws.Cells(i, 12).Value > maxVolume Then
                maxVolume = ws.Cells(i, 12).Value
                tickerVolume = ws.Cells(i, 9).Value
            End If
        Next i
        
        ' Display results in specified cells
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("P2").Value = tickerIncrease
        ws.Range("P3").Value = tickerDecrease
        ws.Range("P4").Value = tickerVolume
        
        ws.Range("Q1").Value = "Value"
        ws.Range("Q2").Value = maxIncrease
        ws.Range("Q3").Value = maxDecrease
        ws.Range("Q4").Value = maxVolume
        
        ' Format Q2 and Q3 cells as percentages
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' Autofit columns for better visibility
        ws.Columns("O:Q").AutoFit
    Next ws
End Sub

