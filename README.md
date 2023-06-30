Sub StockAnalysis()
    Dim ws As Worksheet
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ' Find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'Summary data
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize summary row
        summaryRow = 2
        
        ' Loop through the rows of data
        For i = 2 To lastRow
            ' Check if we are still within the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Set the closing price
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change
                yearlyChange = closingPrice - openingPrice
                
                ' Calculate the percent change
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                
                ' Summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Format the percent change as a percentage
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                
                ' Conditional formatting
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                ' Reset for the next ticker
                openingPrice = 0
                totalVolume = 0
                
                ' summary row
                summaryRow = summaryRow + 1
            End If
            
            '  first record for the ticker
            If openingPrice = 0 Then
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            '  total stock volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        Next i
        
        ' greatest increase
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If
                
                ' greatest decrease
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
            End If
            
End Sub
