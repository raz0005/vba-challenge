Sub StockAnalysis()

    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    
    ' Initialize total volume variable
    totalVolume = 0
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        openingPrice = ws.Cells(2, 3).Value
        ' Set initial values for summary
        summaryRow = 2
        ws.Cells(summaryRow, 9).Value = "Ticker"
        ws.Cells(summaryRow, 10).Value = "Yearly Change"
        ws.Cells(summaryRow, 11).Value = "Percent Change"
        ws.Cells(summaryRow, 12).Value = "Total Volume"
        
        ' Loop through rows to analyze data
        For i = 2 To lastRow
            
            ' Check if the ticker symbol changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Get ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Get opening and closing prices
        
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly change
                yearlyChange = closingPrice - openingPrice
                
                ' Calculate percent change
                If openingPrice <> 0 Then
                    percentChange = (yearlyChange / openingPrice)
                Else
                    percentChange = 0
                End If
                
                ' Get total stock volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Output results
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Move to the next row for summary
                summaryRow = summaryRow + 1
                
                ' Reset total volume for the next ticker
                totalVolume = 0

        ' Reset the openprice for next ticker
                openingPrice = ws.Cells(i + 1, 3).Value
                
            Else
                ' Accumulate total volume for the same ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        Dim rg As Range
        Dim cond As FormatCondition

'specify range to apply conditional formatting
        Set rg = ws.Range("J2:J" & summaryRow)

'clear any existing conditional formatting
rg.FormatConditions.Delete

'apply conditional formatting
Set cond = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0")
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")

'define conditional formatting to use
    With cond
    .Interior.Color = vbGreen
    End With
        
    With cond2
    .Interior.Color = vbRed
    End With
        
'greatest percent increase, decrease and ticker
        Dim gpinc, gpdec As Double
        Dim gtvol As LongLong
        Dim gpinc_ticker, gpdec_ticker, gtvol_ticker As String
        
        
        gpinc = ws.Cells(2, 11).Value
        gpdec = ws.Cells(2, 11).Value
        gtvol = ws.Cells(2, 12).Value
        
        gpinc_ticker = ws.Cells(2, 9).Value
        gpdec_ticker = ws.Cells(2, 9).Value
        gtvol_ticker = ws.Cells(2, 9).Value
        
        
        
        For j = 2 To summaryRow
        
                If ws.Cells(j, 11).Value > gpinc Then
                    gpinc = ws.Cells(j, 11).Value
                    gpinc_ticker = ws.Cells(j, 9).Value
                End If
                
                If ws.Cells(j, 11).Value < gpdec Then
                    gpdec = ws.Cells(j, 11).Value
                    gpdec_ticker = ws.Cells(j, 9).Value
                End If
                
                If ws.Cells(j, 12).Value > gtvol Then
                    gtvol = ws.Cells(j, 12).Value
                    gtvol_ticker = ws.Cells(j, 9).Value
                End If
        
        Next j
        
        ws.Cells(2, 14).Value = gpinc
        ws.Cells(2, 15).Value = gpinc_ticker
        ws.Cells(3, 14).Value = gpdec
        ws.Cells(3, 15).Value = gpdec_ticker
        ws.Cells(4, 14).Value = gtvol
        ws.Cells(4, 15).Value = gtvol_ticker
        
        
    Next ws

End Sub

