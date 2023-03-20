Sub YearlyChangeFinal()
'Loop through all sheets
    For Each ws In Worksheets

        'Add all my headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(, 12).Value = " Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        'Declaring all my variables
        Dim tickerName As String
        Dim firstRow As Integer
        Dim i, j As Integer
        Dim stockVol As LongLong
        Dim openYear As Double
        Dim closeYear As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVol As LongLong
        

        'Initialize the variables
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        firstRow = 2
        stockVol = 0
        openYear = ws.Cells(2, 3).Value
        closeYear = 0
        percentChange = 0
        maxIncrease = 0
        maxDecrease = 0
        maxVol = 0
        

        'Start a for Loop to find and print tickers and associated data
        For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            tickerName = (ws.Cells(i, 1).Value)
            stockVol = stockVol + (ws.Cells(i, 7).Value)
            closeYear = closeYear + (ws.Cells(i, 6).Value)
            yearlyChange = closeYear - (openYear)
            percentChange = ((closeYear - openYear) / openYear) * 100
            
            'Adding Location of my variables
            ws.Cells(firstRow, 9).Value = tickerName
            ws.Cells(firstRow, 12).Value = stockVol
            ws.Cells(firstRow, 10).Value = yearlyChange
            ws.Cells(firstRow, 11).Value = percentChange
            
            'Add to firstRow
            firstRow = firstRow + 1
            
            'Reset stockVol
            stockVol = 0
            
            'Redefine openYear
            openYear = ws.Cells(i + 1, 3).Value
            
            'Reset closeYear
            closeYear = 0
            
            'Reset percentChange
            percentChange = 0
            
        Else
            'Add to stockVol
            stockVol = stockVol + (ws.Cells(i, 7).Value)
            
        End If
        Next i

        'Format the columns Fill Color according to change
        For i = 2 To lastRow
        If ws.Cells(i, 11).Value > 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        ElseIf ws.Cells(i, 11).Value < 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        Next i
    

        
        'This piece will find Max Increase and Max Decrease
        For i = 2 To lastRow
        If (ws.Cells(i, 11).Value) > maxIncrease Then
        'Print in new field
                maxIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 17).Value = maxIncrease
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
        ElseIf (ws.Cells(i, 11).Value) < maxDecrease Then
            'Print in new field
                maxDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 17).Value = maxDecrease
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
        Else
        End If
        
        Next i

        'Now I will go to the maxStock
        For i = 2 To lastRow
            If (ws.Cells(i, 12).Value > maxVol) Then
            'Print it in new field
                maxVol = ws.Cells(i, 12).Value
                ws.Cells(4, 17).Value = maxVol
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
            Else
            End If

        Next i

        'Format percentage columns
        ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 18).NumberFormat = "0.00%"
        
    Next ws

End Sub
