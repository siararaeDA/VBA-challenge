Sub findAndFormatStockData():

    ' Declare variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim priceOpen As Double
    Dim priceClose As Double
    Dim priceChange As Double
    Dim percentChange As Double
    Dim stockVolume As Double
    Dim lastRow As Long
    Dim outputRow As Long
    
    For Each ws In Worksheets

        ' Initialize variables
        ticker = ""
        outputRow = 2
        priceOpen = Range("C2").Value
        priceClose = 0
        priceChange = 0
        percentChange = 0
        stockVolume = 0

        ' Set the column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("I1").Font.Bold = True
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("J1").Font.Bold = True
        ws.Range("K1").Value = "Percent Change"
        ws.Range("K1").Font.Bold = True
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("L1").Font.Bold = True

        ' Find the end of the data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through tickers to find starting and closing prices, and the price differential
        For i = 2 To lastRow

            ' Keep a tally of the stock volume
            stockVolume = stockVolume + ws.Cells(i, 7).Value

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Get the ticker value
                ticker = ws.Cells(i, 1).Value
                'Output
                ws.Cells(outputRow, 9).Value = ticker
                ' Since we've reached the end of a ticker's values and that matches the end of the year, get the closing price.
                priceClose = ws.Cells(i, 6).Value
                ' Find the change in price over the year
                priceChange = priceClose - priceOpen
                ' Write the change in price in the output table
                ws.Cells(outputRow, 10).Value = priceChange
                ' Format the output
                If priceChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                End If
                ' Get the percent change over the year
                If priceOpen = 0 Then
                    ws.Cells(outputRow, 11).Value = 0
                Else
                    percentChange = (priceChange / priceOpen)
                    ' Write the change percentage in the output
                    ws.Cells(outputRow, 11).Value = Format(percentChange, "Percent")
                End If

                ' Write the stock volume in the output
                ws.Cells(outputRow, 12).Value = stockVolume
                ' Get the opening price of the next stock
                priceOpen = ws.Cells(i + 1, 3).Value
                ' Increment the output row for the next stock
                outputRow = outputRow + 1
                ' Reset stock volume to zero for next stock
                stockVolume = 0
            End If

        Next i

        ' Call the bonus functions in each worksheet
        ws.Select
        greatestPercentIncrease ws
        greatestPercentDecrease ws
        highestVolume ws

    Next ws

End Sub

Sub greatestPercentIncrease(ws As Worksheet):

    Dim currentNum As Double
    Dim nextNum As Double
    Dim maxIncrease As Double
    Dim maxIncreaseTicker As String
    Dim lastRow As Double
    
    currentNum = 0
    nextNum = 0
    maxIncrease = 0
    
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'MsgBox (ws.Cells(2, 11).Value)
    
    For i = 2 To lastRow
    
        'MsgBox (ws.Cells(i, 11).Value)
    
        currentNum = ws.Cells(i, 11).Value
        If currentNum > maxIncrease Then
            maxIncrease = currentNum
            maxIncreaseTicker = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    ws.Range("O2").Value = "Greatest Percent Increase:"
    ws.Range("P2").Value = maxIncreaseTicker
    ws.Range("Q2").Value = Format(maxIncrease, "Percent")

End Sub

Sub greatestPercentDecrease(ws As Worksheet):

    Dim currentNum As Double
    Dim nextNum As Double
    Dim maxDecrease As Double
    Dim maxDecreaseTicker As String
    Dim lastRow As Double
    
    currentNum = 0
    nextNum = 0
    maxDecrease = 0
    
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'MsgBox (Cells(2, 11).Value)
    
    For i = 2 To lastRow
    
        currentNum = ws.Cells(i, 11).Value
        If currentNum < maxDecrease Then
            maxDecrease = currentNum
            maxDecreaseTicker = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    ws.Range("O3").Value = "Greatest Percent Decrease:"
    ws.Range("P3").Value = maxDecreaseTicker
    ws.Range("Q3").Value = Format(maxDecrease, "Percent")

End Sub

Sub highestVolume(ws As Worksheet):

    Dim currentNum As Double
    Dim nextNum As Double
    Dim maxVolume As Double
    Dim maxVolumeTicker As String
    Dim lastRow As Double
    
    currentNum = 0
    nextNum = 0
    maVolume = 0
    
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'MsgBox (Cells(2, 11).Value)
    
    For i = 2 To lastRow
    
        currentNum = ws.Cells(i, 12).Value
        If currentNum > maxVolume Then
            maxVolume = currentNum
            maxVolumeTicker = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    ws.Range("O4").Value = "Greatest Total Volume:"
    ws.Range("P4").Value = maxVolumeTicker
    ws.Range("Q4").Value = maxVolume

End Sub
