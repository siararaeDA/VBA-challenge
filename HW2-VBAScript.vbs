Sub tickerOutput():

    ' Declare variables
    Dim ticker As String
    Dim lastRow As Long
    Dim outputRow As Long
    
    ' Initialize variables
    outputRow = 2

    ' Set the column header
    Range("I1").Value = "Ticker"
    Range("I1").Font.Bold = True
    
    ' Find the end of the data
    ' Source of the function: https://www.excelcampus.com/vba/find-last-row-column-cell/#endcode
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through Column A and compare values to find each new ticker
    For i = 2 To lastRow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            Cells(outputRow, 9).Value = ticker
            outputRow = outputRow + 1
        End If
    
    Next i

End Sub

Sub yearPriceDifferential():

    ' Declare Variables
    Dim priceOpen As Double
    Dim priceClose As Double
    Dim priceChange As Double
    Dim ticker As String
    Dim lastRow As Long
    Dim outputRow As Long

    ' Initialize variables
    outputRow = 2
    priceOpen = Range("C2").Value
    priceClose = 0
    priceChange = 0
    
    ' Set the column header
    Range("J1").Value = "Yearly Change"
    Range("J1").Font.Bold = True

    ' Find the end of the data
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through tickers to find starting and closing prices, and the price differential
    For i = 2 To lastRow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Since we've reached the end of a ticker's values and that matches the end of the year, get the closing price.
            priceClose = Cells(i, 6).Value
            ' Find the change in price over the year
            priceChange = priceClose - priceOpen
            ' Write the change in price in the output table
            Cells(outputRow, 10).Value = priceChange
            ' Format the output
            If priceChange > 0 Then
                Cells(outputRow, 10).Interior.ColorIndex = 4
            Else
                Cells(outputRow, 10).Interior.ColorIndex = 3
            End If
            ' Get the opening price of the next stock
            priceOpen = Cells(i + 1, 3).Value
            ' Increment the output row for the next stock
            outputRow = outputRow + 1
        End If
    
    Next i

End Sub
