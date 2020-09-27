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
