Sub stocks():

' Loop through worksheets
For Each ws In Worksheets

' Part 1 - First summary Table
Dim i, LastRow, TableRow As Integer
Dim TotalStock As LongLong
Dim Current, Previous As String
Dim OpenPrice, ClosePrice, YearlyChange, PercentChange As Double
    'Determine number of rows
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' set starting point for summary table row
        TableRow = 2
    ' Set starting point for total stock volume as second row
        TotalStock = ws.Range("G2")
     'Set starting point for openning price at begining of year
        OpenPrice = ws.Range("C2").Value
    ' Creating headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume" 
    'Loop through rows starting from row 2 to the second to last row
    For i = 2 To LastRow - 1
    ' Set values for current and next ticker symbols
        CurrentTicker = ws.Cells(i, 1).Value
        NextTicker = ws.Cells(i + 1, 1).Value
    ' If current ticker and next one are the same add stock volume
        If CurrentTicker = NextTicker Then
            TotalStock = TotalStock + ws.Range("G" & i + 1).Value
    ' If curret ticker and next one are different (i.e., stock for a new year) - write into summary table
        Else
        ' calculate yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
           ClosePrice = ws.Range("F" & i).Value
           YearlyChange = ClosePrice - OpenPrice
        ' calculate the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
           PercentChange = YearlyChange / OpenPrice
        ' Add values to summary table
            ws.Range("I" & TableRow).Value = CurrentTicker
            ws.Range("J" & TableRow).Value = YearlyChange
            ws.Range("K" & TableRow).Value = PercentChange
            ws.Range("L" & TableRow).Value = TotalStock
        ' Formatting summary table
            If YearlyChange < 0 Then
                ws.Range("J" & TableRow).Interior.ColorIndex = 3
            Else
                ws.Range("J" & TableRow).Interior.ColorIndex = 4
            End If
            ws.Range("K" & TableRow).NumberFormat = "%0.00"
        ' re-start the open value and the total stock with the current next values
            TotalStock = ws.Range("G" & i + 1).Value
            OpenPrice = ws.Range("C" & i + 1).Value
        ' advance summary table row
            TableRow = TableRow + 1
        End If
      Next i
      
' Part 2 - Second summary table
Dim LastTableRow As Integer
Dim GreatVolume, NextVolume As LongLong
Dim GreatIncrease, GreatDecrease As Double
Dim Ticker_D, Ticker_I, Ticker_V As String
    ' Constract new table headers
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    'Determine number of rows
    LastTableRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    ' set starting points for percent change and volume
    GreatVolume = ws.Range("L2").Value
    GreatDecrease = ws.Range("K2").Value
    GreatIncrease = ws.Range("K2").Value
    Ticker_I = ws.Range("I2").Value
    Ticker_D = ws.Range("I2").Value
    'Loop through summary table
    For i = 3 To LastTableRow
        PercentChange = ws.Range("K" & i).Value
        'Greatest decrease check
        If PercentChange < GreatDecrease Then
                GreatDecrease = PercentChange
                Ticker_D = ws.Range("I" & i)
        'Greatest increase check
        ElseIf PercentChange > GreatIncrease Then
                GreatIncrease = PercentChange
                Ticker_I = ws.Range("I" & i)
        End If
        'Greatest Volume check
        NextVolume = ws.Range("L" & i + 1).Value
        If GreatVolume < NextVolume Then
            GreatVolume = NextVolume
            Ticker_V = ws.Range("I" & i + 1)
        End If
    Next i
    'Write values to table
    ws.Range("P2").Value = Ticker_I
    ws.Range("P3").Value = Ticker_D
    ws.Range("P4").Value = Ticker_V
    ws.Range("Q2").Value = GreatIncrease
    ws.Range("Q3").Value = GreatDecrease
    ws.Range("Q2:Q3").NumberFormat = "%0.00"
    ws.Range("Q4").Value = GreatVolume
    ' Format columns width
    ws.Columns("L").ColumnWidth = 15
    ws.Columns("O").ColumnWidth = 20
    ws.Columns("Q").ColumnWidth = 15

Next ws
End Sub