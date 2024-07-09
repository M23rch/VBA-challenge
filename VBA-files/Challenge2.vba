Attribute VB_Name = "Module1"
Sub Ticker_Insert()
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim OutputRow As Long
    Dim i As Long

    sheetNames = Array("Q1", "Q2", "Q3", "Q4")

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        ws.Cells(1, 9).Value = "Ticker" ' Label for Ticker Column

        OutputRow = 2 ' Initialize output row for each worksheet

        For i = 2 To 93001 Step 62
            ws.Cells(OutputRow, 9).Value = ws.Cells(i, 1).Value
            OutputRow = OutputRow + 1
        Next i
    Next sheetName
End Sub

Sub Quarterly_Change()
    Dim ws As Worksheet
    Dim RowCount As Long
    Dim i As Long
    Dim ClosingPrice As Double
    Dim OpeningPrice As Double
    Dim QuarterlyChange As Double
    Dim StartRow As Long
    Dim OutputRow As Long
    Dim sheetNames As Variant
    Dim sheetName As Variant

    sheetNames = Array("Q1", "Q2", "Q3", "Q4")

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        StartRow = 2 ' Initialize start row for each worksheet
        OutputRow = 2 ' Initialize output row for each worksheet

        ' Set column headers
        ws.Cells(1, 10).Value = "Quarterly Change"

        For i = 2 To RowCount
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = RowCount Then
                ' Capture the opening price at the beginning of the quarter
                OpeningPrice = ws.Cells(StartRow, 3).Value
                
                ' Capture the closing price at the end of the quarter
                ClosingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the quarterly change
                QuarterlyChange = ClosingPrice - OpeningPrice
                
                ' Output the result in the columns I and J
                ws.Cells(OutputRow, 9).Value = ws.Cells(StartRow, 1).Value ' Ticker Symbol in Column I
                ws.Cells(OutputRow, 10).Value = QuarterlyChange ' Quarterly Change in Column J

                ' Apply conditional formatting based on the value of QuarterlyChange
                If QuarterlyChange < 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                ElseIf QuarterlyChange > 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                Else
                    ws.Cells(OutputRow, 10).Interior.ColorIndex = xlNone ' No color for zero change
                End If
                
                ' Move to the next output row
                OutputRow = OutputRow + 1
                
                ' Reset the start row for the next ticker
                StartRow = i + 1
            End If
        Next i
    Next sheetName
End Sub


Sub Percentage_Change()
    Dim ws As Worksheet
    Dim RowCount As Long
    Dim i As Long
    Dim ClosingPrice As Double
    Dim OpeningPrice As Double
    Dim PercentageChange As Double
    Dim StartRow As Long
    Dim OutputRow As Long
    Dim sheetNames As Variant
    Dim sheetName As Variant

    sheetNames = Array("Q1", "Q2", "Q3", "Q4")

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        StartRow = 2 ' Initialize start row for each worksheet
        OutputRow = 2 ' Initialize output row for each worksheet

        ' Set column headers
        ws.Cells(1, 11).Value = "Percentage Change"

        For i = 2 To RowCount
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = RowCount Then
                ' Capture the opening price at the beginning of the quarter
                OpeningPrice = ws.Cells(StartRow, 3).Value
                
                ' Capture the closing price at the end of the quarter
                ClosingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the percentage change
                If OpeningPrice <> 0 Then
                    PercentageChange = ((ClosingPrice - OpeningPrice) / OpeningPrice) * 100
                Else
                    PercentageChange = 0 ' Handle division by zero if opening price is zero
                End If
                
                ' Output the result in the columns I and K
                ws.Cells(OutputRow, 9).Value = ws.Cells(StartRow, 1).Value ' Ticker Symbol in Column I
                ws.Cells(OutputRow, 11).Value = PercentageChange ' Percentage Change in Column K

                ' Move to the next output row
                OutputRow = OutputRow + 1
                
                ' Reset the start row for the next ticker
                StartRow = i + 1
            End If
        Next i
    Next sheetName
End Sub

Sub Total_Volume()
    Dim ws As Worksheet
    Dim RowCount As Long
    Dim i As Long
    Dim TotalVolume As Double
    Dim TickerSymbol As String
    Dim StartRow As Long
    Dim OutputRow As Long
    Dim sheetNames As Variant
    Dim sheetName As Variant

    sheetNames = Array("Q1", "Q2", "Q3", "Q4")

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        StartRow = 2 ' Initialize start row for each worksheet
        OutputRow = 2 ' Initialize output row for each worksheet

        ' Set column headers
        ws.Cells(1, 12).Value = "Total Volume" ' Column L (12th column)

        ' Reset total volume and ticker symbol
        TotalVolume = 0
        TickerSymbol = ws.Cells(StartRow, 1).Value

        For i = StartRow To RowCount
            ' Sum the volume for the current ticker
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value ' Assuming volume is in column G (7th column)
            
            ' Check if the ticker changes or it is the last row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = RowCount Then
                ' Output the ticker and total volume
                ws.Cells(OutputRow, 9).Value = TickerSymbol ' Ticker Symbol in Column I
                ws.Cells(OutputRow, 12).Value = TotalVolume ' Total Volume in Column L (12th column)

                ' Move to the next output row
                OutputRow = OutputRow + 1
                
                ' Reset total volume for the next ticker symbol
                TotalVolume = 0
                
                ' Update ticker symbol
                If i <> RowCount Then
                    TickerSymbol = ws.Cells(i + 1, 1).Value
                End If
            End If
        Next i
    Next sheetName
End Sub

Sub Quarterly_Analysis()
    Dim ws As Worksheet
    Dim RowCount As Long
    Dim i As Long
    Dim ClosingPrice As Double
    Dim OpeningPrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim TickerIncrease As String
    Dim TickerDecrease As String
    Dim TickerVolume As String
    Dim StartRow As Long
    Dim OutputRow As Long
    Dim sheetNames As Variant
    Dim sheetName As Variant

    ' Initialize variables for tracking maximum values
    MaxIncrease = -999999999
    MaxDecrease = 999999999
    MaxVolume = 0
    TickerIncrease = ""
    TickerDecrease = ""
    TickerVolume = ""

    sheetNames = Array("Q1", "Q2", "Q3", "Q4")

    For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Sheets(sheetName)
        RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        StartRow = 2 ' Initialize start row for each worksheet
        OutputRow = 2 ' Initialize output row for each worksheet

        ' Set column headers
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        For i = 2 To RowCount
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = RowCount Then
                ' Capture the opening price at the beginning of the quarter
                OpeningPrice = ws.Cells(StartRow, 3).Value
                
                ' Capture the closing price at the end of the quarter
                ClosingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the quarterly change
                QuarterlyChange = ClosingPrice - OpeningPrice
                
                ' Calculate the percent change
                If OpeningPrice <> 0 Then
                    PercentChange = QuarterlyChange / OpeningPrice
                Else
                    PercentChange = 0
                End If
                
                ' Calculate the total volume
                TotalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(StartRow, 7), ws.Cells(i, 7)))

                ' Output the results in the columns
                ws.Cells(OutputRow, 9).Value = ws.Cells(StartRow, 1).Value ' Ticker Symbol in Column I
                ws.Cells(OutputRow, 10).Value = QuarterlyChange ' Quarterly Change in Column J
                ws.Cells(OutputRow, 11).Value = PercentChange ' Percent Change in Column K
                ws.Cells(OutputRow, 12).Value = TotalVolume ' Total Stock Volume in Column L

                ' Apply conditional formatting based on the value of QuarterlyChange
                If QuarterlyChange < 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                ElseIf QuarterlyChange > 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                Else
                    ws.Cells(OutputRow, 10).Interior.ColorIndex = xlNone ' No color for zero change
                End If
                
                ' Check for greatest increase, decrease, and volume
                If PercentChange > MaxIncrease Then
                    MaxIncrease = PercentChange
                    TickerIncrease = ws.Cells(StartRow, 1).Value
                End If
                
                If PercentChange < MaxDecrease Then
                    MaxDecrease = PercentChange
                    TickerDecrease = ws.Cells(StartRow, 1).Value
                End If
                
                If TotalVolume > MaxVolume Then
                    MaxVolume = TotalVolume
                    TickerVolume = ws.Cells(StartRow, 1).Value
                End If
                
                ' Move to the next output row
                OutputRow = OutputRow + 1
                
                ' Reset the start row for the next ticker
                StartRow = i + 1
            End If
        Next i
        
        ' Output greatest % increase, % decrease, and total volume in the same worksheet
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Total Volume"
        
        ws.Cells(2, 15).Value = TickerIncrease
        ws.Cells(2, 16).Value = MaxIncrease
        ws.Cells(3, 15).Value = TickerDecrease
        ws.Cells(3, 16).Value = MaxDecrease
        ws.Cells(4, 15).Value = TickerVolume
        ws.Cells(4, 16).Value = MaxVolume
        
        ' Apply formatting to values
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
        ' Reset the tracking variables for the next sheet
        MaxIncrease = -999999999
        MaxDecrease = 999999999
        MaxVolume = 0
        TickerIncrease = ""
        TickerDecrease = ""
        TickerVolume = ""
    Next sheetName
End Sub

