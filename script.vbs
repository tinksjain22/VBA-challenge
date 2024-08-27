Sub CalculateStockMetrics()

    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim initialOpeningValue As Double
    Dim totalVolume As Double
    Dim resultRow As Long
    Dim i As Long
    Dim lastRowPercentChange As Long
    Dim maxPercentChange As Double
    Dim minPercentChange As Double
    Dim maxTotalVolume As Double
    Dim tickerMaxPercentChange As String
    Dim tickerMaxTotalVolume As String
    Dim tickerMinPercentChange As String

    On Error GoTo ErrorHandler

    For Each ws In Worksheets

        ' Determine the last row with data in column A
        lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
        
        ' Filter distinct tickers and copy them to column I
        ws.Range("A1:A" & lastRow).AdvancedFilter _
            Action:=xlFilterCopy, _
            CopyToRange:=ws.Range("I1"), _
            Unique:=True
        
        ' Print headers for the results
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ' Initialize variables for processing
        initialOpeningValue = ws.Cells(2, 3).Value
        totalVolume = 0
        resultRow = 2
        
        ' Initialize values for summary calculations
        maxPercentChange = -1E+308 ' Smallest possible value
        minPercentChange = 1E+308  ' Largest possible value
        maxTotalVolume = -1E+308
        tickerMaxPercentChange = ""
        tickerMinPercentChange = ""
        tickerMaxTotalVolume = ""

        ' Calculate Quarterly Change, Percent Change, and Total Stock Volume
        For i = 2 To lastRow

            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                ' Accumulate the volume for the same ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value

            Else
                ' Calculate and record Quarterly Change and Percent Change
                ws.Cells(resultRow, 10).Value = ws.Cells(i, 6).Value - initialOpeningValue
                ws.Cells(resultRow, 11).Value = ws.Cells(resultRow, 10).Value / initialOpeningValue
                ws.Cells(resultRow, 11).NumberFormat = "0.00%"

                ' Apply conditional formatting based on Quarterly Change
                If ws.Cells(resultRow, 10).Value > 0 Then
                    ws.Cells(resultRow, 10).Interior.Color = vbGreen
                ElseIf ws.Cells(resultRow, 10).Value < 0 Then
                    ws.Cells(resultRow, 10).Interior.Color = vbRed
                Else
                    ws.Cells(resultRow, 10).Interior.Color = xlNone
                End If
                
                ' Update the opening value and total volume for the next ticker
                initialOpeningValue = ws.Cells(i + 1, 3).Value
                ws.Cells(resultRow, 12).Value = totalVolume + ws.Cells(i, 7).Value
                totalVolume = 0
                resultRow = resultRow + 1

                ' Update summary values if necessary
                If ws.Cells(resultRow - 1, 11).Value > maxPercentChange Then
                    maxPercentChange = ws.Cells(resultRow - 1, 11).Value
                    tickerMaxPercentChange = ws.Cells(resultRow - 1, 9).Value
                End If
                
                If ws.Cells(resultRow - 1, 11).Value < minPercentChange Then
                    minPercentChange = ws.Cells(resultRow - 1, 11).Value
                    tickerMinPercentChange = ws.Cells(resultRow - 1, 9).Value
                End If
                
                If ws.Cells(resultRow - 1, 12).Value > maxTotalVolume Then
                    maxTotalVolume = ws.Cells(resultRow - 1, 12).Value
                    tickerMaxTotalVolume = ws.Cells(resultRow - 1, 9).Value
                End If

            End If

        Next i
        
        ' Place the summary values in the worksheet
        ws.Cells(2, 16).Value = tickerMaxPercentChange
        ws.Cells(2, 17).Value = Format(maxPercentChange, "0.00%")
        ws.Cells(3, 16).Value = tickerMinPercentChange
        ws.Cells(3, 17).Value = Format(minPercentChange, "0.00%")
        ws.Cells(4, 16).Value = tickerMaxTotalVolume
        ws.Cells(4, 17).Value = maxTotalVolume

    Next ws

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    Resume Next

End Sub


