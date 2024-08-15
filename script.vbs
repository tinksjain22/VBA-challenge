Sub Sum()

'Declaring variables
  Dim ws As Worksheet
  Dim Lastrow As Long
  Dim Opening_Value As Double
  Dim Count As Double
  Dim Row As Integer
 
For Each ws In Worksheets

'Initializaing Variables
      Lastrow = ws.Range("A" & Rows.Count).End(xlUp).Row

'Filtering Distinct tickers

      ws.Range("A1:A" & Lastrow).AdvancedFilter _
      Action:=xlFilterCopy, _
      CopyToRange:=ws.Range("I1"), _
      Unique:=True
     
     
'Printing Headers

     ws.Cells(1, 9).Value = "Ticker"
     ws.Cells(1, 10).Value = "Quaterly Change"
     ws.Cells(1, 11).Value = "Percent Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"
     ws.Cells(2, 15).Value = "Greatest % Increase"
     ws.Cells(3, 15).Value = "Greatest % Decrease"
     ws.Cells(4, 15).Value = "Greatest Total Volume"
   
   
 'Initializaing Variables
 
     Opening_Value = ws.Cells(2, 3).Value
     Count = 0
     Row = 2
     
 'Populating Values for Quarterly Change, Percent Change and Total stock
 
     For i = 2 To Lastrow
 
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
             
                 Count = Count + ws.Cells(i, 7).Value

 
            Else
                 ws.Cells(Row, 10).Value = ws.Cells(i, 6).Value - Opening_Value
                 ws.Cells(Row, 11).Value = FormatPercent(ws.Cells(Row, 10).Value / Opening_Value)
                 
'Conditional Formatting cells

                 If ws.Cells(Row, 10).Value > 0 Then
                 
                         ws.Cells(Row, 10).Interior.Color = vbGreen
                         
                 ElseIf ws.Cells(Row, 10).Value < 0 Then
                         
                         ws.Cells(Row, 10).Interior.Color = vbRed
                         
                 Else
                         ws.Cells(Row, 10).Interior.Color = xlNone
                 End If
 
                 Opening_Value = ws.Cells(i + 1, 3).Value
                 ws.Cells(Row, 12).Value = Count + ws.Cells(i, 7).Value
                 Count = 0
                 Row = Row + 1

             End If
             
 'Calculated Values for Summary
 
             ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
             ws.Cells(2, 17).Value = FormatPercent(Application.WorksheetFunction.Max(ws.Range("K:K")))
             ws.Cells(3, 17).Value = FormatPercent(Application.WorksheetFunction.Min(ws.Range("K:K")))
 
     Next
     
 'Calculated Ticker Value for Summary
 
   Lastrow_Percent_Change = ws.Range("K" & Rows.Count).End(xlUp).Row

     
   For i = 2 To Lastrow_Percent_Change
           
             If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
                   
                   ws.Cells(2, 16).Value = ws.Cells(i, 9)
                   
             ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L:L")) Then
                   
                    ws.Cells(4, 16).Value = ws.Cells(i, 9)
                   
             ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
                   
                    ws.Cells(3, 16).Value = ws.Cells(i, 9)
                   
       
            End If
   Next
   
 
 Next
 
 End Sub



