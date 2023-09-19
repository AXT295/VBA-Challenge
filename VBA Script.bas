Attribute VB_Name = "Module1"

Sub tickerloop()
    Dim ticker_name As String
    
    Dim ticker_volume As Double
    ticker_volume = 0
    
    Dim summary_ticker_row As Integer
    summary_ticker_row = 2
    
    Dim open_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
'Assigng the Columns.
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Columns("J").ColumnWidth = 12
    Cells(1, 11).Value = "Percent Change"
    Columns("K").ColumnWidth = 14
    Cells(1, 12).Value = "Total Stock Volume"
    Columns("L").ColumnWidth = 16


    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        
   
'Looping through the data sheet.
    For i = 2 To lastrow
    
'Keeping track of the total volume for that specific ticker.
        ticker_volume = ticker_volume + Cells(i, 7).Value
    
'Checking the two cells are different. Getting the closing price, percentage change, yearly change and displaying it in a summary ticker row.
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            close_price = Cells(i, 6).Value

            Range("L" & summary_ticker_row).Value = ticker_volume
            
            yearly_change = (close_price - open_price)
            Range("J" & summary_ticker_row).Value = yearly_change
            Range("J" & summary_ticker_row).NumberFormat = "0.00"
            
'Just resetting the volume to 0 for the next of ticker volumes.
            ticker_volume = 0

            Range("I" & summary_ticker_row).Value = ticker_name
                        
            If open_price <> 0 Then
                percentage_change = (yearly_change / open_price)

            Else
                percentage_change = 0
            End If
            Range("K" & summary_ticker_row).Value = percentage_change
            Range("K" & summary_ticker_row).NumberFormat = "0.00%"

            
'Incrementing the rows for each ticker.
            summary_ticker_row = summary_ticker_row + 1
                
        End If
        
        
'Getting the first ticker values.
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            ticker_name = Cells(i, 1).Value
            open_price = Cells(i, 3).Value
     
        End If
                   
    Next i
    
    
'Looping through positive and negative and coloring coding it to
    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastrow_summary_table
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 10
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
'Labeling the cells.
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Columns("O").ColumnWidth = 19
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
 
'Outputing greatest percentage increase, greatest percentage decrease, and greatest total volume.
    For i = 2 To lastrow_summary_table
    
        If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = Cells(i, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
            
        ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = Cells(i, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
        End If
        
        If Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = Cells(i, 12).Value
        End If
    Next i
    
End Sub



