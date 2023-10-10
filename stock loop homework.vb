Sub stockloops()
    'create worksheet variable
    Dim ws As Worksheet
    'Loop through all sheets
    For Each ws In Worksheets
    
        'Create column/row values for new data
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Volume"
        
        'Determine the last row
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set initial variable for ticker, opening price, closing price, yearly change, % change
        Dim ticker As String
        Dim opening as double
        Dim closing as double
        Dim yearly_change as double
        Dim percent_change as double
        
        'Set variables for greatest values
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_volume As Double
        Dim greatest_increase_ticker As String
        Dim greatest_decrease_ticker As String
        Dim greatest_volume_ticker As String
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0
 
        'set initial variable to hold total stock volume
        Dim volume As Double
        volume = 0
        
        'Keep track of Summary Table row
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'create loop for each sheet
        For i = 2 To lastrow
            'run a loop to get opening value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                'set opening value
                opening = ws.Cells(i, 3).Value
            End If
            
            'create loop to get closing value & all other relevant data
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Set the ticker
                ticker = ws.Cells(i, 1).Value
                'set closing value
                closing = ws.Cells(i, 6).Value
                'add to the total stock volume
                volume = volume + ws.Cells(i, 7).Value
                'print the ticker in the summary table
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                'print the total volume in the summary table
                ws.Cells(Summary_Table_Row, 12).Value = volume
                'calculate yearly change
                yearly_change = closing - opening
                'put yearly change into the summary table
                ws.Cells(Summary_Table_Row, 10).Value = yearly_change
                'calculate % change
                percent_change = (yearly_change / opening)
                'put % change into the summary table
                ws.Cells(Summary_Table_Row, 11).Value = FormatPercent(percent_change)
                'add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                'reset the total stock volume
                volume = 0
                
            Else
                'add to the total stock volume
                volume = volume + ws.Cells(i, 7).Value
            End If
            
            'create loop for conditional formatting
            'color negative value ws.Cells red
            If ws.Cells(i, 10) < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            'color positive value ws.Cells green
            ElseIf ws.Cells(i, 10) > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
            
            'find greatest increase
            If ws.Cells(i, 11).Value > greatest_increase Then
                greatest_increase = ws.Cells(i, 11).Value
                greatest_increase_ticker = ws.Cells(i, 9).Value
            End If
            
            'find greatest decrease
            If ws.Cells(i, 11).Value < greatest_decrease Then
                greatest_decrease = ws.Cells(i, 11).Value
                greatest_decrease_ticker = ws.Cells(i, 9).Value
            End If
            
            'find greatest volume
            If ws.Cells(i, 12).Value > greatest_volume Then
                greatest_volume = ws.Cells(i, 12).Value
                greatest_volume_ticker = ws.Cells(i, 9).Value
            End If
            
        Next i
        
        ' Output greatest values and corresponding tickers
        ws.Cells(2, 16).Value = greatest_increase_ticker
        ws.Cells(2, 17).Value = FormatPercent(greatest_increase)
    
        ws.Cells(3, 16).Value = greatest_decrease_ticker
        ws.Cells(3, 17).Value = FormatPercent(greatest_decrease)
      
        ws.Cells(4, 16).Value = greatest_volume_ticker
        ws.Cells(4, 17).Value = greatest_volume
        
    Next ws
    
End Sub
