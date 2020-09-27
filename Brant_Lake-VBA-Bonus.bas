Attribute VB_Name = "Module2"
Sub bonus()

For Each ws In Worksheets
    ws.Activate

'create headers for summary table

    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'identify the last row again
    
    lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'define some increase variables
    
    Dim greatest_increase As Double
    Dim increase_ticker As String
    Dim greatest_decrease As Double
    Dim decrease_ticker As String
    Dim greatest_volume As Double
    Dim volume_ticker As String
    
    'set variables
    
    greatest_increase = Cells(2, 11).Value
    increase_ticker = Cells(2, 9).Value
    greatest_decrease = Cells(2, 11).Value
    decrease_ticker = Cells(2, 9).Value
    greatest_volume = Cells(2, 12).Value
    volume_ticker = Cells(2, 9).Value
    
    'loop through list of tickers
    
    For i = 2 To lastrow
    
        'find greatest percent increase
        If Cells(i, 11).Value > greatest_increase Then
            greatest_increase = Cells(i, 11).Value
            increase_ticker = Cells(i, 9).Value
        End If
        
        'find greatest percent decrease
        If Cells(i, 11).Value < greatest_decrease Then
            greatest_decrease = Cells(i, 11).Value
            decrease_ticker = Cells(i, 9).Value
        End If
        
        'find greatest total volume
        If Cells(i, 12).Value > greatest_volume Then
            greatest_volume = Cells(i, 12).Value
            volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    'add values into summary table
    
    Cells(2, 17).Value = Format(greatest_increase, "Percent")
    Cells(3, 17).Value = Format(greatest_decrease, "Percent")
    Cells(4, 17).Value = greatest_volume
    Cells(2, 16).Value = Format(increase_ticker, "Percent")
    Cells(3, 16).Value = Format(decrease_ticker, "Percent")
    Cells(4, 16).Value = volume_ticker

Next ws
    



End Sub
