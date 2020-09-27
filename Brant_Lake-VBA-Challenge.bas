Attribute VB_Name = "Module1"
Sub stocks()

'define the variables
'volume variable has to be a double not integer or long or you get an overflow error


Dim ws As Worksheet
Dim ticker As String
Dim ticker_count As Integer
Dim volume As Double
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim lastrow As Long




'run the code through every worksheet in workbook

For Each ws In Worksheets
    ws.Activate
    
    'last row calculation
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'create headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    'reset variable to blank
    ticker = ""
    ticker_count = 0
    open_price = 0
    yearly_change = 0
    percent_change = 0
    volume = 0
    
    'loop for going through all tickers
    
    For i = 2 To lastrow
        
        'add value of ticker to ticker variable
        
        ticker = Cells(i, 1).Value
        
        'get coresponding opening price
        
        If open_price = 0 Then
            open_price = Cells(i, 3).Value
        End If
        
        'runs when we run into a different ticker and collects values
    
        volume = volume + Cells(i, 7).Value
        
        If Cells(i + 1, 1).Value <> ticker Then
            ticker_count = ticker_count + 1
            Cells(ticker_count + 1, 9) = ticker
            
            close_price = Cells(i, 6)
            
            
            yearly_change = close_price - open_price
            
            Cells(ticker_count + 1, 10).Value = yearly_change
            
            'set color shading of cells
            
            If yearly_change > 0 Then
                Cells(ticker_count + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(ticker_count + 1, 10).Interior.ColorIndex = 3
            Else
                Cells(ticker_count + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            'calculate percent changes
            
            If open_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / open_price)
            End If
            
            
            'format into a percentage
            Cells(ticker_count + 1, 11).Value = Format(percent_change, "Percent")
            
            'color shading of percent change
            
            If percent_change > 0 Then
                Cells(ticker_count + 1, 11).Interior.ColorIndex = 4
            ElseIf percent_change < 0 Then
                Cells(ticker_count + 1, 11).Interior.ColorIndex = 3
            Else
                Cells(ticker_count + 1, 11).Interior.ColorIndex = 6
            End If
            
            
            'reset values once you hit a new ticker
            
            open_price = 0
            
            Cells(ticker_count + 1, 12).Value = volume
            
            volume = 0
            
        End If
        
    Next i
    
Next ws
    
                

End Sub
