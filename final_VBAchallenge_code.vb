Sub TickerCounter():

    
    For Each ws In Worksheets
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Label Summary Table
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Set Values for Summary Table
    Dim ticker_symbol As String
    Dim price_change As Double
    Dim ticker_table As Double
    ticker_table = 2
    Dim ticker_start As Double
    ticker_start = 2
    Dim volume As Double
    volume = 0
    
    ' Label second summary table
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    ' Set Values for second summary table
    Dim percent_inc As Double
    percent_inc = 0
    Dim percent_dec As Double
    percent_dec = 0
    Dim greatest_volume As Double
    greatest_volume = 0
    Dim percent_inc_ticker As String
    Dim percent_dec_ticker As String
    Dim greatest_volume_ticker As String
        
        ' Loop through ticker symbols to find differences
        For i = 2 To lastrow
            
            ' Find tickers
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Store ticker symbol
                ticker_symbol = ws.Cells(i, 1).Value
                ws.Cells(ticker_table, 9).Value = ticker_symbol
                
                ' Find end of year value
                price_change = ws.Cells(i, 6) - ws.Cells(ticker_start, 3)

                ' Place price change in summary table
                ws.Cells(ticker_table, 10) = price_change
                
                ' Apply conditional formatting to Yearly Change
                Select Case price_change
                
                    Case Is > 0
                    
                    ' Format as green
                        ws.Cells(ticker_table, 10).Interior.ColorIndex = 4
                        
                    Case Is < 0
                    
                    ' Format as red
                        ws.Cells(ticker_table, 10).Interior.ColorIndex = 3
                    
                    Case Else
                
                End Select
                
                ' avoid dividing by zeros
                If ws.Cells(ticker_start, 3) <> 0 Then
                
                    ' Calculate percent change
                    ws.Cells(ticker_table, 11) = price_change / ws.Cells(ticker_start, 3)
                    
                    ' Format as percentage
                    ws.Cells(ticker_table, 11).NumberFormat = "0.00%"
                
                Else
                    ws.Cells(ticker_table, 11) = ws.Cells(ticker_start)
                    
                End If
                
                ' Place ticker volume in summary table
                ws.Cells(ticker_table, 12) = volume
                
                ' =============================================================================
                
                ' Make new row in summary table
                ticker_table = ticker_table + 1
                
                ' change ticker start to new ticker
                ticker_start = i + 1
            
            Else
                
                ' Add ticker volume
                volume = volume + ws.Cells(i, 7)
                
            End If
            
            
        ' ====================================================================================
        Next i
        
        ' Iterate through first summary table
        For i = 2 To lastrow
        
            ' =================================================================================
            ' Search for maximum percent increase
            If ws.Cells(i, 11).Value > percent_inc Then
                
                ' change greatest percent increase to the higher value
                percent_inc = ws.Cells(i, 11).Value
                
                ' place greatest percent increase ticker in second summary table
                percent_inc_ticker = ws.Cells(i, 9).Value
                ws.Range("P2") = percent_inc_ticker
                
                ' place greatest percent increase value in second summary table
                ws.Range("Q2") = percent_inc
                
                ' format as a percentage
                Range("Q2").NumberFormat = "0.00%"
                
            End If
            
            ' =================================================================================
            ' Search for greatest percent decrease
            If ws.Cells(i, 11).Value < percent_dec Then
                
                ' change greatest percent decrease to the lower value
                percent_dec = ws.Cells(i, 11).Value
                
                ' place greatest percent decrease ticker in second summary table
                percent_dec_ticker = ws.Cells(i, 9).Value
                ws.Range("P3") = percent_dec_ticker
                
                ' place greatest percent decrease value in second summary table
                ws.Range("Q3") = percent_dec
                
                ' format as a percentage
                Range("Q3").NumberFormat = "0.00%"
                
            End If
            
            ' =================================================================================
            ' Search for highest total volume
            If ws.Cells(i, 12).Value > greatest_volume Then
                
                ' change greatest total volume to the higher value
                greatest_volume = ws.Cells(i, 12).Value
                
                ' place greatest total volume ticker in second summary table
                greatest_volume_ticker = ws.Cells(i, 9).Value
                ws.Range("P4") = greatest_volume_ticker
                
                ' place greatest total volume value in second summary table
                ws.Range("Q4") = greatest_volume
                
            End If
        
        Next i
        
        ' Autofit columns
        ws.Cells.EntireColumn.AutoFit
        
        
    Next ws
    
End Sub
