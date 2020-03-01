Sub StockInfo()
    
    Dim TotalVol, TickNum, WSCount As Integer
    Dim HighVol As Double
    Dim Ticker, HighPercTick, LowPercTick, HighVolTick As String
    Dim OpenPrice, ClosePrice, HighPerc, LowPerc As Double
    Dim PercentChange As Double
    Dim lastrow, I As LongPtr
    Dim ws As Worksheet
    


    
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Select

        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker" 'Setting up the format of the data
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        Range("K:K").NumberFormat = "0.00%"
        Range("P2", "P3").NumberFormat = "0.00%"
        
        TickNum = 1 'Setting variables so they can be compared later
        HighVol = 0
        HighPerc = 0
        LowPerc = 0

    
        For I = 2 To (lastrow + 1)
            If ws.Cells(I, 1).Value <> ws.Cells((I - 1), 1) Then
            
                If TickNum <> 1 Then 'This If statement only runs when the Ticker symbol changes, not on the first one
                    ClosePrice = ws.Cells((I - 1), 6).Value 'Grabs the final closing price of the last stock
                    
                    If OpenPrice <> 0 Then
                        PercentChange = ((ClosePrice - OpenPrice) / OpenPrice) 'Gets the percentage increase/decrease of the last stock price
                    Else
                        PercentChange = 0
                    End If
                    
                    If PercentChange > HighPerc Then 'If this PercentChange is greater than the current highest, updates HighPerc and notes the Ticker
                        HighPerc = PercentChange
                        HighPercTick = ws.Cells((I - 1), 1).Value
                    ElseIf PercentChange < LowPerc Then 'If this PercentChange is lower than the current lowest, updates LowPerc and notes the Ticker
                        LowPerc = PercentChange
                        LowPercTick = ws.Cells((I - 1), 1).Value
                    End If
                    
                    If TotalVol > HighVol Then 'If the TotalVol is greater than the current highest, updates HighVol and notes the Ticker
                        HighVol = TotalVol
                        HighVolTick = ws.Cells((I - 1), 1).Value
                    End If
                    
                    ws.Cells(TickNum, 9).Value = Ticker
                    ws.Cells(TickNum, 10).Value = (ClosePrice - OpenPrice)
                    ws.Cells(TickNum, 11).Value = PercentChange
                    ws.Cells(TickNum, 12).Value = TotalVol
                    
                    If (ClosePrice - OpenPrice) > 0 Then
                        ws.Cells(TickNum, 10).Interior.ColorIndex = 4
                    ElseIf (ClosePrice - OpenPrice) = 0 Then
                        ws.Cells(TickNum, 10).Interior.ColorIndex = 6
                    Else
                        ws.Cells(TickNum, 10).Interior.ColorIndex = 3
                    End If
                End If
                
                Ticker = ws.Cells(I, 1).Value
                OpenPrice = ws.Cells(I, 3).Value
                TotalVol = ws.Cells(I, 7).Value
                TickNum = TickNum + 1
            Else
                TotalVol = TotalVol + ws.Cells(I, 7).Value
            End If
            
        Next I
    
        ws.Cells(2, 15).Value = HighPercTick
        ws.Cells(2, 16).Value = HighPerc
        ws.Cells(3, 15).Value = LowPercTick
        ws.Cells(3, 16).Value = LowPerc
        ws.Cells(4, 15).Value = HighVolTick
        ws.Cells(4, 16).Value = HighVol
        TickNum = 1
    Next ws
End Sub

