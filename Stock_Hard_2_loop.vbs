Sub WorkSheet_Loop()

For Each ws In Worksheets
        
        
        'Set the StockTicker as a string
        Dim Ticker, Great_Percent_Ticker, Least_Percent_Ticker, Great_Volume_Ticker As String
        
        
        'Set variables to capture Opening/Closing and volume totals plus outputs
        Dim Opening_Total, Closing_Total, Total_Volume, r, Great_Percent, Least_Percent, Great_Volume As Double
        Opening_Total = 0
        Closing_Total = 0
        Total_Volume = 0
        Great_Percent = 0
        Least_Percent = 0
        Great_Volume = 0
        
        'Calculate the Last Row in the first column of the sheet
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set variable to manage outputs while in loop
        r = 2
        Least_Date = ws.Cells(2, 2).Value
        Max_Date = 0
        
        ' Create Header Rows
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Loop to Calculate the Opening Total, Closing Total, Volume, and product outputs
        
        For i = 2 To lrow
            Ticker = ws.Cells(i, 1).Value
            'Ensure Date is the lease date for opening total
            If ws.Cells(i, 2).Value <= Least_Date Then
                Opening_Total = ws.Cells(i, 3).Value
                Least_Date = ws.Cells(i, 2).Value
            End If
                
            ' If Closing Closing_Total
            If ws.Cells(i, 2).Value >= Max_Date Then
                Closing_Total = ws.Cells(i, 6).Value
                Max_Date = ws.Cells(i, 2).Value
            End If
            
            'Add Total_Volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            'Produce Outputs when card changes and reset values
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(r, 9).Value = Ticker
                ws.Cells(r, 10).Value = Closing_Total - Opening_Total
                If Opening_Total = 0 Then
                    ws.Cells(r, 11).Value = "Opened at Zero"
                Else
                    ws.Cells(r, 11).Value = Format(((Closing_Total - Opening_Total) / Opening_Total), "Percent")
                End If
                If (Closing_Total - Opening_Total) < 0 Then
                    ws.Cells(r, 10).Interior.ColorIndex = 3
                ElseIf (Closing_Total - Opening_Total) < 0 Then
                    ws.Cells(r, 10).Interior.ColorIndex = 6
                Else
                    ws.Cells(r, 10).Interior.ColorIndex = 4
                End If
                
                ws.Cells(r, 12).Value = Total_Volume
                
                'Evaluate to determine greatest increase, decrease, and trading volume
                If Opening_Total <> 0 Then 'Ensure if opening value was zero, it is not evaluated
                    If Great_Percent < ((Closing_Total - Opening_Total) / Opening_Total) Then
                        Great_Percent = ((Closing_Total - Opening_Total) / Opening_Total)
                        Great_Percent_Ticker = Ticker
                    End If
                    
                    If Least_Percent > ((Closing_Total - Opening_Total) / Opening_Total) Then
                        Least_Percent = ((Closing_Total - Opening_Total) / Opening_Total)
                        Least_Percent_Ticker = Ticker
                    End If
                End If
                        
                If Great_Volume < Total_Volume Then
                    Great_Volume = Total_Volume
                    Great_Volume_Ticker = Ticker
                End If
                        
                'Reset Variables for Next Ticker Loop
                Opening_Total = 0
                Closing_Total = 0
                Total_Volume = 0
                Least_Date = ws.Cells(i + 1, 2).Value
                Max_Date = 0
                r = r + 1
            End If
        Next i
        
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = Great_Percent_Ticker
        ws.Cells(2, 17).Value = Format(Great_Percent, "Percent")
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = Least_Percent_Ticker
        ws.Cells(3, 17).Value = Format(Least_Percent, "Percent")
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = Great_Volume_Ticker
        ws.Cells(4, 17).Value = Great_Volume
        
        
            ' Autofit to display data
        ws.Columns("A:Q").AutoFit

Next ws


End Sub
