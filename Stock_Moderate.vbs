Sub Stock_Checker()

'Set the StockTicker as a string
Dim Ticker As String

'Set variables to capture Opening/Closing and volume totals plus outputs
Dim Opening_Total, Closing_Total, Total_Volume, r As Double
Opening_Total = 0
Closing_Total = 0
Total_Volume = 0

'Calculate the Last Row in the first column of the sheet
lrow = Cells(Rows.Count, 1).End(xlUp).Row

'Set variable to manage outputs while in loop
r = 2
Least_Date = Cells(2, 2).Value
Max_Date = 0

' Create Header Rows
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Loop to Calculate the Opening Total, Closing Total, Volume, and product outputs

For i = 2 To lrow
    Ticker = Cells(i, 1).Value
    'Ensure Date is the lease date for opening total
    If Cells(i, 2).Value <= Least_Date Then
        Opening_Total = Cells(i, 3).Value
        Least_Date = Cells(i, 2).Value
    End If
        
    ' If Closing Closing_Total
    If Cells(i, 2).Value >= Max_Date Then
        Closing_Total = Cells(i, 6).Value
        Max_Date = Cells(i, 2).Value
    End If
    
    
    'Add Total_Volume
    Total_Volume = Total_Volume + Cells(i, 7).Value
    
    'Produce Output when stock ticker changes and reset values
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Cells(r, 9).Value = Ticker
        Cells(r, 10).Value = Closing_Total - Opening_Total
        If Opening_Total = 0 Then 'Prevent Errors on annual change if there was no opening value or opening value was zero
            Cells(r, 11).Value = "Opened at Zero"
        Else
            Cells(r, 11).Value = Format(((Closing_Total - Opening_Total) / Opening_Total), "Percent")
        End If
        If (Closing_Total - Opening_Total) < 0 Then 'Format Green for positive, yellow no change, and red if negative
            Cells(r, 10).Interior.ColorIndex = 3
        ElseIf (Closing_Total - Opening_Total) < 0 Then
            Cells(r, 10).Interior.ColorIndex = 6
        Else
            Cells(r, 10).Interior.ColorIndex = 4
        End If
        
        Cells(r, 12).Value = Total_Volume 
        'Reset Variables for Next Ticker Loop
        Opening_Total = 0
        Closing_Total = 0
        Total_Volume = 0
        Least_Date = Cells(i + 1, 2).Value
        Max_Date = 0
        r = r + 1
    End If
Next i


End Sub

