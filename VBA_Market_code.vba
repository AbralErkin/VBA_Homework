Sub VB_market()
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    'Clear all the existing formats for all the empty cells
    Range("I:Z").ClearContents
    Range("I:Z").ClearFormats
    
    'Assign headers to the corresponding columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly_Change"
    Cells(1, 11).Value = "Percen_Change"
    Cells(1, 12).Value = "Total_Stock_Volume"
    Cells(1, 15).Value = "Criteria"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'Create variables
    Dim Total As Double
    Dim Num As Double
    Dim ch_yr(8000, 8000) As Double
    Dim chng As Double
    Dim max_inc As Double
    Dim max_dec As Double
    Dim max_vol As Double
    Dim ticker1 As String
    Dim ticker2 As String
    Dim ticker3 As String
    Dim lsrw As Double
    'Assign values to the variables
    Total = 0
    Num = 0
    max_inc = -1
    max_dec = 1
    max_vol = 0
    ticker1 = zero
    ticker2 = zero
    ticker3 = zero
    lsrw = Cells(Rows.Count, 1).End(xlUp).Row
    
    'search through the data set
    For i = 2 To lsrw + 1
        
        'calculate the total volume
        Total = Total + Cells(i, 7).Value
        
        'Fing the unique tickers
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            'number of unique tickers
            Num = Num + 1
            'assign the opening value of the begining of the year to each ticker
            ch_yr(Num, 1) = Cells(i, 3).Value
            
            If i > 2 Then
                'assign the total value of the previous ticker (after search through all the numbers for that ticker)
                Cells(Num, 12).Value = Total
                Total = 0
                
                'calculate the yearly change and percent_change for non-zero values
                ch_yr(Num - 1, 2) = Cells(i - 1, 6).Value
                
                If ch_yr(Num - 1, 2) * ch_yr(Num - 1, 1) <> 0 Then
                    chng = ch_yr(Num - 1, 2) - ch_yr(Num - 1, 1)
                    Cells(Num, 10).Value = chng
                    'format the cells
                    If chng < 0 Then
                        Cells(Num, 10).Interior.ColorIndex = 3
                    Else
                        Cells(Num, 10).Interior.ColorIndex = 4
                    End If
                
                    'calculate the yearly percent_change
                    Cells(Num, 11).Value = chng / ch_yr(Num - 1, 1)
                    Cells(Num, 11).NumberFormat = "0.00%"
                Else
                    Cells(Num, 10).Value = 0
                    Cells(Num, 11).Value = 0
                End If
                
                'Find the greatest % increase
                If Cells(Num, 11).Value < max_inc Then
                    Cells(2, 17).Value = max_inc
                    Cells(2, 16).Value = ticker1
                Else
                    max_inc = Cells(Num, 11).Value
                    ticker1 = Cells(Num, 9).Value
                    Cells(2, 17).Value = max_inc
                    Cells(2, 16).Value = ticekr1
                End If
                Cells(2, 17).NumberFormat = "0.00%"
                
                'Find the greatest % decrease
                If Cells(Num, 11).Value > max_dec Then
                    Cells(3, 17).Value = max_dec
                    Cells(3, 16).Value = ticker2
                Else
                    max_dec = Cells(Num, 11).Value
                    ticker2 = Cells(Num, 9).Value
                    Cells(3, 17).Value = max_dec
                    Cells(3, 16).Value = ticker2
                End If
                Cells(3, 17).NumberFormat = "0.00%"
                
                'Find the Greatest Total volume
                If Cells(Num, 12).Value < max_vol Then
                    Cells(4, 17).Value = max_vol
                    Cells(4, 16).Value = ticker3
                Else
                    max_vol = Cells(Num, 12).Value
                    ticker3 = Cells(Num, 9).Value
                    Cells(4, 17).Value = max_dec
                    Cells(4, 16).Value = ticker3
                End If
                
            End If
            
            'assign the ticker value to the cells (pick the first value for each ticker)
            Cells(Num + 1, 9).Value = Cells(i, 1).Value
        End If
        
    Next i
    
    Next ws
End Sub