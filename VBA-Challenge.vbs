Sub StockAnalysis():

    'Declaring All Variables
    Dim ws As Worksheet
    
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    Dim TotalVolume As LongLong
    Dim SummaryRow As Long
    Dim SummaryRowNEW As Long
    Dim LastRow As Long
    Dim LastRowNEW As Long
    
    Dim TickerNEW As String
    Dim YearlyChangeNEW As Double
    Dim PercentChangeNEW As Double
    Dim TotalVolumeNEW As LongLong
    
    'Looping through all worksheets
    For Each ws In Worksheets
        'Naming Output Columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        'Initialize the values
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        SummaryRow = 2
        TotalVolume = 0
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVolume = 0
        
        'Declaring Values for first ticker
        Ticker = Cells(2, 1).Value
        OpeningPrice = Cells(2, 3).Value
        TotalVolume = Cells(2, 7).Value
        
        'Looping through each of the rows in worksheet
        For i = 3 To LastRow
            
            'Defining condition: If next ticker symbol is not equal to the current ticker symbol
            If (Ticker <> ws.Cells(i, 1).Value) Then
                'Calculating the Closing Price, Yearly Change and Percent Change
                ClosingPrice = ws.Cells(i - 1, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                PercentChange = YearlyChange / OpeningPrice
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                ' Format the percent change as a percentage and yearly change as currency as $
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                ws.Cells(SummaryRow, 10).NumberFormat = "$0.00"
                'Incrementing the summary table row
                SummaryRow = SummaryRow + 1
                
                'Define the values for new ticker symbol
                OpeningPrice = ws.Cells(i, 3).Value
                Ticker = ws.Cells(i, 1).Value
                TotalVolume = ws.Cells(i, 7).Value
            
                Else
                ' If the current row has the same ticker symbol as the previous row,
                ' update the total volume for the ticker symbol
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
            End If
        
        Next i
        
        'Define the values for last ticker symbol
        ClosingPrice = ws.Cells(LastRow, 6).Value
        PercentChange = YearlyChange / OpeningPrice
        ws.Cells(SummaryRow, 9).Value = Ticker
        ws.Cells(SummaryRow, 10).Value = YearlyChange
        ws.Cells(SummaryRow, 11).Value = PercentChange
        ws.Cells(SummaryRow, 12).Value = TotalVolume
        
        ' Format the percent change as a percentage and yearly change as currency as $
        ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
        ws.Cells(SummaryRow, 10).NumberFormat = "$0.00"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Declaring values for the new table range for columns I,J,K,L
        LastRowNEW = ws.Cells(Rows.Count, 9).End(xlUp).Row
        SummaryRowNEW = 2
        'Looping through the new table range for columns I,J,K,L
        For j = 2 To LastRowNEW
        
            TickerNEW = ws.Cells(j, 9).Value
            YearlyChangeNEW = ws.Cells(j, 10).Value
            PercentChangeNEW = ws.Cells(j, 11).Value
            TotalVolumeNEW = ws.Cells(j, 12).Value
            
            ' Set conditional formatting for yearly change cell
            If YearlyChangeNEW > 0 Then
                ws.Cells(SummaryRowNEW, 10).Interior.ColorIndex = 4
            ElseIf YearlyChangeNEW < 0 Then
                ws.Cells(SummaryRowNEW, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(SummaryRowNEW, 10).Interior.ColorIndex = 0
            End If
            
            'Compare values to find the maximum and minimum
            If PercentChangeNEW > ws.Cells(2, 17).Value Then
                ws.Cells(2, 16).Value = TickerNEW
                ws.Cells(2, 17).Value = PercentChangeNEW
                
            End If
            
            If PercentChangeNEW < ws.Cells(3, 17).Value Then
                ws.Cells(3, 16).Value = TickerNEW
                ws.Cells(3, 17).Value = PercentChangeNEW
               
            End If
            
            If TotalVolumeNEW > ws.Cells(4, 17).Value Then
                ws.Cells(4, 16).Value = TickerNEW
                ws.Cells(4, 17).Value = TotalVolumeNEW
               
            End If
            SummaryRowNEW = SummaryRowNEW + 1
        Next j
        
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Columns("A:Z").AutoFit
        
    Next ws

End Sub


