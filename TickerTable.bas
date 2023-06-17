Attribute VB_Name = "Module1"
Sub TickerTable():
    
    Dim openingPrice, closingPrice, volumeTotal As Double
    
    Dim currentStock As String
    
    Dim ws As Worksheet
    
    'Loop which will analyze stocks for each sheet'
    For Each ws In Worksheets
    
        Dim resultRowIndex, lastRow As Double
        
        volumeTotal = 0
        resultRowIndex = 2
        
        'Count number of rows in worksheet'
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Variables for tracking greatest % increase, decrease, total volume'
        Dim greatestPercentIncrease, greatestPercentDecrease, greatestVolume As Double
        Dim greatestPercentIncreaseName, greatestPercentDecreaseName, greatestVolumeName As String
        greatestPercentIncrease = 0
        greatestPercentDecrease = 0
        greatestVolume = 0
        greatestPercentIncreaseName = ""
        greatestPercentDecreaseName = ""
        greatestVolumeName = ""
        
        
        'Set up stock results table'
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Set up Greatest % Increase/Decrease table'
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Set openingPrice to first <open> price in the table'
        openingPrice = ws.Cells(2, 3).Value
        
        'Set currentStock to first ticker name in the table'
        currentStock = ws.Cells(2, 1).Value
        
        'Loop through all rows in stocks table'
        For I = 2 To (lastRow + 1)
            
            'Code for when loop reaches end of a given stock'
            If (ws.Cells(I, 1).Value <> currentStock) Then
            
                'Record stock values at row given by resultRowIndex'
                ws.Cells(resultRowIndex, 9) = currentStock
                ws.Cells(resultRowIndex, 10) = closingPrice - openingPrice
                ws.Cells(resultRowIndex, 11) = (closingPrice - openingPrice) / openingPrice
                ws.Cells(resultRowIndex, 12) = volumeTotal
                
                'Set color of "Yearly Change" based on sign of value'
                If (ws.Cells(resultRowIndex, 10).Value >= 0) Then
                
                    ws.Cells(resultRowIndex, 10).Interior.ColorIndex = 4
                    
                Else
                    
                    ws.Cells(resultRowIndex, 10).Interior.ColorIndex = 3
                    
                End If
                
                'Format "Percent Change" column as percent'
                ws.Cells(resultRowIndex, 11).NumberFormat = "0.00%"
                
                'Check for new greatest % inc, % dec, total volume'
                If (ws.Cells(resultRowIndex, 11).Value > greatestPercentIncrease) Then
                
                    greatestPercentIncrease = ws.Cells(resultRowIndex, 11)
                    greatestPercentIncreaseName = currentStock
                    
                End If
                
                If (ws.Cells(resultRowIndex, 11).Value < greatestPercentDecrease) Then
                    
                    greatestPercentDecrease = ws.Cells(resultRowIndex, 11)
                    greatestPercentDecreaseName = currentStock
                    
                End If
                
                If (volumeTotal > greatestVolume) Then
                
                    greatestVolume = volumeTotal
                    greatestVolumeName = currentStock
                
                End If
                
                'Reinitialize tracking variables'
                currentStock = ws.Cells(I, 1).Value
                openingPrice = ws.Cells(I, 3).Value
                
                'Set volumeTotal equal to the stock's day one volume'
                volumeTotal = ws.Cells(I, 7).Value
                
                'Increment the resultRowIndex'
                resultRowIndex = resultRowIndex + 1
                
            
            'Code for when loop is still processing a given stock'
            ElseIf (ws.Cells(I, 1).Value = currentStock) Then
            
                'Add the day's stock volume to total'
                volumeTotal = volumeTotal + ws.Cells(I, 7)
            
            End If
            
            'If statement which sets closingPrice on last occurrence of a stock'
            If (ws.Cells(I + 1, 1) <> currentStock) Then
            
                closingPrice = ws.Cells(I, 6)
                
            End If
            
        Next I
        
        'Populate greatest % increase, % decrease, total volume table'
        ws.Range("P2").Value = greatestPercentIncreaseName
        ws.Range("Q2").Value = greatestPercentIncrease
        ws.Range("P3").Value = greatestPercentDecreaseName
        ws.Range("Q3").Value = greatestPercentDecrease
        ws.Range("P4").Value = greatestVolumeName
        ws.Range("Q4").Value = greatestVolume
        
        'Format cells Q2 and Q3 as percents'
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Resize columns I through Q to fit text'
        ws.Columns("I:Q").AutoFit
    Next
    
End Sub

