Sub testscript()

    'Defining variables
    '------------------
    Dim ticker As String
    Dim openprice, closeprice As Double
    Dim row, resultsrow As Long
    Dim lastrow As Double
    Dim bonusLastRow As Double
    Dim volume As Double
    Dim ws As Worksheet
    Dim greatestVal As Double
    Dim greatestTick As String
    Dim leastVal As Double
    Dim leastTick As String
    Dim highVolume As Double
    Dim highTick As String
    Dim maxPercent As Double
    Dim minPercent As Double
    Dim maxVolume As Double
    
    'Starts a count for each worksheet and initializes it to 1
	'---------------------------------------------------------
    Dim sheetInc As Double
    sheetInc = 1

    'Sets baseline results table row
    '-------------------------------
    resultsrow = 2

    'Starts looping through each worksheet
    For Each ws In Worksheets
    
        Sheets(sheetInc).Activate

        'Clears table area for data to insert into
        ws.Range("I2:Q1048576").Clear
    
        'Creates column headers for the results to print under
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change ($)"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'Determine the last row of the data
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
        'Sets initial ticker value
        ticker = ws.Cells(2, 1).Value
    
        'Starts the loop process through data
        For row = 2 To lastrow
    
            'If ticker is the same, set open price value and start adding up volume for that ticker
            If ws.Cells(row, 1).Value = ws.Cells(row + 1, 1).Value Then
                If openprice = 0 Then
                    openprice = ws.Cells(row, 3).Value
                Else
                End If
            
                volume = volume + ws.Cells(row, 7).Value
            
            'If ticker is not the same, set close price. Also, if at end of data, still set close price to last data row value
            Else
                If ws.Cells(row, 1).Value <> ws.Cells(row - 1, 1).Value Then
                    closeprice = ws.Cells(row, 6).Value
                Else
                    closeprice = ws.Cells(row, 6).Value
                End If
            
                'Calculates, prints, and formats results table data
                ws.Cells(resultsrow, 9).Value = ws.Cells(row, 1).Value
                ws.Cells(resultsrow, 10).Value = closeprice - openprice
                    If ws.Cells(resultsrow, 10).Value > 0 Then
                        ws.Cells(resultsrow, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(resultsrow, 10).Value < 0 Then
                        ws.Cells(resultsrow, 10).Interior.ColorIndex = 3
                    End If
                ws.Cells(resultsrow, 11).Value = (closeprice / openprice) - 1
                    ws.Cells(resultsrow, 11).NumberFormat = "0.00%"
                volume = volume + ws.Cells(row, 7).Value
                ws.Cells(resultsrow, 12).Value = volume
                    ws.Cells(resultsrow, 12).NumberFormat = "0"
        
                'Adds to results table row count
                resultsrow = resultsrow + 1
        
                'Resets volume, open price, and close price values for next loop
                volume = 0
                openprice = 0
                closeprice = 0
        
            End If
        
        Next row
    
        'Uses AutoFit function to format results table columns
        ws.Columns("I:Q").AutoFit
    
        'Resets results table row value for use in next worksheet
        resultsrow = 2
        
        'Finds last row of first results table
        bonusLastRow = ws.Cells(Rows.Count, 11).End(xlUp).row
        
        
        'Finds the max/min percentage, and max volume values from initial results table
        maxPercent = WorksheetFunction.Max(Range("K2:K1501"))
        minPercent = WorksheetFunction.Min(Range("K2:K1501"))
        maxVolume = WorksheetFunction.Max(Range("L2:L1501"))
        
        'Start looping through first results table to gather desired info and compare it to the above values
        'Then, print and format those values into the second results table
        For row = 2 To bonusLastRow
            If ws.Cells(row, 11).Value = maxPercent Then
                greatestTick = ws.Cells(row, 9).Value
                greatestVal = ws.Cells(row, 11).Value
                ws.Cells(2, 16).Value = greatestTick
                ws.Cells(2, 17).Value = greatestVal
                ws.Range("Q2").NumberFormat = "0.00%"
            End If
        
            If ws.Cells(row, 11).Value = minPercent Then
                leastTick = ws.Cells(row, 9).Value
                leastVal = ws.Cells(row, 11).Value
                ws.Cells(3, 16).Value = leastTick
                ws.Cells(3, 17).Value = leastVal
                ws.Range("Q3").NumberFormat = "0.00%"
            End If
        
            If ws.Cells(row, 12).Value = maxVolume Then
                highTick = ws.Cells(row, 9).Value
                highVolume = ws.Cells(row, 12).Value
                ws.Cells(4, 16).Value = highTick
                ws.Cells(4, 17).Value = highVolume
                ws.Cells(4, 17).NumberFormat = "0.00E+00"
            End If
            
        Next row
        
        'Reset various values that were used above, to use again in the next worksheets
        greatestVal = 0
        greatestTick = " "
        leastVal = 0
        leastTick = " "
        highVolume = 0
        highTick = " "
        maxPercent = 0
        minPercent = 0
        maxVolume = 0
        
        'Takes sheet index variable and increments for next loop
        sheetInc = sheetInc + 1
        
    Next ws
                
End Sub