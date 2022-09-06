# Stock Analysis
UC Berkeley BootCamp challenge 2 (VBA)
## Overview of Project
Refactor the code successfully make the VBA script run faster.
```
'2. Initialize an Array of all tickers
    Dim tickers(12) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

'3. Prepare for an analysisof tickers.
    '1. Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingPrice As Single
    '2. Activate the data worksheet
    Worksheets(yearValue).Activate
    '3. Find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
'4. Loop through the tickers
    For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
             Worksheets(yearValue).Activate
            '5. Loop through rows in the data
            For j = 2 To RowCount
                '1. Find total volume for the current ticker.
                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If
                '2. Find starting price for the current ticker
                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                    startingPrice = Cells(j, 6).Value
                End If
                '3. Find ending price for the current ticker
                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
            Next j
'6. Output the datafor the current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(7 + i, 2).Value = ticker
        Cells(7 + i, 3).Value = totalVolume
        Cells(7 + i, 4).Value = (endingPrice / startingPrice) - 1
    Next 
```
##Results
Getting red of nested loop:
```
''2a) Create a for loop to initialize the tickerVolumes to zero.
    tickerIndex = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
## Summary
