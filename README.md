# Stock Analysis
UC Berkeley BootCamp challenge 2 (VBA)
## Problem Statement
Steve has just graduated with his finance degree. His are going to be his first clients. They are passionate about green energy, they have decided to invest all their money into DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. But Steve thinks that his parents' funds should be more diversified, so he wants to analyze several green energy stocks, in addition to DAQO stock.

Steve has given us an Excel file containing the stock data he wants us to analyze. We'll be using an extension to Excel, built to automate tasks: Visual Basic for Applications, usually referred to as VBA.
 Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. 

## Overview of Project
Refactor the code successfully make the VBA script run faster.
```
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
```
## Results
<img src="https://github.com/ArmineKhanan/stock-analysis/blob/main/ASA%20Runtime%20for%202018.png" width="450" />         <img src="https://github.com/ArmineKhanan/stock-analysis/blob/main/ASA%20Runtime%20for%202018%20if%20refactored.png" width="450" />
## Summary
