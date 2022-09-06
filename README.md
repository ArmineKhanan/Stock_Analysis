# Stock Analysis
UC Berkeley BootCamp challenge 2 (VBA)
## Problem Statement
Steve's parents are passionate about green energy, and are eager to invest all their money into DAQO New Energy Corporation. Yet, they decided to first seek advice from Steve, who's just graduated with his finance degree. The letter, inhis turn applied to us for assistance in analysis. We solved the problem using an extension to Excel, built to automate tasks: Visual Basic for Applications, usually referred to as VBA.

Steve is happy. Though, to do a little more research for his parents, Steve wants to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen stocks, it might take a long time to execute for thousands of stocks.

In this challenge, we’ll edit, or refactor, our initial solution code. Afterwards, we’ll determine whether refactoring our code successfully made the VBA script run faster. 

## Overview of Project
Our purpose is to refactor the code successfully and make the VBA script run faster. In search of an effective scenario where we can go through all the data one time we came to the conclusion to get rid of the nested loop. 
```
'Loop through the tickers
    For i = 0 To 11
            'Loop through rows in the data
            For j = 2 To RowCount
                ***
                ***
                End If
            Next j

    Next i
```

Getting rid of nested loop: initialising tickerIndex variable and increasing it's value by one whenever ticker name changes.
```
'Initialize tickerIndex
Dim tickerIndex As Integer
tickerIndex = 0
        
    'Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    
        ***
        ***

            'Increase the tickerIndex.
            If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
```
## Results
<img src="https://github.com/ArmineKhanan/stock-analysis/blob/main/ASA%20Runtime%20for%202018.png" width="450" />         <img src="https://github.com/ArmineKhanan/stock-analysis/blob/main/ASA%20Runtime%20for%202018%20if%20refactored.png" width="450" />
## Summary
<span style="color: red;">text</span>
