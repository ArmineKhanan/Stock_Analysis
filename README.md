# Stock Analysis
UC Berkeley BootCamp challenge 2 (VBA)
## Overview of Project
Steve's parents are passionate about green energy and are eager to invest all their money into DAQO New Energy Corporation. Yet, they decided to seek advice from Steve, who graduated with his finance degree not long ago. The letter, in his turn, applied to us for assistance in analysis. We solved the problem using an extension to Excel, built to automate tasks: Visual Basic for Applications, usually referred to as VBA.

Steve is happy. But to do a little more research for his parents, Steve wants to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen stocks, it might take a long time to execute for thousands of them.

In this challenge, we will edit and refactor our initial solution code. Afterward, we plan to determine whether refactoring our code made the VBA script run faster.

## Results
#### Script Editing
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
        For i = 2 To RowCount
            ***
            ***
            'Increase the tickerIndex.
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                tickerIndex = tickerIndex + 1
            End If
        Next i
```
Loop through previously defined arrays to output the ticker name, daily volume, and return.
```
For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(7 + i, 2).Value = tickers(i)
        Cells(7 + i, 3).Value = tickerVolumes(i)
        Cells(7 + i, 4).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
```
#### VBA code runtime recorded
In order to assess the affect of changes on the speed of report production we decorated the whole presedure with the folowng script:
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
        ***
        ***
        ***
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```
The screenshots below prove that the refactored script is almost 2.5 faster in average.

<img src="https://github.com/ArmineKhanan/stock-analysis/blob/main/ASA%20Runtime%20for%202017.png" width="450" />                                    <img src="https://github.com/ArmineKhanan/stock-analysis/blob/main/ASA%20Runtime%20for%202017%20if%20refuctored.png" width="450" />

<img src="https://github.com/ArmineKhanan/stock-analysis/blob/main/ASA%20Runtime%20for%202018.png" width="450" />                                    <img src="https://github.com/ArmineKhanan/stock-analysis/blob/main/ASA%20Runtime%20for%202018%20if%20refactored.png" width="450" />

## Summary
Code refactoring may be time-consuming and even a risky endeavor. For example, if one does not have an entire understanding of the code, they better refrain from refactoring. Though for code maintenance, its readability and effectiveness refactoring can be fruitful. 
In our case, code refactoring resulted in speedy performance. Being aware of Steve's plans to broaden the analysis, we believe our efforts will be even more rewarding in the future. 
Yet, the refactored script has a limitation as compared to the initial. It will work effectively if only the ticker names have the same sequence in the data source as in the script.

