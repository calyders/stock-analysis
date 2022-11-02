# Stocks Analysis VBA Challenge
## Overview of Project
### Purpose
The purpose of this project was to refactor the code written in Module 2. The purpose behind this edit was to create a faster code that would gather the same data in a shorter amount of time. 
## Results
### Analysis
Provided below are the changes I made to the code to create the new refactored code.

    '1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
According to the data gathered by the code, there is a significant difference in stock performance between the years 2017 and 2018. Namely, while 2017 had largely successful returns, as evident by the 11 out 12 cells being highlighted green, 2018 had largley negative returns with 10 out of 12 cells being highlighted red. Looking at this data now, as of 2018, ENPH and RUN seem to be the only stocks worth investing in.

![2017_Stocks](https://user-images.githubusercontent.com/115501756/199622030-2a8ee970-f070-4fa7-8a9c-f236cec020b5.png)
![2018_Stocks](https://user-images.githubusercontent.com/115501756/199622042-10e93f24-353d-43d3-bf51-448e4fa3effb.png)

As for the differences between the original code and the edited, refactored code, the refactored code performed much faster. For the 2017 and 2018 data, the original code came back with times of 0.69 seconds and 0.68 seconds, while the refactored code provided times of 0.10 seconds and 0.14 seconds.

![2017_Original](https://user-images.githubusercontent.com/115501756/199622209-2431a362-f152-403f-91b7-425e93725c63.png)
![2018_Original](https://user-images.githubusercontent.com/115501756/199622214-9dae7895-1b24-4d84-94e4-bc6c55204dce.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/115501756/199622217-ab2e9d17-97aa-4708-b0c1-16fb0ffddf45.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/115501756/199622222-0e538484-2819-4930-aa41-8bdf56c00901.png)

## Summary
### What are the advantages or disadvantages of refactoring code?
In terms of the code used for this project, the advantage that comes from refactoring code would be an improvement in runtime for the macro. As is seen in the analysis, the refactored code came back with times half a second shorter than the original code. A disadvantage that I thought of comes from the idea of refactoring someone elses code. Someone else might not write their script as well-organized as myself which could make trying to edit it rather difficult. It would be hard to refactor a code I didn't write myself or don't know how to read because of poor organization.
### How do these pros and cons apply to refactoring the original VBA script?
Some pros of refactoring code would be better organization, making the script easier to read for ourselves and for others who might use the code after us. 
