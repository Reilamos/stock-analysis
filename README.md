# Green Stocks Analysis

# Project Overview

Steve loved the workbook that was prepared for him to analyze an entire data set. Now he is asking to expand the dataset to include the entire stock market. Using (VBA) Visual Basic Application and Excel to provide return on investment and the annual volume depending on what year is input.

Another thing monitored will be how long it takes to run the code. If Steve wants to analyse a higher amount of stocks I want to make it as fast as possible.

# Results

The first thing to do was set up new arrays tickerStartingPrices as Single and tickerEndingPrices as Single and tickerVolumes = 0.
    
    For i = 0 To 11
       ticker = tickers(i)
        tickerVolumes = 0
    Dim tickerSartingPrices As Single
    Dim tickerEndingPrices As Single

Using a for loop with If Then to store the data for each stock and then display the results of the perfomaing stocks volume and return:

  For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
         If Cells(j, 1).Value = ticker Then
          
            tickerVolumes = tickerVolumes + Cells(j, 8).Value

           End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               tickerStartingPrices = Cells(j, 6).Value
               
         End If
              
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowÕs ticker doesnÕt match, increase the tickerIndex.
        'If  Then
            
             If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               tickerEndingPrices = Cells(j, 6).Value

           End If
            

            '3d Increase the tickerIndex.
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               tickerStartingPrices = Cells(j, 6).Value
               
         End If
         
          If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               tickerEndingPrices = Cells(j, 6).Value

           End If
               
        'End If

    Next j
    
See posted screenshots of results of All Stocks 2017 and 2018 data. To monitor how long it takes for the code to run I used:

    Dim startTime As Single
    Dim endTime  As Single
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

# Summary

There are advantages and disadvantages  to refactoring code. On one hand it does give you some insight into what the person who wrote the code is working on. But on the other hand the code doesn't necesarily fit with your existing data. You also have to comb through it and make sure there are no issues with existing code.

Refactoring VBA script is workable. There are resources online if you run into issues. I did run into a few errors that took awhile to settle (code 6 and code 9 I believe).


