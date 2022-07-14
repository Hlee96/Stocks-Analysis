# Green Stock Analysis
## Proejct Overview
The overall project was to provide Steve with an analysis of number of Green Stocks to help the decision making for his parents, on what stock is it worth to invest in. The analysis work has been done through the usage of VBA (Visual Basic Application) in Excel to find each stock's total daily volumne and annual return for the year 2017 and 2018. By analyzing 12 different green stocks, the results gave the right stock for Steve's parents to invest in.
## Purpose
Within the dataset of stocks analysis,it included dataset of over 6000 rows for the year 2017 and 2018. Displaying each ticker, the specific date's open, high, low, close, and volumne, there was a need of an efficient way to anlayze the stocks and find the right stock for Steve's parents to invest. By utilizing the VBA tool it took less than a second to analyze which was the best stock to invest in. By refactoring my code the analysis has been done in a more efficient way.
#Results
## Analysis
Prior refactoring the code, first, I downloaded the VBS file that was needed to create the appropriate input box, headers, ticker arrays, and the appropriate worksheet to activate. The file has provided the instruction into refactoring the code.
## Refactored Code
   '1a) Create a ticker Index
   
        tickerindex = 0
    
    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
            Dim tickerStartingPrices(12) As Single
                Dim tickerEndingprices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingprices(i) = 0
    Next i
    
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
    tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
    
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
     If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
        tickerStartingPrices(tickerindex) = Cells(i, 6).Value
        
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row ticker doesn't match, increase the tickerIndex.
        'If  Then
      If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
        tickerEndingprices(tickerindex) = Cells(i, 6).Value
        
        End If
            

            '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerindex = tickerindex + 1
        
        End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumnes(i)
        Cells(4 + i, 3).Value = tickerEndingprices(i) / tickerStartingPrices(i) - 1
             
    Next i
By the results from running the macro, the results showed that the highest return green stock for the year 2018 was "RUN" (ticker), with a return of 84% and a total daily volume of $502,757,100. 
 ![VBA_Challenge_2018](https://user-images.githubusercontent.com/108282027/178908901-2cff399a-09eb-4dc5-b95e-1a54b02895b7.png)
 ![VBA_Challenge_2017](https://user-images.githubusercontent.com/108282027/178908953-fc2e9491-29b8-4e44-94c8-c8bc72feaa56.png)
Attached above is the result for the 2017 and 2018 stock analysis with the time that it took to generate the results.

##Summary
### Advantage vs Disadvantages of refactoring a code
The main advtange is the time efficiency that the refactored code is providing. The macro analysis was run about approximately .5 seconds less. When dealing with a large data set, time is crucial and more so being efficient with the time given to run the analysis is important. It also helped to eliminate redundancies in comparison to the original code. It is essentially aiming to achieve a better code and more effective one. Some of the disadvantage can relate to user error, when the refactored code introduces additional bugs that wasn't present prior. A refactored code might also be difficult for the user to understand, while we know the original code clearly, the process of refactoring a code might lead to more confusion.
### Orginal Code Time
![Screenshot 2022-07-14 093929](https://user-images.githubusercontent.com/108282027/179009083-db0f67cf-569c-41f7-8a3a-9c1ef1fb390a.png)
![Screenshot 2022-07-14 093953](https://user-images.githubusercontent.com/108282027/179009151-286d1ac5-3479-47ae-8c39-075df3259224.png)
### Advatnage seen in Stock Analysis
The decreased time to run the macro was the biggest advantage. When the original code was run it was closer to 1 second, and by refactoring it decreased the time for more than 50%. Also, there was no too much hands on work required to change the code from the original code, and it was still able to be more effective and time efficient.
### Refactored Code Time
![VBA_Challenge_2017](https://user-images.githubusercontent.com/108282027/179009276-e3a37681-6fb1-42ed-aa93-867635be7815.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/108282027/179009283-788e1257-2e45-4857-b3d3-a1dc2ed1da76.png)



