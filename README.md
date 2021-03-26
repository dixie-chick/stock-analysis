# stock-analysis
Performing stock analysis on DQ if a sound investment
Overview of Project: Explain the purpose of this analysis.

## Can a Visual Basic Macro be refactored for increased efficiency?
    - This stock analysis through VBA provides a snapshot of performance across 12 stocks during 2017 and 2018 to make an informative decision  of a sound long term investment.
    - First a code was created to run a value analysis on DQ stock
    - Next the code was reused, now focusing on All Stocks
    - Then a code initialized an index to loop through 12 tickers
    -   Within this loop, ticker starting and ending prices were assigned based off of the Ticker Index
    - Finally, the anaylsis runs evaluating performance across the stock array for 2017 & 2018
   
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
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
                
            End If

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        

  Before Refactor
        ![2018_Original](https://user-images.githubusercontent.com/79612565/112698808-bbd43180-8e47-11eb-9dff-fe8c09121996.png)


## Results of Refactor with Efficient Performance
![VBA_Challenge_2017_Refactored](https://user-images.githubusercontent.com/79612565/112692273-e5875b80-8e3b-11eb-9cae-95b29daf6cdc.png)
![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/79612565/112692275-e8824c00-8e3b-11eb-9a46-b841ac7136f7.png)


## Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
