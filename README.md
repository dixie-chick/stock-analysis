# VBA Stock Analysis
This stock analysis using Excel's VBA function provides a performance snapshot for 2017 and 2018 across 12 tickers to make an informative decision on a hold ivestment strategy.

1. First, a code was created to run specifically on DQ ticker performance
```
Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
```

2. Next the code was reused to focus on All Stocks in the index

```
 '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
3. Then, a code initialized the index to loop through all 12 tickers
     - Within this loop, ticker starting and ending prices were assigned based off the Ticker Index
     ```'2a) Create a for loop to initialize the tickerVolumes to zero.
    
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
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
                
            End If
                     
            
        'End If
        ```
        
4. Finally, the analysis runs evaluating performance across the stock array for 2017 and 2018
```
  Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
```

## Can a Visual Basic Macro be refactored for increased efficiency?

### Lets take a look at the time to run from the original code created:

![2018_Original](https://user-images.githubusercontent.com/79612565/112698808-bbd43180-8e47-11eb-9dff-fe8c09121996.png)


### Now lets compare this to the run time after refactoring:

![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/79612565/112692275-e8824c00-8e3b-11eb-9a46-b841ac7136f7.png)

![VBA_Challenge_2017_Refactored](https://user-images.githubusercontent.com/79612565/112704794-26da3400-8e59-11eb-8474-cd09f4a36277.png)



![Great_Sucess](https://user-images.githubusercontent.com/79612565/112704720-e11d6b80-8e58-11eb-9b16-00cc3636664e.jpg)

## YES! Refactored VBA has successfully increased efficiency

## Pros & Cons of Refactoring
# What are the advantages or disadvantages of refactoring code?
Advantages include: 
1. Increased speed when running
2. Clean, easy to read functions
3. Debugged and easier to maintain

Disadvantages include:
1. Time consuming: Refactoring and Debugging takes a full process!
2. Complexity: Risk of incorrectly assigning new code
3. Chance of mistakes: accidentally introduce new bugs


## To Wrap it Up

Refactoring ultimately led to a cleaner, organized code improving software, design, debugging and faster programming. Sharing this code with others is more understandable benefitting end users who might want to leverage in their own analyses.    
