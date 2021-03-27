# VBA Stock Analysis
This stock analysis using Excel's VBA function provides a performance snapshot for 2017 and 2018 across 12 tickers to make an informative decision on a hold ivestment strategy.

1. First, a code was created to run specifically on DQ ticker performance
2. Next the code was reused to focus on All Stocks in the index
3. Then, a code initialized the index to loop through all 12 tickers
  - Within this loop, ticker starting and ending prices were assigned based off the Ticker Index
4. Finally, the analysis runs evaluating performance across the stock array for 2017 and 2018


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
