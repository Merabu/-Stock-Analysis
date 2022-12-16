# Stock-Analysis
## Overview of Project: 
I take pride in trying to simplify the lives of others. The stock analysis project is aimed at helping Steve’s parents to research and expand the dataset to include the entire stock market over the last week

### Explain the purpose of this analysis.

The purpose of the stock analysis is to reuse the already built worksheet and customize it to the needs of Steve’s parents. I am asked to edit the solution for the previous solution to determine whether refactoring my code made the VBA script run faster. I am expected to make the codes more efficient, using less memory, or improving the logic of the code to make it easier for future users to read

## Analysis


Step 1a:

Create a tickerIndex variable and set it equal to zero before iterating over all the rows.
  Dim tickerIndex As Single

               tickerIndex = 0
               

 (this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step) 


Step 1b:

Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
The tickerVolumes array should be a Long data type.
The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

Step 2a:

Create a for loop to initialize the tickerVolumes to zero.

 For i = 0 To 11
    
     tickerVolumes(i) = 0

Step 2b: Create a for loop that will loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
Step 3a:  increases the current tickerVolumes (stock ticker volume) 


                   tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value




Step 3b:

Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable

If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If


Step 3c:

Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.

If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            


Step 3d:

If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
Step 4:

Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.

## Results compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

![All stock 2017](https://user-images.githubusercontent.com/115379848/207991801-3f92ea2f-75d1-4686-87b1-796ef9a85402.JPG)

![All stock 2018](https://user-images.githubusercontent.com/115379848/207991917-899c0b25-b21c-4996-a385-aadb7262d3d0.JPG)



## Summary: 
### In a summary statement, address the following questions.



### What are the advantages or disadvantages of refactoring code?


### How do these pros and cons apply to refactoring the original VBA script?
