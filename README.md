# Stock-Analysis
## Overview of Project: 
I take pride in trying to simplify the lives of others. The stock analysis project is aimed at helping Steve’s parents to research and expand the dataset to include the entire stock market over the last week

## Explain the purpose of this analysis.

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

## Results 


## compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
While looking at the results from the analysis I confirmed that Steve's parents will view the stock results quicker in 2017 than 2018 as shown in the screenshot


![VBA_Challenge_2018](https://user-images.githubusercontent.com/115379848/208025840-6ee5bd73-b9d5-4c77-92d5-bbbbc77a1858.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/115379848/208025850-87a531cc-b3bf-449e-b220-023113842f36.png)


![All stock 2017](https://user-images.githubusercontent.com/115379848/207991801-3f92ea2f-75d1-4686-87b1-796ef9a85402.JPG)

![All stock 2018](https://user-images.githubusercontent.com/115379848/207991917-899c0b25-b21c-4996-a385-aadb7262d3d0.JPG)



## Summary: 
### advantages of refactoring code?
Logical error easily appear in well structure code

The VBA interpretation code can reveal patterns that are not easily seen in the source

Refactoring code decreases micro time run from 0.679 for 2017 and 0.671 for 2018 to 0.203 and 0.429 respectively


![VBA_Challenge_2017](https://user-images.githubusercontent.com/115379848/208019018-804110b5-509e-446e-b927-63d51991477d.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/115379848/208019027-7a48a29e-842f-4868-b972-ad8c8d5d5d38.png)

### Disadvantages of refactoring code?
Refactoring can ffect the testing outcome

### How do these pros and cons apply to refactoring the original VBA script?
Refactoring helps to make codes cleaner and better data origanized which improves the efficiency of the programmming which becomes easy for our users to view and read.
Unfortunately due to luck of proper test cases for existing codes we are not get the luxury of doing it.
