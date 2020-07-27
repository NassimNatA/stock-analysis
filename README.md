# Stock Analysis Using VBA 
## Overview
The purpose of this analysis is to produce a comprehensive examination of stock data to find the total daily volume and yearly return for each stock. Since this analysis was initially done with a smaller subset, the analysis below is an expansion of the dataset to include the entire stock market over the past couple years. This analysis was executed by refactoring from the inital examination of the subset data, and to determine the efficiency of refactoring VBA code and measuring performance by script run time. 

--
## Results 
Refactoring of code was executed as below. 
*Step 1a/1b: Create a tickerIndex variable and set it equal to zero before iterating over all the rows.* 
*Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.*
`Dim tickerVolumes(12) As Long
    Dim tickerstartingPrices(12) As Single
    Dim tickerendingPrices(12) As Single
    tickerIndex = 0`

*Step 2a: Create a for loop to initialize the tickerVolumes to zero.*
`Loop to initialize tickerVolumes to 0
    For i = 0 To 11
    ticker = tickers(tickerIndex)
    tickerVolume = 0`

*Step 2b/3a: Use this for loop to loop over all the rows and write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.*
`For j = 2 To RowCount
     If Cells(j, 1).Value = tickers(tickerIndex) Then
               tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
               End If`

*Step 3b: statement to check if the current row is the first row with the selected tickerIndex. If it is,  assign the current closing price to the tickerStartingPrices variable.*
 `If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
 tickerstartingPrice = Cells(i, 6).Value
               End If`
               
*Step 3c: statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable*
`If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = ticker(tickerIndex) Then
startingPrice = Cells(j, 6).Value
 End If`
               
*Step 4a: Use a for loop to loop through your arrays to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.*
  `Worksheets("AllStocksAnalysis").Activate
       Cells(4 + i, 1).Value = ticker(i)
       Cells(4 + i, 2).Value = tickerVolume(i)
       Cells(4 + i, 3).Value = tickerendingPrices(i) / tickerstartingPrices(i) - 1`

               
