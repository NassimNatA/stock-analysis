# Stock Analysis Using VBA 
## Overview
The purpose of this analysis is to produce a comprehensive examination of stock data to find the total daily volume and yearly return for each stock. Since this analysis was initially done with a smaller subset, the analysis below is an expansion of the dataset to include the entire stock market over the past couple years. This analysis was executed by refactoring from the inital examination of the subset data, and to determine the efficiency of refactoring VBA code and measuring performance by script run time. 

---
## Results 
Refactoring of code was executed as below. 

**Step 1a/1b: Create a tickerIndex variable and set it equal to zero before iterating over all the rows.** 

Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

   `Dim tickerVolumes(12) As Long
    Dim tickerstartingPrices(12) As Single
    Dim tickerendingPrices(12) As Single 
    tickerIndex = 0`
    
---
**Step 2a: Create a for loop to initialize the tickerVolumes to zero.**

   `Loop to initialize tickerVolumes to 0
    For i = 0 To 11
    ticker = tickers(tickerIndex)
    tickerVolume = 0`
    
---

**Step 2b/3a: Use this for loop to loop over all the rows and write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.**

`For j = 2 To RowCount
     If Cells(j, 1).Value = tickers(tickerIndex) Then
               tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
               End If`
               
---

**Step 3b: statement to check if the current row is the first row with the selected tickerIndex. If it is,  assign the current closing price to the tickerStartingPrices variable.**

 `If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
 tickerstartingPrice = Cells(i, 6).Value
               End If`
               
 --- 
 
**Step 3c: statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable**

`If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = ticker(tickerIndex) Then
startingPrice = Cells(j, 6).Value
 End If`
 
 --- 
 
**Step 4a: Use a for loop to loop through your arrays to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.**

  `Worksheets("AllStocksAnalysis").Activate
       Cells(4 + i, 1).Value = ticker(i)
       Cells(4 + i, 2).Value = tickerVolume(i)
       Cells(4 + i, 3).Value = tickerendingPrices(i) / tickerstartingPrices(i) - 1`
---
## Conclusion

 Although refactoring was done per above with original code, comprehensive analysis of the dataset for year 2017 and 2018 could not be obtained due to further debugging required in the code.
 
 ---
## Summary
### Advantages and disadvantages of refactoring code.

Refactoring code is advantegous for reusing existing code and applying it to larger datasets. This can allow scripts to be multipurpose and cross-functional across different project. This can also be time-efficent rather than starting a new script for every new scenario or question being asked. A disadvantage is the risk of missing variables to change or correct for when recfactoring a new code for a new dataset, which may lead to more debugging required. Likewise without careful documentation on github refactoring can lead to overwriting valuable code. 

---
### Advantages and disadvantages of the original and refactored VBA script 

Advantages of the original VBA script is that a smaller dataset allowed us to manipulate and format the data more easily. With only a small number of datapoints, buttons/color-coding can be adding to easier user access. A disadvantage to the original VBA script is that it was not comprehensive in informing us about the stock data across years 2017,2018. 

-
Advantage of the refactored VBA script is that it does give us a thorough look at sotkc data across multiple years and data points. A disadvantage is that the processing time can be assumed to be slower in VBA given the larger about of data to manipulate. 
