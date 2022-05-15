# Stock Analysis Using VBA


## Overview of Project

The goal of this project is to efficiently and effectively determine key information regarding various stocks related to green energy. This was accomplished using VBA Macros to gather information about the Annual Return Rate and Total Volume Traded for these stocks. The code was then refactored to run the program quicker and more efficiently.

## Results

### Analysis of 2017 Stocks

#### ORIGINAL VBA SCRIPT FOR 2017 STOCKS
![ian-zukowski](Original_VBA_2017.png)


#### REFACTORED VBA SCRIPT FOR 2017 STOCKS
![ian-zukowski](VBA_Challenge_2017.png)

Looking at the results from the 2017 stock performance it would appear that "DQ", "ENPH", "FSLR" and "SEDG" had the best performances in terms of Annual Return Rate. To find this annual return rate the final closing price of the year (from 12/29/17) was compared with the first closing price of the year (01/03/17). Each of the aforementioned stocks had a return of at least +100%, meaning the closing price on 12/29 was had at least doubled its value in the 360 days in between. The results also show that the only one of the analyzed stocks to actually decrease in value at all was "TERP", which only decreased by -7.2%.

In terms of Total Volume of stocks traded, "FSLR" and "SPWR" were the two most commonly traded stocks, each having a volume of over 500 million stocks traded over the course of the year. These values were found by going iteratively through the table and adding the volume of daily traded stocks to the previously counted stocks for that company.


### Analysis of 2018 Stocks

#### ORIGINAL VBA SCRIPT FOR 2018 STOCKS
![ian-zukowski](Original_VBA_2018.png)   


#### REFACTORED VBA SCRIPT FOR 2018 STOCKS
![ian-zukowski](VBA_Challenge_2018.png)

The results from the 2018 stock performance show a bleaker outlook than the previous years results. During this year the only two companies to profit at all were "ENPH" and "RUN". Both of these companies had a respectable +80% return rate, which is still a signifcant increase even though "ENPH" had previously achieved a +129.5% return in the previous year. Most companies by comparison had their stocks lose value in the 363 days that were recorded in 2018. A similar method was used to obtain these results, this time comparing the final closing price (12/31/2018) to the first closing price (01/02/2018).

In terms of Total Volume of stocks traded, "ENPH", "RUN" and "SPWR" were the most commonly traded stocks, all having a volume of over 500 million stocks traded over the course of the year. This volume in 2018 is a stark increase for both "ENPH" (an extra 400 million compared to 2017) and "RUN" (an extra 250 million compared to 2017). These values were again found by going iteratively through the table and adding the volume of daily traded stocks to the previously counted stocks for that company.


### Efficiency of Original Code Compared to Refactored Code
As can be seen in the pictures above, the refactored code was able to run quicker than the original code suggested by the readings in the module. For the 2017 code, the refactored version ran 1.28x quicker than the original. For the 2018 code, the refactored version ran 1.59x quicker than the original. 

The main difference in the codes is seen in the conditional statements that were used to find Total Volume, First Closing Values, and Final Closing Values. In the original program the code runs through the spreadsheet, gathers the relevant information for the first stock ("AY"), then activates the "All Stocks Analysis" worksheet and displays those results. After displaying those results, it returns to the worksheet with the daily information and then repeats the process for the next stock.

In the refactored code, Total Volume, First Closing Values, and Final Closing Values are all stored as arrays which will be associated with the stock ticker at the end of the process. This allows the code to do all neccessary work in the worksheet for 2017/2018 daily information without having to go back and forth with the "All Stocks Analysis" worksheet. Only after all of the values have been found does the program open up the worksheet to then display each of the relevant results.


### Code Used to Obtain Results

#### Code to find Total Volume of Traded Stocks
     For j = 2 To RowCount
        If Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        End If

This code goes through each row of data and adds in the daily volume of traded stocks (Cells(j,8).Value) if and only if the first cell in that row is the current ticker.


#### Code to find Annual Return
        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If

        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        End If

    For i = 0 To 11
        tickerIndex = i
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1

This code goes through the data and sets the "Starting Price" for the year (01/02/18 or 01/03/17) by finding the first row with the desired ticker value in cells(j,1). It then establishes the "Ending Price" for the year (12/31/18 or 12/29/17) by finding the last row with the desired ticker value, which must be right above a row with a different ticker value. Then finally to establish the Annual Return Percentage the EndingPrice is divided by the StartingPrice to find the growth/decay factor for the year. And finally that growth/decay factor is subtracted by 1 to establish the rate of growth/decay compared to the starting "100%" value which was the starting price.


## Summary: In a summary statement, address the following questions.

  What are the advantages or disadvantages of refactoring code?
  
  How do these pros and cons apply to refactoring the original VBA script?
