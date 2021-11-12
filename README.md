# Stock Analysis
the purpose of this analysis is to evaluate performance of different types of tickers in stock market in 2017 & 2018.

## Technical points
This analysis is done based on Total volume, Return in 2017 & 2018 - Total volume means a sum of Volumes and Return means ending price / starting price -1 .
We have 12 different types of tickers,  therefore,  tickers is defined as an array; 
 - (Dim tickers(12) As String)
Volume, Endingprice and Startingprice are initialized as a array because we need to have 12 of each ones for each tickers;
 - Dim tickerVolumes(12) As Long
 - Dim tickerStartingPrices(12) As Single
 - Dim tickerEndingPrices(12) As Single
First, we initialize tickerVolumes(tickerindex) = 0 before starting our calculation as an index.
our stock data that we have in this project are sorted by time and tickers so starting price of each tickers is first price of its first row and its ending price is the price of last row of each tickers. So if- then is used to find first row and last row of each tickers.
The final results are shown in a table based on our calculation of each year. the headers of table are defined; 		 
 - Cells(2, 2).Value = "Ticker"
 - Cells(2, 3).Value = "Total Daily Volume"
 - Cells(2, 4).Value = "Return"
The below codes prepare our result in the table. For loop is used;		
 For tickerindex = 0 To 11
   - Cells(tickerindex + 3, 2) = tickers(tickerindex)
   - Cells(tickerindex + 3, 3) = tickerVolumes(tickerindex)
   - Cells(tickerindex + 3, 4) = tickerEndingPrices(tickerindex) / tickerStartingPrices(tickerindex) - 1
 Next tickerindex
 
## Refactoring code
Refactoring code means the process of restructuring existing computer code—changing the factoring—without changing its external behavior. Refactoring improves the design of code, we make it easier and easier to work.
  ### Benefits of Refactoring 
Simplified support and code updates. Clean code is much easier to update and improve.
Saved time and money in the future. 
Reduced complexity for easier understanding.
Maintainability and scalability.

## Result
Based on our calculations and result, In 2017 all tickers have positive return except TERP. In 2018 ENPH & RUN have positive return. In 2018, ENPH volume becomes 2.7 times compare with its volume in 2017 and its return increase from 5.5% to 84%. So based on performance of RUN ticker in 2017 & 2018, it is a good ticker for investing on it.



