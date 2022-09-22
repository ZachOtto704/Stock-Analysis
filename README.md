Stock Analysis for Green Energy Stocks

Overview of Project:
The purpose of this project was to help our client, Steve, analyze multiple green energy stocks. We used a data file that held trading information for 12 different stocks, with data from each trading day of the year for each stock. We sorted through all this data by giving Steve the total yearly return for each individual stock, and the total volume for the year for each individual stock. We did this by creating macros that sorted through the data and returned the information we wanted on our output sheet. After creating a macro script to run our test, we also created a refactored code that ran the same test even more efficiently. 

Results:
Although, both macros returned the same data, the refactored code ran much more efficiently. Looking at the run times below, we can see the difference.

Original Code Performance:
![non_refactored_code_performance](https://user-images.githubusercontent.com/112716673/191826534-7c8154c3-6275-456a-83f0-99c4aa9d0310.png)

Refactored Code Performance:
![refactored_code_performance](https://user-images.githubusercontent.com/112716673/191826628-b8e924fc-566c-4f0e-85de-2e5a3644a256.png)

Here, we can see that the refactored code ran in 1/10th of the time the time the original code ran. 

Differences in Code:
The main difference between the original and the refactored code is as follows. The original code loops through all the tickers, and then loops through all the rows of data, and correlates the data in the spreadsheet to the appropriate ticker. For example, 

   For i = 0 to 11
     ticker=tickers(i)
     
Tells excel to make 12 calculations, one for each stock, and 

   totalVolume = totalVolume + Cells(j, 8).Value
   
then adds the volume for each trading day to calculate a total volume relating to each stock. As the trading volume is stored in the 8th column (hence j,8)

However in the refactored code, we used the same initial loop of 

For i = 0 To 11
    ticker = tickers(i)
    
And then created an output array for the ticker volumes as well. 

     tickerVolumes(i) = 0
     
In order to make this code work, we needed a variable to relate back to our ticker index. So we created a tickerIndex variable and set it to zero.

     tickerIndex = 0
     
And then when we calculated stock volume in our output array, we used the tickerIndex variable to move through the stock tickers in the index:

     tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
     
Since tickerIndex is a static variable, we had to create a script to increase the index to move from one ticker to the next. We did this with:

     tickerIndex = tickerIndex + 1


For the stock starting and ending prices we had to use the script:

     If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
     tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
     
This tells excel that if the value of the data in the ticker column is the equal to the ticker we are looking for, AND the ticker after it (j+1) is different, then this is the last trading day for that specific ticker. So the ending price for the year would be the value of the ticker's closing price, stored in the 6th column of the spreadsheet ((j,6).value). 

Summary / Advantages and Disadvantages:
As we stated before, both codes work, and give us the correct output for the desired calculations we wanted. However the refactored code works much more efficiently. So an obvious advantage of refactoring a code, is that you make the program run more efficiently. However, it also takes a lot of work, and can get confusing at times. In trying to refactor the code I ran into many, many errors. And the code now runs 1 second faster. So I would say in this case in VBA, it is okay to use the original code if it is giving us the correct output. However, in a much larger project, it would definitely be necessary to refactor the code. In this case the code ran 1 second faster, but that is also 10x faster than the original code. Imagine a large code that has a lot to process. Using this same 10x difference, imagine if running a program took 2 minutes or 120 seconds. Refactoring the code could result in the code running in 12 seconds instead. If there are many codes that need to be ran in a program, the program will run much smoother if the code is refactored.
