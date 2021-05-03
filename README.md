# Green Stock Aanalysis

## Overview
This analysis uses VBA to evaluate the performance of 12 green stocks using data from 2017 and 2018. The dataset contains daily historical data for each stock. I used loops and conditionals to run through the rows and extract the data I was looking for. One the code ran smoothly, I refactored the code to improve run time, in hopes that it could be used for hundreds of stocks at once. 

## Purpose
The purpose of this analysis is to easily compute the performance of 12 green stocks, specifically I wanted to know the total trading volume per year and yearly return.  

## Results 

### Stock Analysis 

After running the analysis on both 2017 and 2018, it appears that the green stocks overall did better in 2017 than in 2018. 

![Screen Shot 2021-04-30 at 3 43 37 PM](https://user-images.githubusercontent.com/80648379/116833629-541caf00-ab88-11eb-8c04-dd7b4146b2b0.png)

![Screen Shot 2021-04-30 at 3 43 08 PM](https://user-images.githubusercontent.com/80648379/116833631-567f0900-ab88-11eb-95e6-12677fe02ce2.png)

In 2017 we can see that 11 out of 12 of the stocks had a positive yearly return. While in 2018, only two stocks, ENPH and RUN had positive returns. The only stock that had a positive return in both years was ENPH. 
 
### Code Analysis

The original code took about 0.8 seconds to run for both 2017 and 2018. 

![Screen Shot 2021-04-29 at 5 45 11 PM](https://user-images.githubusercontent.com/80648379/116833639-5f6fda80-ab88-11eb-8eba-ebee8fe34462.png)

![Screen Shot 2021-04-29 at 5 44 52 PM](https://user-images.githubusercontent.com/80648379/116833641-61399e00-ab88-11eb-9eaa-ddfbf1503332.png)

To improve the run time, I refactored the code to only run through the data once. In the first code I used a nested loop. The first for loop looped through the tickers, and for each ticker we used the second loop to loop though the data again to find the Total Volume, Starting Price and Ending Price. 
```
For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0

            Worksheets("2018").Activate

            For j = 2 To RowCount

                If Cells(j, 1).Value = ticker Then
                    totalVolume = totalVolume + Cells(j, 8).Value
                End If

                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                    startingPrice = Cells(j, 6).Value
                End If

                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                    endingPrice = Cells(j, 6).Value
                End If
         Next j
Next i
```
To improve this code we created three arrays for Ticker Volumes, Starting Prince and Ending Price, as well as a ```tickerIndex``` variable that is used to access the correct ticker index in each array. The ```tickerIndex``` variable enables us to only run through the data once, which cuts down on the run time. 

```
For i = 0 To 11
        tickerVolumes(i) = 0
Next i

For i = 2 To RowCount

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    
            tickerIndex = tickerIndex + 1
            
        End If
    
 Next i
```
When using the refactored code and looping through the data once, it only takes about 0.15 second to run for each year. 

![VBA Challenge 2017](https://user-images.githubusercontent.com/80648379/116833648-6dbdf680-ab88-11eb-8a77-7df3d6aca80a.png)

![VBA Challenge 2018](https://user-images.githubusercontent.com/80648379/116833651-70205080-ab88-11eb-9b28-9392d95074a7.png)

## Summary 

### Advantages of refactoring 

An advantage to refactoring code is a quicker run time, in this case the code ran about 81% faster. Another advantage is shortening and simplifying the code, making is easier to read. I was able to make the code simpler by only having to run through the data once instead of twice.  

### Disadvantages of refactoring 

One of the main disadvantages of refactoring code is the time it takes to refactor. If the code you have runs perfectly well, it might not be worth the time to refactor it. With only evaluating 12 stocks, ran about 0.7 seconds faster and might not have been worth the time and effort to refractor. 
![image](https://user-images.githubusercontent.com/80648379/116833611-40714880-ab88-11eb-92b4-309b366ca73f.png)
