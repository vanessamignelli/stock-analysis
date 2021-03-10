# Green Stocks Analysis

## Overview of Project

### The purpose of this project is to examine the total daily volume and yearly return for each of the green stocks provided in the existing excel document, to assist in determining the best option to invest in. The goal was to create a fast and efficient way to quickly run this analysis using VBA. The full analysis can be found here [VBA_challenge](https://github.com/vanessamignelli/stock-analysis/blob/main/VBA_challenge.xlsm).

## Results

### Stock Results
Overall, stocks performed considerably better in 2017 compared to 2018. In 2017, approximately 91% of the green stocks analyzed had a postive return rate. The average daily trading volume was 263,886,592 and the total daily trading volume across all stocks for 2018 was 3,166,639,100. The only two most profitable stocks to invest in in 2017 were DQ which offered a return of 199.4% and SEDG which offered a return of 184.5%. Only one stock, TERP, offered a return rate that was below zero (see image below).

![VBA_ChallengeOutput_2017](https://github.com/vanessamignelli/stock-analysis/blob/main/Resources/VBA_ChallengeOutput_2017.png)

In comparision to 2018, approximately 82% of the green stocks analyzed had a return rate of less than zero, indicating investors would have lost money. The average daily trading volume was 275,503,183 and the total daily trading volume across all stocks for 2018 was 3,306,038,200. The only two profitable stocks to invest in in 2018 were ENPH which offered a return of 81.9% and RUN which offered a return of 84.0% (see image below). 

![VBA_ChallengeOutput_2018](https://github.com/vanessamignelli/stock-analysis/blob/main/Resources/VBA_ChallengeOutput_2018.png)

### Code Results
The initial code developed to run this analysis worked, but had room for efficiencies. Originally, only one array was utilized to run a loop and access the list of stocks, to get the corresponding total volumes and starting and ending price to be able to calculate the rate of return. With this format, the code was able to run in just over a second (see images below).

![prefectoredCode_2017](https://github.com/vanessamignelli/stock-analysis/blob/main/Resources/prerefactoredCode_2017.png)
![prefectoredCode_2018](https://github.com/vanessamignelli/stock-analysis/blob/main/Resources/prerefactoredCode_2018.png)

The refactored code offered a different approach. Here, four arrays were created and only one index was used to access values in all four of the arrays. Additionally, another loop in accordance with an if-then statement, to automatically increase it's value at the end of the loop. This allowed for the next value in the index to be accessed and evaluated. The new refactored code was able to run in 0.148 seconds (see image below).

![VBA_Challenge_2017](https://github.com/vanessamignelli/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/vanessamignelli/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Pros and Cons of Refactoring Code
One of the benefits of refactoring code includes the ability to focus on creating efficiencies. Since there is a base or template already in place, there wouldn't be any additional time spent creating script to select or format cells, instead the focus may be on how to optimize more complex or confusing areas of the code. One of the cons related to refactoring code is you may need to take some additional time to read and understand the code that has been written prior. If there is a lack of appropriate comments/descriptions, it may be difficult to understand what the code is trying to accomplish, or how the ouput is being generated. 

### Pros and Cons of the Initial vs. Refactored VBA script
One of the advantages of the original VBA code created for this analysis, was it was more intuiative. When viewing the code, it is very evident that there is one array and one corresponding index, and it is easy to understand how the individual stocks are being accessed in the first for loop, using the ticker variable. Another benefit to the original VBA code is there are no magic numbers. Everything is clearly defined, and it is easy to see why certain numbers are being used. The con to the original code is it is not optimized. There is room to create a faster more efficient code that can better filter through the date. With regards to the refactored code, one of the benefits is the exclusion of nested for loops, which can be somewhat difficult to follow or understand. Additionally, there is the addition of the tickerIndex variable that is used to access all 4 arrays in the VBA code. The downside to this refactored code is in somewhat contains a magic number. When creating the volumes, starting price and ending price array, they are denoted to show 12 values contained in them. It may not be entirely clear when viewing the code where the number 12 came from.
