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
- text

### Pros and Cons of the Initial vs. Refactored VBA script
- text
