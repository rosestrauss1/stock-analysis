# VBA Challenge

## Overview of Project

### Purpose

The purpose of this project is to use VBA to refactor code that uses for loops to determine the total daily volume and return for different stocks. Originally, this process was done using VBA to loop through the lines of codes until the values were identified, but in this project I refactored the code to make the process faster by only looping through the values once.

## Results

### Comparison of Stock Performance Between 2017 and 2018

The stock performance in 2018 was worse than it was in 2017. This can be determined from the data because in 2017, most of the tickers had a positive return. This means that there was a positive percent increase from the starting price of the stock to the price the stock ended at. However, in 2018, almost all the tickers had a negative return. This can be easily visualized from the formatting included in VBA, which shows values that had a positive return in green and a negative return in red.

2017 Returns


<img width="300" alt="Screen Shot 2021-11-05 at 12 18 40 AM" src="https://user-images.githubusercontent.com/92963227/140472912-6f33132b-e3b1-4ca5-a728-2cf15ecaf10b.png">


2018 Returns


<img width="302" alt="Screen Shot 2021-11-05 at 12 16 16 AM" src="https://user-images.githubusercontent.com/92963227/140472819-8b5d3827-6552-4820-be77-0ac68716182a.png">



### Comparison of Original Script and Refactored Script

The original script and the refactored script were different in a few key ways. In the original script, the starting price and ending price for the tickers was determined by looping over all tickers in the data set. 

To find the starting price, the code I used was:
If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        startingPrice = Cells(j, 6).Value

To find the ending price, I used a similar process: 

 If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        endingPrice = Cells(j, 6).Value

This utilized an if then statement within a for loop to loop through all the rows of data and find the ending price. In both cases, the if then statement evaluted whether the cell before or after matched the ticker cell value, and if it did not, it was evaluated as either a starting or ending price.

In the refactored code, the main difference was that instead of looping over all data in the data set, I used an index value called tickerIndex which allowed me to iterate through the data faster. This was because instead of looping through each entire row to find starting price, ending price, and volume, I was able to move past rows that did not have the starting or ending values for the ticker.

## Summary

### Advantages and Disadvantages of Refactored Code
Refactoring the code made the VBA script run considerably faster than it did with the original code. For 2018, the regular script ran at a time of 0.67 seconds compared to 0.12 seconds. In 2017, the original script ran at 0.64 seconds and the refactored code ran at 0.12 seconds. The benefits of this is that the code is more efficient and could be easier for some people to understand. In my opinion, this efficiency would be a larger factor to consider if the script took a longer time to run, but because this script ran quickly, to me it was not worth refactoring the code because it took a very long time.
