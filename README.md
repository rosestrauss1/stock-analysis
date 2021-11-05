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

This utilized an if then statement to find if the ticker value was 

To find the ending price, I used a similar process: 

 If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        endingPrice = Cells(j, 6).Value

This also used an if then statement to 

In the refactored code, the main difference was that instead of looping over all data in the data set, I checked to see if the current row was the first or last row of a ticker index. This made it so I only had to loop through the data one time, because I was able to obtain the starting and ending price without iterating through the data.

## Summary

### Advantages and Disadvantages of Refactored Code
