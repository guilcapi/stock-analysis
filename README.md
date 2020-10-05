# Stock Analysis showing Total Daily Volume and Return

## Overview of Project

### Purpose
Steve wants to at a click of a button analyze an entire dataset of stock information. We are presented with data for 2017 and 2018 stock daily volume and returns for a variety of stocks. 

Through a code refactor, I need to loop through all the provided data one time and see if this method runs fasters than my previous code.

## Results

The Stock Analysis script allowed comparison between 2017 and 2018 stock performance. Comparing the Module 2 script with the refactored script showed a variety of differences in efficiency and execution time. 

2017 showed that it was a good year for the majority of the stocks with some exceeding 100% returns. The highest return was the stock DQ at 199.4% while the lowest was at -7.2% returns.

![](/resources/VBA_Challenge_2017.png)

2018, however, showed a bad year for the majority of stocks though not as drastic as the positive returns from the previous year. Stock DQ which was the highest performer in 2017 was now the lowest performer with -62.6% returns. Stocks ENPH and RUN continued with positive returns this year at much higher returns than the previous year. 

![](/resources/VBA_Challenge_2018.png)

In terms of efficiency, the refactored code combined the formatting and stock analysis calculations into one sub statement allowing for processing with one sub statement rather than two. The execution time also decreased for both years when running the refactored script. Execution time went from 0.625 seconds for the year 2017 to 0.090 seconds. Below are the orginal script execution times. 

![](/resources/Previous_code_2017.png)

Execution time went from 0.633 seconds for the year 2018 to 0.086 seconds.

![](/resources/Previous_code_2018.png)

Within Module 2, Steve was doing the analysis of the stock market to see if stock DQ was a good investment for his parents and based on performance from 2017 and 2018, I would suggest that another stock be chosen to the high volatility in the returns but an analysis of more years would be advised. 

## Summary

- What are the advantages or disadvantages of refactoring code? The advantages of refactoring code are that it can improve efficiency in the script, take fewer steps, use less memory, and allow the script to be better understood with improved logic. Disadvantages with refactoring code can be that since it was not created by the new user, it may be hard to understand the original intent of parts of the script if the script was not commented well. 

- How do these pros and cons apply to refactoring the original VBA script? These pros and cons directly apply to refactoring the orignal VBA script because as was seen in this challenge, the efficiency of the script improved, processing speed improved, and had better logic ready for more data sets. The first attempt worked but it does not mean improvements could not have been made.
