# Stocks-Analysis

## Overview of Project

### Background & Purpose
The purpose of this project is to refactor the original VBA script put together to analyze stocks found in Steve's parents portfolio. The original workbook prepared for Steve contained a VBA script that analyzed a dozen stocks looping through the entire data set per each stock.  Steve is looking to expand his reserach so he can analyze more stocks in an efficient manner.  The original script needs to be refactored in a way that the entire stock market can be analyzed if necessary using one loop versus looping throw the entire data set for each stock.  A final determination needs to be reached on whether the refactored script performed more efficiently than the original script all the while confirming the new output still matches the original output.

## Results

### Stock Performance

#### 2018 vs 2017 Stock Performance
Table 1 below compares total daily volume and returns twelve different stocks contained in the data set.  It also includes analysis of what the value of $100 investment would of been at the beginning of 2017 through the end of 2018.   Overall, a positive return would of been made on all but three different stocks which are JKS, SPWR, & TERP.  The only stocks with positive returns in both years are ENPH and RUN which would of netted you the 1st and 3rd best returns in two years overall.  Although SEDG had a negative return in in 2018, the high positive return in 2017 gave this stock the 2nd best return over the two years.   ENPH, SEDG, and RUN also had the biggest increases in volumes between 2017 and 2018 possibly signaling strong and rising markets for the particular stock.  Finally, if you would of invested $100 in all stocks at the beginning of 2017 through the end of 2018, your net gain would of been 56.6%

##### Table 1
![](https://github.com/sbretag/stocks_analysis/blob/main/resources/VBA_Challenge_2018vs2017.png)


##### Data Analysis File
 [Go to Analysis (Refactored Script)](https://github.com/sbretag/stocks_analysis/blob/main/VBA_Challenge.xlsm)

### Analysis Execution Times
As mentioned in the overview, the analysis of the stock performance was ran using two different VBA scripts.  The table 2 below show the execution times of the original scripts for both 2017 & 2018 data sets and table 3 shows the execution times using the refactored script for both 2017 & 2018 data sets.  As you can see the refactored script ran faster which can be attributed to only having to loop through the entire dataset once, versus looping through the data set once for every stock.  The difference in execution times is just over a half of second which may seem immaterial however if this script was applied to a data set with thousands of stocks, the time would be relatively longer.  When trading in the stock markets, every second counts.

#### Table 2
![](https://github.com/sbretag/stocks_analysis/blob/main/resources/VBA_Challenge_OrigScript_2017and2018.png)

##### Data Analysis File w/Original Script
 [Go to Analysis w/Original Script](https://github.com/sbretag/stocks_analysis/blob/main/resources/green_stocks.xlsm)

#### Table 3
![](https://github.com/sbretag/stocks_analysis/blob/main/resources/VBA_Challenge_RefactoredScript_2017and2018.png)


## Summary

### Advantages & Disadvantages of Refactoring Code in General

In general, advantages of refactoring code include optimizing code to increase run times, implement bug fixes, and improving the overall readiability for future refactoring all while keeping the behavior the same.  Disadvantages include time, cost, and the risk of introducing new bugs. 

### Advantages & Disadvantages of the original and refactored VBA script

The original VBA script in this analysis had an advantage that it simply worked, even though the code was inefficient and had a higher run time than the refactored code, if you never planned to increase your data set to thousands of stocks then the half second of time savings would not be worth the time spent refactoring.  Especially if the analysis was done less frequently, like on an annual basis.  The clear disadvantage of the original VBA script would be if you planned to expand your analysis to a data set that contained thousands of stocks that you needed to run on a daily basis.  The time saved in running the code would most likely exceed the time spent on refactoring the code.  This type of planning should be taking into consideration before refactoring code.


