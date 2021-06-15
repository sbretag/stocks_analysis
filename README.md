# Stocks-Analysis

## Overview of Project

### Background & Purpose
The purpose of this project is to refactor the original VBA script put together to analyze stocks found in Steve's parents portfolio. The original workbook prepared for Steve contained a VBA script that analyzed a dozen stocks looping through the entire data set per each stock.  Steve is looking to expand his reserach so he can analyze more stocks in an efficient manner.  The original script needs to be refactored in a way that the entire stock market can be analyzed if necessary using one loop versus looping throw the entire data set for each stock.  A final determination needs to be reached on whether the refactored script performed more efficiently than the original script all the while confirming the new output still matches the original output.

## Results

### Stock Performance

#### 2018 vs 2017 Stock Performance
Table 1 below compares total daily volume and returns twelve different stocks contained in the data set.  It also includes analysis of what the value of $100 investment would of been at the beginning of 2017 through the end of 2018.   Overall, a positive return would of been made on all but three different stocks which are JKS, SPWR, & TERP.  The only stocks with positive returns in both years are ENPH and RUN which would of netted you the 1st and 3rd best returns in two years overall.  Although SEDG had a negative return in in 2018, the high positive return in 2017 gave this stock the 2nd best return over the two years.   ENPH, SEDG, and RUN also had the biggest increases in volumes between 2017 and 2018 possibly signaling strong and rising markets for the particular stock.  Finally, if you would of invested $100 in all stocks at the beginning of 2017 through the end of 2018, your net gain would of been 56.6%

##### Table 1


##### Data Analysis File
 [Go to Analysis (Refactored Script)](https://github.com/sbretag/stocks_analysis/blob/main/VBA_Challenge.xlsm)

### Analysis Execution Times
As mentioned in the overview, the analysis of the stock performance was ran using two different VBA scripts.  The table 2 below show the execution times of the original scripts for both 2017 & 2018 data sets and table 3 shows the execution times using the refactored script for both 2017 & 2018 data sets.  As you can see the refactored script ran faster which can be attributed to only having to loop through the entire dataset once, versus looping through the data set once for every stock.  The slight difference in timing would be more apparent if the data set contained thousands of stocks versus 12.

#### Table 2


##### Data Analysis File w/Original Script
 [Go to Analysis w/Original Script](https://github.com/sbretag/stocks_analysis/blob/main/resources/green_stocks.xlsm)

#### Table 3





## Summary

### Advantages & Disadvantages of Refactoring Code in General

#### Advantages
1. placeholder
2. placeholder

#### Disadvantages
1. placeholder
2. placeholder

### Advantages & Disadvantages of the original and refactored VBA script

#### Advantages
1. placeholder
2. placeholder

#### Disadvantages
1. placeholder
2. placeholder

