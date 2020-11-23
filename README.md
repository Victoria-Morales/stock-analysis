# Analyze Stock Data with VBA

## Overview of Project
### This project uses VBA in excel to automate analysis of a stock file for Steve. After providing our initial anaylsis to Steve, he now needs the script to be refactored and run faster. 

### Purpose
#### The purpose of this project is to provide analysis for Steve of green energy stocks based on 2017 and 2018 data. Steve's parents have invested in DAQO stock, so Steve would like an analysis of their stock as well as other green energy stocks, so that he will be able to advise them properly. The the script runs fairly quickly for this small amount of stocks, he needs the script to be refactored to run more efficiently in case he has a larger amount of data.


## Results

#### Comparison of stock performance between 2017 and 2018

Overall, the green stocks we analyzed performed much better in 2017 than in 2018 in terms of yearly return. The two stocks that had positive returns both years are ENPH and RUN. Additionally, these two stocks increased their daily trading volume.
Notably, the DQ stock that Steve's parents invested in, had a significant loss in 2018.


![2018](/Resources/2018Chart.png)


#### Execution time of original script and refactored script

Refactoring the script, significantly improved the run time.  See examples below for the comparison of run time.  

Original 2017

![2017Original](/Resources/VBA_Challenge_2017_Old_Way.png)

Refactored 2017

![2017Refactored](/Resources/VBA_Challenge_2017.png)

Original 2018

![2018Original](/Resources/VBA_Challenge_2018_Old_Way.png)

Refactored 2018

![2018Refactored](/Resources/VBA_Challenge_2018.png)

## Summary
There are both advantages and disadvantages to refactoring code.  The advantage of refactoring code is that it can make your code more flexible.  It can also make your code more efficient and run faster. A disadvantage of refactoring code is that is can be very time consuming. In regards to this project, these advantages and disadvantages definitely apply. Refactoring the code did make it run quicker.  However taking the time to refactor the script was very time consuming, and in this project, not a big enough payoff in terms of time saved to be worth doing.  However, if this script were to be used on a much bigger set of data, the refactoring process would be well worthwhile.



