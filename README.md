# A ***Brief*** Analysis of the Advantage and Disadvantage of Refactoring VBA Code

## Overview of Project
#### The purpose of this project was to determine the advantage or disadvantage of refactoring ***working*** VBA code.  
Given a data set of the annual performance of 12 "green" stocks, a VBA code script was written in which over 3,000 rows of data were calculated and summarized according to each stock's ticker at the click of a button.  

An ***input box*** prompted the user to enter the value for the year to be reported.  Upon entering that value, a timer was initiated which calculated the total elapsed time of the following processes:  **collection**, **storage**, **output**, and **formatting**.  These results were recorded on a separate sheet.  A message box was then displayed with the total run time of of those processes.

To accomplish this, a VBA Subroutine, or *Macro*, was written to hold each stock's ticker name in an array and then loop through every row of data and output the **ticker name**, **total volume**, and **return** for the year.  Because the script ran through every row of data for each ticker, a total of over ***36,000*** rows were processed each time the button was clicked.  By refactoring the working code, we were able to achieve the same results by creating three ***output*** arrays which stored the data while the remaining rows were processed.  This resulted in the refactored code running through the rows of data one time for a total of just over ***3,000*** rows.

## Results
#### By comparing the run times between the original code and the refactored code, we can measure the improvements in processing speed of the subroutine script used to process the data.  As seen in the images below it's clear that the refactored code ran ***almost 6 times quicker*** than the original script. 

The factors that contribute to this increase in time are the number of tickers and the overall number of rows of data.  This is because of the nature of the nested loops used in the original code.  To quote Colin Bartoe, a fellow classmate in this cohort:  *"You will be much more efficient if you have to run an obstacle course 1 time with a couple things to do at each station than to run the whole obstacle course for each obstacle, then do it again for the next one."*


***2017 Analysis run with original code script and with all data being reviewed for each unique ticker (over 36,000 rows)***

![2017_analysis.png](https://github.com/frostbrosracing/stock-analysis/blob/main/Resources/2017_analysis.PNG)


***2017 Analysis run with refactored code script and with each ticker being summarized as all the data was processed one time (just over 3,000 rows)***

![VBA_Challenge_2017.png](https://github.com/frostbrosracing/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)


***2018 Analysis run with original code script and with all data being reviewed for each unique ticker (over 36,000 rows)***

![2018_analysis.png](https://github.com/frostbrosracing/stock-analysis/blob/main/Resources/2018_analysis.PNG)


***2018 Analysis run with refactored code script and with each ticker being summarized as all the data was processed one time (just over 3,000 rows)***

![VBA_Challenge_2018.png](https://github.com/frostbrosracing/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

## Summary


#### Advantages and Disadvantages of Refactoring Code
Refactoring code can offer a variety of advantages.  The resources needed to process lines of code are the memory and processing speed of the computer running that code.  The ability of a computer to process code quickly and efficiently is based on a number of factors including, but not limited to:  total amount of data being read, the number of process being run over the range of that data, additional programs running concurrently, and the amount of memory available.  Because of the finite resources available within the computing power of a machine, it can be a benefit to find ways to simplify the code being run.  Because future users of code may also need to make changes, writing comments within the code is also important.  These future users may not think the same way as the original author of the code, so writing code logically is equally important.  Because of the small sample size of this data the difference in run times shown above is less than a second, but because the overall factor of time is based on a multiple rather than a constant, greater benefit will be seen over larger sample sizes. 

Depending on the circumstances, a particular subroutine of code may not be used very often.  In this instance it might not be beneficial to work exhaustively through the code in order to tweak its efficiency or readability.  "If it works, it works."  

#### Advantages and Disadvantages of the Original and Refactored VBA Script Used in This Analysis
The sample size of data in this analysis was very limited.  It was over a very few number of stocks and over a relatively short period of time.  Specific to applications of data analysis like the stock market, split second decisions need to be made and the processing speed of a subroutine can be a significant benefit or detriment to the user of that code.

Depending on the size of data to be processed by a code and the frequency in which it will be run, perhaps it doesn't offer enough benefit to sift exhaustively through that code just to save a fraction of a second.  For code that is relatively few lines in length refactoring may not be beneficial overall, but for code that contains many lines refactoring may in fact prove to be quite beneficial.






