# A ***Brief*** Analysis of the Benefit and Disadvantage of Refactoring VBA Code

## Overview of Project
#### The purpose of this project was to determine what the benefit or disadvantage might be of refactoring ***working*** VBA code.  
Given a data set of the annual performance of 12 "green" stocks, a VBA code script was written in which over 3,000 rows of data were able to be calculated and summarized according to each stock's ticker at the click of a button.  

Upon been prompted by a user ***input box*** to enter the value for the year to be reported, a timer was initiated which calculated the process time of **collection**, **storage**, **output**, and **formatting** of the data being recorded on a separate sheet.  A message box was then displayed with the run time of the script.

A VBA Subroutine, or *Macro*, was written to hold each stock's ticker name in an array and then loop through every row of data and output the **ticker name**, **total volume**, and **return** for the year.  Because the script ran through every row of data for each ticker, a total of over *36,000* rows were processed each time the button was clicked.  As a result of refactoring the working code, we were able to achieve the same results by creating three ***output*** arrays which stored the data while the remaining rows were processed.  This resulted in the refactored code running through the rows of data one time for a total of just over ***3,000*** rows.

## Results


## Summary
#### Advantages and Disadvantages of Refactoring Code
Refactoring code can offer a variety of advantages.  The resources needed to process lines of code are the memory and processing speed of the computer running that code.  The ability of a computer to process code quickly and efficiently is based on a number of factors including, but not limited to:  total amount of data being read, the number of process being run over the range of that data, additional programs running concurrently, and the amount of memory available.  Because of the finite resources available within the computing power of a machine, it can be a benefit to find ways to simplify the code being run.  Because future users of code may also need to make changes, writing comments within the code is also important.  These future users may not think the same way as the original author of the code, so writing code logically is equally important.  

Depending on the circumstances, a particular subroutine of code may not be used very often.  In this instance it might not be beneficial to work exhaustively through the code in order to tweak its efficiency or readability.  "If it works, it works."  

#### Advantages and Disadvantages of the Original and Refactored VBA Script Used in This Analysis
The sample size of data in this analysis was very limited.  It was over a very few number of stocks and over a relatively short period of time.  Because of this, the resulting delta in processing time was negligible as seen in the screenshots below.  However; because it's clear from the images that the run time for the code prior to refactoring code was ***almost 6 times longer*** than after the refactor, it can clearly be seen that over a larger data set the overall processing time could be much more significant than just fractions of a second.  Specific to applications of data analysis like the stock market, split second decisions need to be made and the processing speed of a subroutine can be a significant benefit or detriment to the user of that code.

Depending on the size of data to be processed by a code and the frequency in which it will be run, perhaps it doesn't offer enough benefit to sift exhaustively through that code just to save a fraction of a second.  For code that is relatively few lines in length refactoring may not be beneficial overall, but for code that contains many lines refactoring may in fact prove to be quite beneficial.






