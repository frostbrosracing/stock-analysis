# A ***Brief*** Analysis of the Benefit and Disadvantage of Refactoring VBA Code

## Overview of Project
#### The purpose of this project was to determine what the benefit or disadvantage might be of refactoring ***working*** VBA code.  
Given a data set of the annual performance of 12 "green" stocks, a VBA code script was written in which over 3,000 rows of data were able to be calculated and summarized according to each stock's ticker at the click of a button.  

Upon been prompted by a user ***input box*** to enter the value for the year to be reported, a timer was initiated which calculated the process time of **collection**, **storage**, **output**, and **formatting** of the data being recorded on a separate sheet.  A message box was then displayed with the run time of the script.

Code was written to hold each stock's ticker name in an array and then loop through every row of data and output the **ticker name**, **total volume**, and **return** for the year.  Because the script ran through every row of data for each ticker, a total of over *36,000* rows were processed each time the button was clicked.  As a result of refactoring the working code, we were able to achieve the same results by creating three ***output*** arrays which stored the data while the remaining rows were processed.  This resulted in the refactored code running through the rows of data one time for a total of just over ***3,000*** rows.

##

There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

Structure, Organization, and Formatting Requirements (8 points)
The written analysis contains the following structure, organization, and formatting:

There is a title, and there are multiple paragraphs (2 pt).
Each paragraph has a heading (2 pt).
There are subheadings to break up text (2 pt).
Links are working, and images are formatted and displayed where appropriate (2 pt).








