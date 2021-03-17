# A ***Brief*** Analysis of the Benefit or Disadvantage of Refactoring Code

## Overview of Project
The purpose and background are well defined (2 pt).
#### The purpose of this project was to determine what the benefit or disadvantage might be of refactoring working VBA code.  
Given a data set of the annual performance of 12 "Green" stocks, a VBA code script was written in which over 3,000 rows of data were able to be calculated and summarized according to each stock's ticker at the click of a button.  

Upon been prompted by a user ***input box*** to enter the value for the year to be reported, a timer was initiated which calculated the **collection**, **storage**, **output**, and **formatting** of the data being recorded on a separate sheet.

Code was written to loop through every row of data based on an array of the 12 stock tickers and output the **ticker name**, **total annual volume**, and **return**.  Because the script ran through every row of data for each ticker, a total of over 36,000 rows were processed.  By collecting the same information through the process of creating output arrays the same outcome was achieved by refactoring the script to allow the code to process through the entire 3,000+ rows of data one time, collect the data of the output arrays (according to their index) and speed up the entire process. 
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

Structure, Organization, and Formatting Requirements (8 points)
The written analysis contains the following structure, organization, and formatting:

There is a title, and there are multiple paragraphs (2 pt).
Each paragraph has a heading (2 pt).
There are subheadings to break up text (2 pt).
Links are working, and images are formatted and displayed where appropriate (2 pt).








