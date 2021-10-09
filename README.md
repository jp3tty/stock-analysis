# Green Stock Analysis
Performing analysis on stocks with Excel and VBA.

## Overview of Project

The goal of this project is to create an efficient, user-friendly Excel interface for performing stock analysis. To manage this task, a VBA subroutine was employed that reports an array of stock performaces (Total Daily Volume and Yearly Return Percentage) for a given year. As an input, the subroutine prompts the user to pick the year they wish to analyze. The data is then processed and the desired information is output to its own Excel worksheet for the user to view.

# Analysis and Challenges

The VBA subroutine created to perform this analysis steps through the following operations:

1. Format the report worksheet:
   * add a title with the year that is being evaluated.
   * "Ticker, Total Daily Volume, and Return" headers
3. Reference an array of tickers for evaluation.
4. Execute a For loop to total the value of a given ticker index.
5. Submit the results of the For loop to the report worksheet.
6. Format the report worksheet:
   * make the headers text bold
   * include a line to separate the headers form their associated rows of results
   * conditionally format the results with green or red cells depending on the stocks performance

For ease of use, the report sheet was given two buttons as a user interface. The first button, labeled "Perform Stock Analysis," initializes the steps listed above. The second button, labeled "Clear Worksheet," initializes a separate subroutine that clears the erases the information from the report worksheet.

@ INCLUDE IMAGE OF STOCK PERFORMACE FOR BOTH YEARS
@ INCLUDE IMAGE OF INITIAL CODE
@ INCLUDE IMAGE OF MODIFID CODE

## Measuring the VBA Subroutines Performance

@ INCLUDE IMAGE OF ORIGINAL FOR LOOP

## Improving the VBA Subroutines Performace

@ INCLUDE IMAGE OF UPDATED FOR LOOP

## Challenges and Difficulties Encountered

Having minimal experience with VBA not knowing where to activate worksheets and place variable callouts.

# Results
The ticker performance per year performance remained unchanged between the original code and the refactored code. The code efficiency is greatly improved however in the refactored code. 

'[Performance of Refactored 2017 Code](https://github.com/jp3tty/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
'[Performance of Refactored 2018 Code](https://github.com/jp3tty/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

@ IMAGE OF PERFORMACE OF ORIGINAL CODE
@ IMAGE OF PERFORMACE OF MODIFIED CODE

# Summary

