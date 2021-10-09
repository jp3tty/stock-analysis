# Green Stock Analysis
Performing analysis on stocks with Excel and VBA.

## Overview of Project

The goal of this project is to create an efficient, user-friendly Excel interface for performing stock analysis. To manage this task, a VBA subroutine was employed that reports an array of stock performaces ("Total Daily Volume" and "Yearly Return Percentage") for a given year. As an input, the subroutine prompts the user to pick the year they wish to analyze. The data is then processed and the desired information is output to its own Excel worksheet for the user to view.

# Analysis and Challenges

The VBA subroutine created to perform this analysis steps through the following operations:

1. Format the report worksheet:
   * add a title of evaluated year.
   * add "Ticker, Total Daily Volume, and Return" headers
3. Reference an array of tickers for evaluation.
4. Execute a For loop to total the value of a given ticker index.
5. Submit the results of the For loop to the report worksheet.
6. Format the report worksheet:
   * make the headers text bold
   * include a line to separate the headers form their associated rows of results
   * conditionally format the results with green or red cells depending on the stocks performance

For ease of use, the report sheet was given two buttons as a user interface. The first button, labeled "Perform Stock Analysis," initializes the steps listed above. The second button, labeled "Clear Worksheet," initializes a separate subroutine that erases the information from the report worksheet.

A report for 2017 can be seen here:

![Output for 2017](https://github.com/jp3tty/stock-analysis/blob/main/Resources/2017_analysis_OG.PNG)

A report for 2018 can be seen here:

![Output for 2018](https://github.com/jp3tty/stock-analysis/blob/main/Resources/2018_analysis_OG.PNG)

The buttons to initialize and clear the analysis can be seen in both images. Added to the subroutine is an ability for it to measure it's own performance in terms of runtime. We see, from the above images that the 2017 analysis took 0.9726563 seconds to run and 2018 took 0.890625 seconds.

The code used to perform this analysis used a nested For loop shown below:

![OG All Stock Analysis](https://github.com/jp3tty/stock-analysis/blob/main/Resources/OG_code.PNG)

## Refactoring the VBA Subroutines

Refactoring was applied in an attempt to improve the performance of the VBA subroutine. In this case, the code indexes a given ticker to determine its yearly performance. The changes can be seen here:

![Refactored For Loop](https://github.com/jp3tty/stock-analysis/blob/main/Resources/refactoredNestedLoop(2).PNG)

## Results
The ticker performance per year remained unchanged between the original code and the refactored code, but codes efficiency is greatly improved due to refactoring. This can be seen in the 

![Refactored2017CodePerformance]

![Refactored2018CodePerformance]

# Summary

## Advantages of refactoring:
* It can improve the performance of a subroutine, and is a key part of the coding process.
* It helps make the code appear cleaner, making it easier to understand.

 ## Disadvantages of Refactoring:
* It can be time consuming.
* It could alter the subroutines outcomes.
* It might be difficult to refactor large subroutines.
