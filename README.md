# stock-analysis

## Overview of Project
-This project is an analysis of the recent performance of a certain subset of stocks.
### Purpose
-The purpose of this analysis is to instruct considering buyers as to whether or not to purchase stocks from the set they were considering with money they set aside to invest.

## Results
-The stocks that this data would isntruct me to recommend purchasing shares of are NPH and RUN.
### Stock Performance Year Over Year
-Every stock had a negative return year over year with the exception of NPH and RUN, and both of those stocks also had positive returns in the year before.
### Original Script vs Refactored Script Performance
-The refactored script was significantly faster to run than the original script. Storing the values in an array and retrieving them seems to be a more efficient method for displaying that information than spitting the values out into the spreadsheet during each iteration in the parent for loop.

## Summary
### 1. What are the advantages or disadvantages of refactoring code?
-There certainly are some advantages to refactoring your code. One that comes to mind is the ability to make your code more readable to someone outside of the code's author. That, on top of the fact that you can make your code run more smoothly and efficiently, which is particularly important when working with larger data sets. You also gain the ability to see where and how joining multiple macros under one subroutine becomes useful, removing steps from the process for yourself should you reuse similar code down the line.
-As far as disadvantages, there are only really two that come to mind. The first of which is that you stand the chance of causing yourself a major headache or series of headaches in accidentally breaking your code. The second, and probably the more common concern, would be that of time. Refactoring code takes time at every step: analysis, planning, execution, and reimplementation. Sometimes time is a luxury you just don't have available to spend on such things, especially ones that already work well enough.
### 2. How do these pros and cons apply to refactoring the original VBA script?
-The refactoring has more steps broken down more clearly and would be a more useful teaching tool than even the more basic original script. However, it also complicated things enough to make a process that was very similar take almost twice as long to reimplement correctly.
