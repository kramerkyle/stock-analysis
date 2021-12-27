# Stock Analysis

## Overview of Project
This repository contains an excel file and png files that pertain to a stock analysis that was executed in Excel using Visual Basic Application. The user may run an analysis on the twelve stocks that are included in the file by clicking the "Run Analysis" button and entering the year in question into the Message Box that appears.

## Results
### Process
The analysis hinges on a couple of metrics: volume and return. The portfolio manager believes that a high value in the former may indicate a fairer price due to increased discovery. The return shows how much money would have been made over the course of the year.

At it's simplest, the code sums the volume of shares traded, then pulls the starting and ending price, and finally calculates the yield. The outputs are placed into the "All Stock Analysis" sheet.

The analysis as it stands, includes two modules of code within the VBA Developer view. Module 1 contains subs that were written while walking through course material. Module 2 contains refactored code that is more efficient. The refactored macro has been assigned to the "Run Analysis" button on the All Stock Analysis sheet.

### Outcomes
Within 2018, ENPH traded up 81.9% on the year on over 600,000,000 in volume. RUN was the leader of the year with a yield of 84% and approximately 500,000,000 in volume. All other stocks lost value.

In 2017, the only stock to lose value was TERP. There was a wide variety in upsides from RUN's 5% to DQ's 199%.

The original time to execute code for 2017 and 2018 was 0.2617188 and 0.2539062 seconds, respectively. Through refactoring code, the time has been reduced significantly. Please see the two figures below for details.

![VBA_Challenge_2017](https://github.com/kramerkyle/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA_Challenge_2018](https://github.com/kramerkyle/stock-analysis/blob/main/VBA_Challenge_2018.png)

## Summary
### Pros & Cons of Refactoring Code
Refactoring code is an excellent way to optimize code to run faster and more reliably. Not only that, refactoring provides a great opportunity to ensure effective documentation has occurred and that the code is clean and understandable. While these are clear benefits, refactoring code is not without a con or two.

A potential con of refactoring code is that the time invested may not be earned back on the running of the code. This project was a prime example of this, as the original code ran well under a second. A second prospective issue is that the person attempting the refactoring may not be successful in improving the code. This would burn valuable time and resources on an unsuccessful venture.

### Original vs. Refactored
The initial code used "for" loops to track data based on the stock ticker, looping through every row of the sheet once per ticker. By refactoring the original VBA script, users have greater speed, reliability, and clarity of code. However, the time saved by the refactoring is negligible compared to the time spent refactoring. The second potential negative was not realized.
