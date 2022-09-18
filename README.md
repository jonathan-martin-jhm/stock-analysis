# VBA_Challenge

## Overview of Project

### Background

Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job

### Purpose

The purpose of this project was to refactor VBA code to run faster than it did in the stock analysis performed in Module 2. The goal was to refactor code from Module2_VBA_Script to loop through the data one time and collect all of the information.

## Results

### Refactoring VBA code

#### Step 1

a) Create a tickerIndex variable and set it equal to zero before iterating over all the rows.

![VBA_challenge_1a](https://user-images.githubusercontent.com/111193280/190913084-25871883-8099-4b58-b266-290c7f738369.png)

b) Create three output arrays: tickerVolumes, tickerStartingPrices and tickerEndingPrices.

![VBA_Challenge_1b](https://user-images.githubusercontent.com/111193280/190913163-537974b4-393c-4c87-9869-0f9096f1d9e2.png)

#### Step 2

a) Create a for loop to initialize the tickerVolumes to zero.

![VBA_challenge_2a](https://user-images.githubusercontent.com/111193280/190913244-b967642b-08b1-418a-a831-035b1d6582d2.png)

b) Loop over all the rows in the spreadsheet.

![VBA_challenge_2b](https://user-images.githubusercontent.com/111193280/190913330-1382bec4-cd96-41db-bd66-df9bcb64139b.png)

#### Step 3

a) Increase volume for current ticker.
b) Check if the current row is the first row with the selected tickerIndex.
c) Check if the current row is the last row with the selected ticker.
d) Increase the tickerIndex.

![VBA_challenge_3](https://user-images.githubusercontent.com/111193280/190913467-8a3ae049-8e2b-474d-961e-329ddc6f4c53.png)

#### Step 4

Loop through the arrays to output the Ticker, Total Daily Volume and Return.

![VBA_challenge_4](https://user-images.githubusercontent.com/111193280/190913570-3a90b5d5-39ac-48e3-8cf2-b28166aab7e9.png)

#### Results for 2017

![VBA_Challenge_2017](https://user-images.githubusercontent.com/111193280/190913704-fc51796f-1113-49d2-9a1f-81900a36ef82.png)

#### Results for 2018

![VBA_Challenge_2018](https://user-images.githubusercontent.com/111193280/190913742-a171bd21-3346-4be0-927e-8c8e72fa8b02.png)

## Summary

### Overview of Advantages and Disadvantages to Refactoring Code 

#### Advantages
- Improves design of the software and efficiency of the code.
- Improves the legibility and makes it easier to understand.
- Indentifies bugs and inefficiencies within the code.
- Improves the ease and speed of programming in the future.
- Improves adaptability of the code.

#### Disadvantages

- Refactoring code can be time consuming and costly. The person or persons refactoring code are often not the original designer, so there can be a learning curve to identify the original intentions or design.
- Refactoring code runs the risk of introducing new bugs that were avoided in the original design.
- Refactoring does not change the end result of the code, so the cost and time spent to improve the code may not be seen as beneficial to customers, clients or supervisors.

#### Advantages and Disadvantages of the Original and Refactored VBA Script

The original and refactored script were neatly written, which facilitates with readability and helps determine the logic of the design. The disadvantage of the orignal code compared to the refactored script was the processing speed to obtain the variables used to analyze the 12 stocks. The disadvatange of the refactored code was the time and effort spent to improve the efficiency and adaptibility of the code while achieving the same end goal of the data analysis.







