# Stocks with VBA (An Analysis of Green Engergy Stocks Using VBA)

## Overview of Project
### Background
Steve, a new finance graduate, parent's have hired him to review DAQO New Energy Corp's stock. DAQO New Energy Corp's stock is a green engery company and it is their belief that green energy is going to be more and more relevant in the future. Steve want to look at diverifying his parents money, as he has concerns about investing it all in one place. Becasue of this he wants to analyze several green energy companies. 

### Process
Steve created an Excel Workbook containg stock data for several green energy companies, including DAQ0 New Engery Corp. Steve has requested that we, the Analytics Department, help with this request. It has been decided that since this information is already in Excel we will use Visual Basic for Applications (VBA) to analyze the data. Once this was decided a series of steps took place:
1. Show Steve how to enable Macros on his computer
2. Write VBA Code for DAQO New Energy Corp stocks so Steve could could get that information to his parents (this included a hardcoded year and stock)
3. Write a new Macro/new Code so that it will loop through all of the stocks (lets you loop through all stocks not just DAQO or a different hard coded stock)
4. Write a new Macro/Iterate the VBA code so that it formats the data (including bolding, edges, number format, autofit and interior color)
5. Write a Macro to clear the worksheet (this lets you start with a clean slate)
6. Add button(s) to DQ Analysis and All Stocks Analysis Worksheets (This allows you to clear and run the data directly from the sheet ather than through the developer tab)
8. Write a new Macro/Iterate on the VBA code so that any year can be choosen (an input box is added to the code replaing the hardcoded year values)
9. Add a timer to the code to measure performance

## Results
![VBA_Challenge_2017.png](https://github.com/AprilVilmin/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/AprilVilmin/stock-analysis/blob/main/VBA_Challenge_2018.png)

## Summary
### What are the advantages or disadvantages of refactoring code?

#### Advantages
1. Increased System Performance (Runtime)
2. More Efficent Code
3. Adding Comments/Notes to the Code

#### Disadvantages
1. Time Consuming
2. New Errors or Vulunerabilities Could Be Added to the Code in Error/By Mistake

### How do these pros and cons apply to refactoring the original VBA script?
1. The code is less complicated in the fact that there are not any inner loops.
2. The run times decreased.

#### Prior to Refactoring 

![2017 Prior to Refactoring.png](https://github.com/AprilVilmin/stock-analysis/blob/main/2017%20Prior%20to%20Refactoring.png)

![2018 Prior to Refactoring.png](https://github.com/AprilVilmin/stock-analysis/blob/main/2018%20Prior%20to%20Refactoring.png)

#### After Refactoring

![VBA_Challenge_2017.png](https://github.com/AprilVilmin/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/AprilVilmin/stock-analysis/blob/main/VBA_Challenge_2018.png)
