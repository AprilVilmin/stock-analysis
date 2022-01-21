# Stocks with VBA (An Analysis of Green Energy Stocks Using VBA)

## Overview of Project
### Background
Steve is a new finance graduate whose parents have hired him to review DAQO New Energy Corp's stock. DAQO New Energy Corp's stock is a green energy company. Steve’s parents believe that green energy is going to be more relevant in the future. Steve wants to look at diversifying his parents’ money, as he has concerns about investing it all in one place. Because of this Steve decided to analyze twelve green energy stocks including DAQO.

### Process
Steve created an Excel Workbook containing the stock data for the twelve green energy companies, including DAQ0 New Energy Corp. Steve has requested that we, the Analytics Department, help with this request. The Analytics Department decided that since this information is already in Excel, we will use Visual Basic for Applications (VBA) to analyze the data. Once the Analytics Department decided upon VBA, a series of steps took place:
1. Show Steve how to enable Macros on his computer
2. Write VBA Code for DAQO New Energy Corp stocks so Steve could get that information to his parents (this included a hardcoded year and stock)
3. Write a new Macro/new Code so that it will loop through all of the stocks (lets you loop through all stocks not just DAQO or a different hard coded stock)
4. Write a new Macro/Iterate the VBA code so that it formats the data (including bolding, edges, number format, autofit and interior color)
5. Write a Macro to clear the worksheet (this lets you start with a clean slate)
6. Add button(s) to DQ Analysis and All Stocks Analysis Worksheets (This allows you to clear and run the data directly from the sheet rather than through the developer tab)
8. Write a new macro/Iterate on the VBA code so that any year can be chosen (an input box is added to the code replacing the hardcoded year values)
9. Add a timer to the code to measure performance
10. Refactor the VBA Code to try to get it to run faster

## Results
The Total Daily Volume and Ret
urn results were the same before and after the code was refactored (screenshots provided below). Which is great as the results are not supposed to change due to the data being refactored. One difference to note is that this macro has different formatting on the Total Daily Volume and Return columns than some of the other macros due to Skill Drill within the module. The run times decreased, which was the goal of this refactoring. The details of the decrease are outlined in more detail in the 'Advantages' section in the Summary. 

### Results Prior to Refactoring

#### 2017 

![2017 Before Refactoring.png](https://github.com/AprilVilmin/stock-analysis/blob/main/2017%20Before%20Refactoring.png)

#### 2018
![2018 Before Refactoring.png](https://github.com/AprilVilmin/stock-analysis/blob/main/2018%20Before%20Refactoring.png)

### Results Post Refactoring

#### 2017 
![2017 After Refactoring.png](https://github.com/AprilVilmin/stock-analysis/blob/main/2017%20After%20Refactoring.png)

#### 2018
![2018 After Refactoring.png](https://github.com/AprilVilmin/stock-analysis/blob/main/2018%20After%20Refactoring.png)

### Runtimes Post Refactoring

![VBA_Challenge_2017.png](https://github.com/AprilVilmin/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/AprilVilmin/stock-analysis/blob/main/VBA_Challenge_2018.png)

### Greatest Impact
I think that the removal of the nested loop is one of the things that had the greatest impact on the speed in the refactored code. I have included a snippet of the code that was removed below:

       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If

## Summary
### What are the advantages or disadvantages of refactoring code?

#### Advantages
1. Increased System Performance (Runtime)
2. More Efficient Code
3. Adding Comments/Notes to the Code

#### Disadvantages
1. Time Consuming
2. New Errors or Vulnerabilities Could Be Added to the Code in Error/By Mistake

### How do these pros and cons apply to refactoring the original VBA script?
Yes, several of these apply. See the details below:

1. The code is less complicated/more efficient in the fact that there are not any nested loops. See code snippet below.
2. Refactoring this code took me many hours.
3. The run times decreased. See the screenshots below.

#### Piece of Removed Code (Same as Snippet Above)

      Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If

#### Prior to Refactoring 

![2017 Prior to Refactoring.png](https://github.com/AprilVilmin/stock-analysis/blob/main/2017%20Prior%20to%20Refactoring.png)

![2018 Prior to Refactoring.png](https://github.com/AprilVilmin/stock-analysis/blob/main/2018%20Prior%20to%20Refactoring.png)

#### After Refactoring

![VBA_Challenge_2017.png](https://github.com/AprilVilmin/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/AprilVilmin/stock-analysis/blob/main/VBA_Challenge_2018.png)
