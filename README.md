# Stock Analysis: VBA Challenge
## Purpose
In this challenge, I’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that I did in this module. Then, I’ll determine whether refactoring my code successfully made the VBA script run faster. Finally, I’ll present a written analysis that explains my findings.

# Deliverable 1: Refactor VBA Code and Measure Performance
## Step 1a:
* Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.

![image](https://user-images.githubusercontent.com/87340105/155652160-306ab02e-cd24-44aa-974d-8daf8ef04db2.png)


## Step 1b:
* Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
* The tickerVolumes array should be a Long data type.
* The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

![image](https://user-images.githubusercontent.com/87340105/155652208-9fba41d0-d6dd-4a09-a59c-609944dfbabc.png)

## Step 2a:
Create a for loop to initialize the tickerVolumes to zero.

![image](https://user-images.githubusercontent.com/87340105/155653682-ecd596c4-bc51-4880-bc72-fd2b3d72bd83.png)

## Step 2b:
Create a for loop that will loop over all the rows in the spreadsheet.

![image](https://user-images.githubusercontent.com/87340105/155653739-d35407eb-4a5f-4d40-ba14-58203fa38c8d.png)

## Step 3a:
Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
Use the tickerIndex variable as the index.

![image](https://user-images.githubusercontent.com/87340105/155653839-968bc6d9-dc78-41c9-94d7-33f8b6b4d5e1.png)

## Step 3b:
Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.

![image](https://user-images.githubusercontent.com/87340105/155653895-9c100b32-0ad4-4bfa-8278-905b17f0e6f1.png)

## Step 3c:
Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.

![image](https://user-images.githubusercontent.com/87340105/155653957-ec585f5e-ff09-4382-a8eb-35f963b3939a.png)

## Step 3d:
Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.

![image](https://user-images.githubusercontent.com/87340105/155654770-12803101-e23a-488c-bbbe-e064955b7f42.png)

## Step 4:
Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.

![image](https://user-images.githubusercontent.com/87340105/155655024-219bfcb1-4833-4d9e-9d96-7cefda651783.png)

## Analysis
## Summary
