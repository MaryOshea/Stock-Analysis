# Stock Analysis 


## OVERVIEW: VBA Stock Analysis 

### PURPOSE
The purpose of conducting the refactored stock market analysis was to enhance our previously created VBA code in a way that allows our VBA script to loop through *all* of the data within our workbook with the click of a button.  The ultimate goal that we are trying to achieve by making these adjustments is for our VBA script to run faster and more efficiently.  Lastly, we will be able to determine the outcome of our refactored analysis by comparing the difference in the time it took to run the VBA script pre-refactored updates vs. post-refactored updates. 

### ANALYSIS & CHALLENGES
A high level overview of the actions performed while conducting this refactored stock market analysis include:
- Creating a Ticker Index variable and setting it equal to zero; then using this variable to access the correct index across our ticker array; & tickerVolumes, tickerStartingPrice, and tickerEndingPrice arrays respectively.
- Creating a **loop** to initialize the tickerVolumes to zero while also identifying if the next row's ticker matches the previous rows and increasing the tickerIndex volume respetively if the ticker does not match.
- Created a loop that will include all rows of data in the spreadsheet
- Used an If Then statement to identify weather or not the current row is the first row with the selected ticker Index; and if it is then assigning the cell with the current tickers respective starting/ending price.
- Used conditional formatting within our VBA script in a way that would provide our audience with an easily digestable visual depiction of each stocks yearly performance 
- Added comments as needed to explain the steps taken in developing our code
- Ran the final stock analysis in order to determine if the outputs in our historical analysis match the outputs of our refactored analysis


With that being said, the outputs of my 2017 and 2018 refactored stock analysis **successfully match** the outputs from my previous All Stock Analysis performed at the start of this module.  Please refer to the images below to see the VBA Challenge Analysis & All Stocks Analaysis output images for the 2017 & 2018 to confirm and support my statement findings:

***Final All Stocks Analysis Outputs for 2017 - 2018***

<img width="578" alt="2017 Outcomes" src="https://user-images.githubusercontent.com/77044730/110081755-67fb8e80-7d5a-11eb-8444-68446a64beaf.PNG">

<img width="575" alt="2018 Outcomes" src="https://user-images.githubusercontent.com/77044730/110081821-7fd31280-7d5a-11eb-8f6a-06b607869804.PNG">


***VBA Challenge Analysis Outputs & Elapsed Time for 2017 & 2018***

<img width="608" alt="2017 MessageBox with Outcomes" src="https://user-images.githubusercontent.com/77044730/110081848-8d889800-7d5a-11eb-96dd-98041b951b8c.PNG">

<img width="603" alt="2018 Messagebox with Outcomes" src="https://user-images.githubusercontent.com/77044730/110081894-9d07e100-7d5a-11eb-8ab7-21853ea00919.PNG">



## SUMMARY

**1. What are the advantages or disadvantages of refactoring code?**

- **Advantages:** Logical errors that could be more efficiently structured using excel logic are easily identifiable.  In addition, once identified these errors can be easily fixed using conditional formatting and nested loops to improve the efficiency of how the code runs, as demonstrated within this project.
- **Disadvantages:** It is easy for restructured code to quickly become unstructured and messy.  When this occurs it is best to split the code into different functions and to always remember to use comments and white spaces whenever possible to keep the code clean and easy to follow.  In addition, refactoring also may affect the testing outcome.


**2. How do these pros and cons apply to refactoring the original VBA script?**

> As stated above, when refactoring our origional VBA script it is important to always keep code organization at the forefront of our minds.  It is common and often very likely that our scripts will need to be frequently updated based upon the demands of our stakeholders - in which case, we want our script to be as simple and easy to understand as possible in order to ensure that our future refactoring and/or troubleshooting endevours can be made with ease.



