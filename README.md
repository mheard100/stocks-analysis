# STOCK ANALYSIS WITH VBA + EXCEL

## OVERVIEW: VBA Stock Analysis Project

### Purpose

In this analysis, the Stock Market Dataset is refactoed with VBA code to loop through all data to collect information. Then, the analysis is checked to see if the VBA script ran faster than the module 2 code. 

### Analysis and Challenges
A quick look:

- Prepare 'VBA_Challenge.vbs' file for the project.
- Create resources folder in **GitHub** to hold the run-time pop-up messages shown with screenshots after running refactored analyses for 2017 and 2018.
- Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
- Use the steps **Refactor VBA code and measure performance** to add needed code into the starter code file.

> Use VBA and the starter code provided in this Project to refactor the VBA Script dataset to loop through the data one time and collect all of the information.

#### Our Challenge Data Background
> Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

> In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

> Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

## RESULTS: Refactor VBA Code and Measure Performance
 
### Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:

**1.Create a tickerIndex variable and set it equal to zero before iterating over all the rows.


**2a. Arrays are created for 'tickers', 'tickerVolumes', 'tickerStartingPrices', and 'tickerEndingPrices'.**
>The tickerVolumes array is a long data type while tickerStartingPrices and tickerEndingPrices are single data types

**2b. A for loop that will loop over all the rows in the spreadsheet.

**3. Inside the for loop in Step 2b, a script that increases the current 'tickerVolumes' (stock ticker volume) variable and adds the ticker volume for the current stock ticker was written.
Then, an if-then statement to check if the current row is the first row with the selected 'tickerIndex' was written. If it was, then the current starting price to the 'tickerStartingPrices' variable was assigned.
Following, an if-then statement to check if the current row is the last row with the selected tickerIndex was written. 
If it is, the current closing price was assigned to the 'tickerEndingPrices' variable.
Finally, a script that increases the 'tickerIndex' if the next row’s ticker doesn’t match the previous row’s ticker was added. 


**4.  A for loop to loop through the arrays ('tickers', 'tickerVolumes', 'tickerStartingPrices', and 'tickerEndingPrices') to output the Ticker, Total Daily Volume, and Return columns 
was added to the script. 






***Final VBA Analysis 2017***

![name-of-you-image](https://github.com/mheard100/stocks-analysis/blob/main/VBA_Challenge_2017.png)

***Final VBA Analysis 2018***


![name-of-you-image](https://github.com/mheard100/stocks-analysis/blob/main/VBA_Challenge_2018.png)



## SUMMARY: Our Statement:

### Deliverable with detail analysis:
**1. What are the advantages or disadvantages of refactoring code?**

**Disadvantages:**

> - Can have duplication of code if care not taken
> - Easy to run into errors since code may not be written by same person or at the same time. 
> - Refactoring process can affect the testing outcomes. 


**Advantages:**
> - The code is similar to originally written code.  
> - Increases efficiency of code.
> - Shows patterns.

**2. How do these pros and cons apply to refactoring the original VBA script?**

> Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. In all, refractoring
is successul if it can be done easily over time and by different people. It should be easy to maintain and increase efficiency of the code. 





