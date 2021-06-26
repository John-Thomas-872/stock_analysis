# **VBA Stock Analysis**

### Overview
Steve is in the process of helping his parents invest in the stocks of companies that focus on green energy. He has reached out for help analyzing the stocks of a handful of other green energy companies so that his parents can divefsify their portfolio. Steve has provided an Excel file with the information for the green energy companies stocks from 2017 and 2018 he would like analyzed. The goal in helping Steve is to create a macro in Visual Basic for Applications (VBA) to automate this analysis. The initial macro that was built worked well for the 12 green energy stocks that Steve was originally looking in to. He has since asked if the macro can be refactored to analyze the entire stock market as a whole. Refacctoring of a macro or a code in general, is taking the original code and editing it to fit a new need instead of creating a new marco or code from scratch. 

### Analysis
The orginal macro that was created analyzed 12 green energy stock options to help Steve diversify his parents stock options. The refactored code now allows Steve to run analysis over the entire stock market for a given year and get the Ticker Name, the Total Daily Volumne, and the Percent Return for each. The refactored code also formats the columns, adding commas to the Total Daily Volume column and converting the values in the Return column to percentages. The code also turns the cell in the Return column green if the return is positive and red if the return is negative. Before, the formatting and color coding was all accomplished in seperate macros from the original analysis. 

When refactoring the VBA code, most of the code was able to stay the same. The image below highlights the updates and changes that were made to make the code run the way that Steve has requested. 
![alt text](../resources/VBA_Challenge_Refactored.png)
