# **VBA Stock Analysis**

### Overview
Steve is in the process of helping his parents invest in the stocks of companies that focus on green energy. He has reached out for help analyzing the stocks of a handful of other green energy companies so that his parents can divefsify their portfolio. Steve has provided an Excel file with the information for the green energy companies stocks from 2017 and 2018 he would like analyzed. The goal in helping Steve is to create a macro in Visual Basic for Applications (VBA) to automate this analysis. The initial macro that was built worked well for the 12 green energy stocks that Steve was originally looking in to. He has since asked if the macro can be refactored to analyze the entire stock market as a whole. Refacctoring of a macro or a code in general, is taking the original code and editing it to fit a new need instead of creating a new marco or code from scratch. 

### Analysis
The orginal macro that was created analyzed 12 green energy stock options to help Steve diversify his parents stock options. The refactored code now allows Steve to run analysis over the entire stock market for a given year and get the Ticker Name, the Total Daily Volumne, and the Percent Return for each. The refactored code also formats the columns, adding commas to the Total Daily Volume column and converting the values in the Return column to percentages. The code also turns the cell in the Return column green if the return is positive and red if the return is negative. Before, the formatting and color coding was all accomplished in seperate macros from the original analysis. 

When refactoring the VBA code, most of the code was able to stay the same. The image below highlights the updates and changes that were made to make the code run the way that Steve has requested. 
![alt text](https://github.com/John-Thomas-872/stock_analysis/blob/main/Resources/VBA_Challenge_Refactored.png)

### Results
There is another benefit to refactoring the VBA code other than allowing Steve to analyze more data. The refactored code now runs singinficantly faster. The original 2017 analysis ran in 0.6992188 seconds while the 2018 analysis ran in 0.6914063 seconds. The timer feature was added to determine if refactoring the code saved time.

The 2017 analysis performed with the refactored code ran in 0.1328125 seconds, as seen in the image below. The refactored code perfomed the 2017 analysis roughly 5 times faster than the orignal code.

![alt text](https://github.com/John-Thomas-872/stock_analysis/blob/main/Resources/VBA_Challenge_2017.png)

The 2018 analysis performed with the refactored code ran in 0.09375 seconds, as seen in the image below. The refactored code performed the 2018 analysis more than 7 times faster than the original code.

![alt text](https://github.com/John-Thomas-872/stock_analysis/blob/main/Resources/VBA_Challenge_2018.png)

Without using the timing feature in VBA it would be nearly impossible to tell the difference in the efficiency between the refactored and the original code, especially since both codes ran in less than a second. 

### Summary
For this project it made the most sense to refactor the original code instead of starting from scratch. In doing this all of the original code was kept and If statements and for loops were added to analyze the entire stock market. 

**Pros**

One of the pros of refactoring the existing code it organization. The code and comments for the code are already established so it makes it easier to go back and add the necessary changes in the correct locations. As long as a backup of the original code is saved, there is little to no risk in refactoring code. If something is deleted or the code developes a bug or error, the original code is there to fall back on.

**Cons**

Refactored code can often become long and complicated. Because of the way the original code was created, it might be harder to edit or make the necessary changes in a simple manner. Refactoring code can also be very time consuming and not worth the time and money to start over.  

The original code was very good at doing exactly what Steve asked. To analyze 12 stocks for green energry companies. It is simple and works as asked. 
The refactored code on the other hand is a little bit more complicated, but opens up the options that Steve has in the future. The refactored code also takes in the formatting macro that adds the correct formatting for the numbers and adds color for the positive and negative percentages. 

Finally, to make the analysis easier for Steve, two buttons were added to All Stocks Analysis worksheet. The first button runs the macro after asking the user to input the year they would like to analyze. The second button clears the worksheet so that the first macro can be run again. 

