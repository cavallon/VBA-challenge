# VBA-challenge
Challenge 2 Files

READ ME - Multi_Year_Stock_Data Macro File

1. There were multiple subs created to run the code that computes the results needed. In order for the code to operate correctly,
it must always be run in the order shown in the Main_Macro Sub. If you attempt to run the entire code while in the "View Code" window
in VBA, make sure your cursor is in the Main_Macro Sub. This is first sub all the way at the top of the page. 

2. So that you understand each Sub, I will briefly explain it here:

Sub "Main_Macro" runs these subs in the following order: tickersymbols,
color_code, max_value, and ticker_max. 

Sub "tickersymbols" loops through all of the ticker symbols in Column A. Whenever it finds that a new 
ticker symbol has started, it will populate the required data associated with the ticker symbol. The required data is one ticker symbol from each
group, the yearly change in the stock price (positive or negative) as calculated by subtracting the closing price at the end of the year to the
opening price at the begining of the year, the total percentage of change between the opening price at the begining of the year and closing price
at the end of the year, and the total volume of each stock.    

Sub "color_code" changes the color of Column J to either Red, Green, or No color. It becomes Red if there is a negative value, it becomes Green if
there is a positive value, and it remains with no color if there is a 0 or the cell is blank. 

Sub "max_value" finds the greatest % increase, greatest % decrease from Column K, and greatest total stock volume from Column L. It then populates the 
values in Column Q. 

Sub "ticker_max" finds the ticker symbol that correlates to each of the values in Column Q, and populates Column P with it. 

3. There is a separate sub that is not included in the "Main_Macro" sub. This sub is titled Reset_Workbook and is the code for the reset button that is 
featured in the main excel workbook. Pressing this button will erase all values that have been calculated by the "Main_Macro" sub. 


----IMPORTANT----

4. Lastly, if you have access to my main excel workbook, it will not be necessary to open the "View Code" window in VBA when using this macro. The main excel workbook has both a Calculate button and a Reset button. All you need to do is press the calculate button to calculate everything, and then press the reset button to remove all of the calculations and formatting that was just completed. 
