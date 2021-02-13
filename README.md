# stock-analysis

# Module 2 Challenge

## Overview of Project

Steve graduated with a Finance Degree and his parents will be his first clients. His parents are passionate about green energy and decided to invest
all their money into a green energy stock called DAQO New Energy Corp. Steve prefers his parents diversify and will need to run analysis to check 
DAQO’s performance first then check other green energy stocks performance. Steve created all the stock data that will be used for the project analysis.
The data is provided in Excel format and VBA will be used to automate formulas for calculation to present a summary.

## Results

Stock data is provided in Excel with each year in different worksheets but in the same format.

	![2017 Stock Data](NCarr2021/stock-analysis/Stocks-2017.png)

	![2018 Stock Data](Resources/Stocks-2018.png)

Analysis was done only selecting DAQO (ticker DQ) for 2018 to determine stock performance.

	![DQ](Resources/DQ-Only-Selection.png)

Code only selects ticker 'DQ'.

	![DQOnly](Resources/DQ-Only.png)

Results show the return for 2018 was poor, Column C, (highlighted in yellow) with run time.

	![DQ 2018 Results](Resources/DQ-Analysis.png)

Adding start / end timer shows elapsed time to run code for 2018:

	![DQ 2018 Timer](Resources/DQ-Analysis-2018timer.png)

Except for the change in year, the same code was used to run analysis on 2017.

"" 'Access stock data ""
"" Worksheets("2017").Activate ""
 
Adding the start / end timer shows the elapsed time for running the code.

	![DQ 2017 Timer](Resources/DQ-Analysis-2017timer.png)
	
	![DQ 2018 Timer](Resources/DQ-Analysis-2018timer.png)
	
Another option is select different worksheet data is to use an InputBox.

	"" Value = InputBOx("What year would you like to run the analysis on?")
	
Or Or a Button linked to the specific macro.

	![Button](Resources/ButtonSelection.png)
	
Comparing both years, the stock had a small return in 2017 vs. a loss in 2018.
	
		![Outputs](Resources/DQ-Analysis.png)
		![Outputs](Resources/DQ-Analysis-2017.png)
		
Further analysis was performed on all other stocks for 2018 and shows performance for the majority is poor.
Only two stocks ENPH and RUN did very well. 

	![All Stocks](Resources/All-Stocks-Analysis.png)

However stocks in 2017 did well.

	![All Stocks}(Resources/All-Stocks-Analysis-2017.png)

## Summary

1.	The disadvantages of refactoring code include potential for causing errors if existing code isn’t fully understood.
Also, the refactored code may not perform as well as existing, and be harder to understand and maintain. 
The advantages of refactoring code include making it easier to read by adding whitespace and comments making 
is simpler to maintain. Performance may also improve if code is simplified.

2.	In this project, refactoring was positive overall. The additional code for indexes and variable added 
an improved approach to summarize all tickers. Coding was added to format output to easily identify gains vs. losses. 

