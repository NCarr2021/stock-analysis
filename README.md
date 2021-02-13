# stock-analysis

# Module 2 Challenge

## Overview of Project

Steve graduated with a Finance Degree and his parents will be his first clients. His parents are passionate about green energy and decided to invest
all their money into a green energy stock called DAQO New Energy Corp. Steve prefers his parents diversify and will need to run analysis to check 
DAQO’s performance and check the performance of other green energy stocks. Steve created a spreadsheet containing all the stock data that will be used
for the project analysis. The data is provided in Excel format and VBA will be used to automate formulas for calculationa to present a summary by stock tickers.

## Results

Stock data is provided in Excel with one format and 2017 and 2018 in different worksheets.
This data was used to create an All Stocks Analysis summary worksheet.

<img src="/Stocks-2017.png" width="600" />

<img src="/Stocks-2018.png" width="600" />

Analysis was first done selecting only DAQO (ticker DQ) for 2018 to determine stock performance.
The code only selected ticker 'DQ'.

<img src="/DQ-Only-Selection.png" width="400" />
<img src="/DQ-Only.png" width="600" />

Results show the return for 2018 was poor, Column C, (highlighted in yellow) with run time.

<img src="/DQ-Analysis.png" width="400" />

Adding start / end timer shows elapsed time to run code.
Except for the change in year, the same code was used to run analysis on 2017.

	"" 'Access stock data ""
	"" Worksheets("2017").Activate ""
 
Adding the start / end timer shows the elapsed time for running the code.
<img src="/TimerStart.png" width="400" />
<img src="/TimerEnd.png" width="300" />

<img src="/DQ-Analysis-2017timer.png" width="400" />

<img src="/DQ-Analysis-2018timer.png" width="400" />


Code changes added another option to select different worksheet data.
Use an InputBox.

	"" Value = InputBOx("What year would you like to run the analysis on?")
	
Or use a Button linked to the specific macro.

<img src="/ButtonSelection.png" width="400" />
	
	
Comparing both years, the DQ stock had a small return in 2017 vs. a loss in 2018.
	
<img src="/DQ-Analysis.png" width="400" />
					
<img src="/DQ-Analysis-2017.png" width="400" />
		
		
Further analysis was performed on all other stocks for 2018 and shows performance for the majority is poor.
Only two stocks ENPH and RUN did well. 

<img src="/All-Stocks-Analysis.png" width="400" />

However, stocks in 2017 did much better.

<img src="/All-Stocks-Analysis-2017.png" width="400" />


## Summary

What are the advantages or disadvantages of refactoring code?

Disadvantages include potential for causing errors if existing code isn’t fully understood.Also, the refactored code may 
not perform as well as existing and be harder to understand and maintain. Refactoring may take additional time not currently available.

Advantages include making it easier to read by adding whitespace and comments making it simpler to maintain. Performance may also improve 
if code is simplified with better structure and using indexes/arrays.

In this project, refactoring was positive. The additional code for indexes and variables added 
an improved approach to summarize all tickers. Including comments throughout explained function's purpose. Coding was added to 
format the output to easily identify gains vs. losses. 


