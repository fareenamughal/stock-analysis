# All Stocks Analysis Refactored


# Refactoring VBA Code and measuring performance
---
## Overview of Project
---
The overview of the project is determine whether the VBA script will run faster after being refactored or edited. 

https://github.com/fareenamughal/stock-analysis/blob/main/Module2Challenge/VBA_Challenge.xlsm   
{Link to the xlsm file containing the VBA code and workings for this challenge.}

### Purpose
---
The purpose of the project is to measure how fast the VBA script runs after being edited or refactored.
An analysis of stock for 2017 and 2018 was carried out to determine the performance of various stocks. We used excel and VBA scripting to provide the stock            performance analysis very quickly. We had previously already generated a VBA script and ran a macro named "All Stocks Analysis". This code was edited to determine      whether we could further improve the VBA code so that it would run faster than before. The edited code was named " All Stocks Analysis Refactored".

---
The two scripts are:
  - Original VBA script - All Stocks Analysis.
  - Refactored VBA script - All Stocks Analysis Refactored.
---
## Results
--- 
### Comparing stock performance results for 2017 & 2018 and All Stocks Analysis for 2018 vs All Stocks Analysis Refactored for 2018.
---
(https://github.com/fareenamughal/kickstarter-analysis/blob/main/Crowdfunding%20analysis/Resources/Theater_Outcomes_vs_Launch.png)
----
#### Stock performance comparison of 2017 - vs - 2018
  - The best performing stock for 2018 was ticker symbol RUN giving a 83% return followed by ENPH with an ~82% return/
  - The worst performing stock for 2018 was ticker symbol DQ followed by JKS.Both had indicated a decrease in returns of over 60%
  - The best performing stock for 2017 was ticker symbol DQ which almost doubled the returns at 199% followed by SEDG at 184.47%
  - The worst performing stock for 2017 was ticker symbol TERP which indicated a decrease in returns of approx.7%.
  - Overall 2017 was a better perfoming year compared to 2018. Though it would be advisable to investige the underlying cause of general market downturn in 2018           before any investment decision is made.	 
----
#### Original VBA script - All Stocks Analysis for 2018
  - The original VBA script for All Stocks Analysis for 2018 was executed in 0.9780273 seconds
  - Screenshot reflecting the execution time
----
#### Refactored VBA script - All Stocks Analysis Refactored for 2018
  - The refactored VBA script for All Stocks Analysis Refactored for 2018 was executed in 0.9440918 seconds 
  - Screenshot reflecting the execution time for the refactored script
  - This reflected a slight improvement in execution of the refactored VBA script for 2018
  - The refactored script was more organized and stated similar items under one area, for instance the header, arrary, index etc were all reflected in one area as much     as the execution allowed.
  - The comments were also better stated and provided a better description of the code following the comments

---
## Summary
---
### The advantages and disadvantages of refactoring
---
#### Advantages of refactoring
  - VBA scipting can be done quickly when not starting from scratch or when refactoring.
  - Refactored scripts are better organised as these are edited and improved versions of the original scripts.
  - Better organization helps us identify patterns alot more easily. 
---
#### Disadvantages of refactoring
  - VBA scipting can sometimes be very tedious and can have sometimes have undesirable effect on the outcomes.
  - There may be duplication which will force rethinking of the code.
---

### How do the pros and cons apply to refactoring the original VBA script
---
  - In this case the time spent on refactoring the scipt was alot more than the time saved. There was a marginal improvement in terms of speed.
  - Whilst trying to improve the code, by moving '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return. to after the loop for the tickers,     starting and ending price to the formatting section the code resulted in an additional row with the tickers(11) being duplicated. I am not sure why this kept           happening and had to change the code back to the original in this section.
  - I had to recheck to see if any areas were being repeated and this can also be time consuming
  - Overall the benefits of reusing the code can be enormous but this shoul be assessed based on the complexity of the code etc. 

