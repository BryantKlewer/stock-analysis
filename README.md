# All Stocks Analysis

## Project Overview
- This analysis is intended to determine if refactored VBA script will collect and output formatted stock data in an Excel worksheet more efficiently than a base VBA script. 

### Purpose
- The purpose of this analysis is to help determine whether using a base VBA script is optimal or if refactoring the script will allow us to obtain the same data results and output them more quickly. For the purposes of this analysis, we will be alalyzing twelve different companies' stock market data from 2017 and 2018. These companies are all from the Green engergy sector.

## Results  

### Code Performance
- In order to determine the efficiencey of both codes, we set up a timer on both the base script and refactored script to capture the start and end times. We then output this information in a message box once the script finished running. ![timer_code](https://github.com/BryantKlewer/stock-analysis/blob/main/Resources/timer_code.png) 
We then ran our base script to produce and output our results for both 2017 and 2018. The speed in which the codes ran were 0.609375 seconds for 2017 and 0.625 seconds for 2018. 
![2017_original_0609375](https://github.com/BryantKlewer/stock-analysis/blob/main/Resources/2017_original_0609375.png) 
![2018_original_0625](https://github.com/BryantKlewer/stock-analysis/blob/main/Resources/2018_original_0625.png) 
After reviewing the entire original scirpt, I determined that it was fairly efficient, which is sometimes the case. I did decide to move some of the cell range formatting data out of a loop and place it in the output sheet formatting section. ___format cell range___  After making this change and a few typographical changes, the refactored scipts were run on both 2017 and 2018's data. In both cases, we were able to successfully duplicate the same results and output formatting. Both speeds for the refactored scripts ran slightly faster than their originals as well. ___refactor png's___ 

### Stock Performance
- What we can see in the 2017 stock analysis is that 11 of the 12 total stocks reviewed posted gains for the year. These gains ranged from 5.5% to 199.4%. Only TERP realized a loss of 7.2% for 2017. ___2017___ On the other hand, only 1/6th of stocks posted gains in 2018, they were ENPH (81.9%) and RUN (84.0%). Losses in 2018 amongst the remaining 10 stocks ranged from 3.5% to 62.6%. ___2018___ We are able to determine from this data that the Green energy sector can be volitile with large gains and losses. 

## Summary

- What are the advantages or disadvantages of refactoring code?
Generally speaking, the main advantages of refactoring code is that it becomes simplified, easier to read, and faster. It is able to accomplish this, all while still retaining its base functions. The primary disadvantage of refactoring code is that you can potentially introduce errors into the script and loose the functionality of the code. 

- How do these pros and cons apply to refactoring the original VBA script?
I think that what we can determine in the case of this particular analysis is that refactoring my own, personally created VBA script, made little difference. I would consider this a positive, it indicates that it was generally well written to begin with. I think another pro is that we were still able to increase efficiently slightly showing that code can always be written better and more efficiently. The only con for this particular analysis is ROI, return on investment. The time taken to analyze the code and make slight modifications only resulted in minimal speed changes.
