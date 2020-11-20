# Module_2_VBA_Challenge - Green Stocks Analysis Using VBA

## Overview of Project
The aim of this project was to help Steve identify green stock alternatives for his parents.  Steve was interested in understanding the performance of DQ stock (which his parents had expressed interest in).  In addition there were other green stocks with potential that we wanted to assess.  Rather than calculate the values manually, Steve was keen for us to write code in VBA that could be used repeatedly to analyze performance for various years.   

### Purpose
The purpose of writing this code was to automate the analysis of green stock performance for different years (2017 & 2018).  The final refactored code removed extranous components, clearly detailed the funtion of each element in the code and enhanced the performance of the code allowing it to run with even greater speed/ efficiencey.  

## Analysis and Challenges
There were 3 key areas of focus:  
1) Analysis of Green Stocks performance for 2017 and 2018
2) Automation of analysis through code in VBA  
3) Refactor the code to ensure its efficiency 


##  Results 

### Analysis of Green Stocks Performance

 Initially we looked at Steve's parents' chosen stock DQ.  The analysis showed DQ had a negative return in 2018 of -62.6%.  DQ had performed well in 2017 with a return of 194.7%.  It definitely performed poorly in 2018 through.  
 
 We then looked across a series of other green stocks.  Those with positive returns in 2018 included ENPH with a positive return of 81.9% and RUN with a positive 2018 return of 84.0%.  All others had negative returns in 2018. 

 Interestingly enough, all of the stocks analyzed, with the exception of TERP, showed positive returns in 2017.  It is evident that something negatively inpacted green stocks overall across our selection in 2018. Given we are only looking at 2 years' worth of data it would be interesting to analyze additional previous years to understand the actual trends before providing Steve and his parents with a game plan for investment in the coming year.  It would be helpful to understand net gains and losses over a longer period.
 
 ![2018 Green Stock Analysis](https://github.com/PatriciaCB1/Module_2_VBA_Challenge/blob/main/2018%20Stocks%20Analysis.png)
 ![2017 Green Stock Analysis](https://github.com/PatriciaCB1/Module_2_VBA_Challenge/blob/main/2017%20Stocks%20Analysis.png)
 
### Automation of Analysis
 By writing code and creating macros we have automated Steve's analysis of green stocks.  By using variables, arrays, for loops and conditionals we have written code that enables Steve run analysis incredibly quickly.  We tried, where possible, not to hard code in variables.  Creating dynamic variables that can adjust as new data points or even worksheets are added (for additional years) has been our focus.

### Refactoring of code
Advantages and Disadvantages of refactoring code.  How these pros and cons apply to refactoring the original VBA script.
In order to be successful, code must work - peforming its intended tasks and providing the needed output/ outcome.  Code performance can also be a key component of design.  Some of the aims of refactoring code include tidying up extraneous lines in the code, duplication, (code smell!) making code cleaner and easier to understand, enhancing performance and reducing run time etc..  The advantages of refactoring code are that it makes the code more efficient (runs faster), uses less memory, makes the code easier to read, allows you and others to comprehend and update more easily, creates a better flow/ more logical order and also allows you to optimize code (i.e.making code dynamic instead of using hard code) creating greater versatility and future-proofing your efforts.  One of the main disadvantages of refactoring code are that it adds time to the process.  You also run the risk of making errors and creating bugs when you start to tinker with a working code.  

When we were initially writing our code for Steve we wrote it piece by piece, function by function.  We focussed on ensuring each individual task was performed correctly.  Once we had the entire code written we were able to see the bigger picture.  Rather than looking at the code as individual functions we were able to work our way through to arrange the code in more logical order, remove redundancies/ duplications and create further efficiences.  We utilized indecies rather than hard code, combined some of our loops, removed extranous page activiation, moved some of the initiation orders etc.  This enabled us to improve our run time significantly.  Run time was reduced from 0.515625 seconds to 0.078125 seconds.  This will allow for much faster analysis for larger datasets moving forward.  During refactoring I encountered a situation where I accidentally introduced a bug into my code when it had been working fine previously.  It took me 2 hours to remedy the situation.  Fortunately I had time to diagnose and correct this but had I been up against a tight deadline I would have had to skip the refactoring process or submit at a later time.

## Summary
In summary through this process we were able to support Steve by creating code allowing him to quickly and easily run analysis on key green stocks by year through the use of a button/ input box.  Since we took the time to refactor the code it now runs much more quickly and efficiently, meaning Steve can analyze very large data sets with ease.  We also provided him with a clean, easy to read code that will allow him (or us) to more readily update the code in future should his needs change. 
