# Stock Analysis
A multi year analysis of stock price performance using VBA macros to quickly and easily guage trends with a variety of tools.  The project then explains the process for refactoring the original VBA code to see if using the new VBA macros written to analyze 12 stocks is appropriate for  use with larger data sets such as the S&P 500 and Russel 2000.


## Overview
TheStock Anyslis is presented in two parts:  an initial survey of VBA macros to compare the perofrmance of a handful of stocks from 2017 and 2018, then refactoring the code used to make the macros run more efficiently to handle larger data sets.  Code refactoring is a process [defined as] (https://refactoring.com/)

> a disciplined technique for restructuring an existing body of code, altering its internal structure without changing its external behavior.

In this case to shorten the run time and therefore improve its performance.  

## Results

### Part I
The first project created an array of stock tickers shown here:

![array](/../Resources/Array of tickers.png)

For each element (specific stock symbol) VBA was used to run a ´for´loop inside a ´for´loop - called a **nested loop** demonstrated here

![nested loop](/../Resources/nested loop.png)

Notice how the ´For j = ´ starting the nest and ´next  j´ definition  is *nested* within the ´For i=´ and next i statements, creating a loop within a loop

While this approach worked great for our first run through the data, a few improvements could still be made to the code.

### Part II
The second part of the project called to reflector the code such that it only loops through the rows once to produce the same outcome.  Given this condition, nested loops are out - we need to create a ´Index´ that will output the data we want as three seperate arrays as shown here:

![tickerIndex](/../Resources/tickerIndex.png)

While you’ll notice I use i, j and k in this example to differentiate between variables used in different steps, I could just have easily have used i throughout the code since it is definitially closed off at each loop
 
![I_K](/../Resources/i_j_k.png)

And Voila!  A process to loop through the rows of data just once by using the ´tickerIndex´
 




## Summary
Refactoring code is as much an art as it is a science.   Throughout the research for this project, more advanced topics approaches came up, such as setting the ´tickerIndexas Single´ and using ´ReDim´ to create the output array:

 ![ReDim](/../Resources/improvement.png)

To borrow another overused phrase, refactoring code is less a destination than a journey.  The need to ship a Minimum Viable Product may take precedence over producing perfect code, but improvements can always be added to product in future updates to improve functionality and respond to a new operating environment.

In this case, the positives of the refactored code seem to be a clear win for the refactored code in terms of reduced processing time as demonstrated here: 

![original](/../Resources/orig_run_time.png)

!(refactored](/../Resources/refactored_run_time.png)

However, depending on the audience it is possible that the run time may be less important than providing a tool that is basic enough for employees to use, build and learn from.  In case the original code using ´nested loops´ may be preferable - it’s all about what the tool is best used for!
