# Stock Analysis
A multi year analysis of stock price performance using VBA macros to quickly and easily guage trends with a variety of tools.  The project then explains the process for refactoring to determine if the new VBA macros written is appropriate for use with larger data sets such as the S&P 500 and Russel 2000.


## Overview
The analysis is presented in two parts:  an initial survey of VBA macros to compare the perofrmance of a handful of stocks from 2017 and 2018, then refactoring the code used to make the macros run more efficiently to handle larger data sets.  Code refactoring is a process [defined as](https:/refactoring.com/)

> a disciplined technique for restructuring an existing body of code, altering its internal structure without changing its external behavior.

In this case to shorten the run time and therefore improve its performance.  

## Results

### Part I
The first project created an array of stock tickers shown here:

![/../stock_analysis/Resources/Array of tickers.png)](https://github.com/halesoysters/stock_analysis/blob/b49559c8231fd73414b18201bb81df2db600bb5a/Resources/Array%20of%20tickers.png)

For each element a **nested loop** was used to design a `for` loop inside a `for` loop demonstrated here:

![nested loop](https://github.com/halesoysters/stock_analysis/blob/b49559c8231fd73414b18201bb81df2db600bb5a/Resources/nested%20loop.png)

Notice how the `For j = ` starting the nest and `next  j` definition  is *nested* within the `For i=` and next i statements, creating a loop within a loop.

While this approach worked great for our first run through the data, a few improvements could still be made to the code.

### Part II
The second part of the project called to refactor the code such that it only loops through the rows once to produce the same outcome.  Given this condition, nested loops are out - we need to create an `Index` that will output the data we want as three seperate arrays as shown here:

![tickerIndex](https://github.com/halesoysters/stock_analysis/blob/b49559c8231fd73414b18201bb81df2db600bb5a/Resources/tickerIndex.png)

While you’ll notice I use i, j and k in this example to differentiate between variables used in different steps, I could just have easily have used i throughout the code since it is definitially closed off at each loop.
 
![i_K](https://github.com/halesoysters/stock_analysis/blob/b49559c8231fd73414b18201bb81df2db600bb5a/Resources/i_j_k.png)

And Voila!  A process to loop through the rows of data just once by using the `tickerIndex`.
 



## Summary
Refactoring code is as much an art as it is a science.   Throughout the research for this project, more advanced approaches came up, such as setting the `tickerIndex` as Single and using `ReDim` to create the output array:

 ![ReDim](https://github.com/halesoysters/stock_analysis/blob/b49559c8231fd73414b18201bb81df2db600bb5a/Resources/improvement.png)

To borrow another overused phrase, refactoring code is less a destination than a journey.  The need to ship a Minimum Viable Product may take precedence over producing perfect code, but improvements can always be added to product in future updates to improve functionality and respond to a new operating environment.

In this case, the positives of the refactored code seem to be a clear win in terms of reduced processing time.  This could be very helpful for use with a very large data set - say like spotting trends within the component companies of the S&P 500 or Russell 2000 over the past decade - as demonstrated here: 

![original](https://github.com/halesoysters/stock_analysis/blob/b49559c8231fd73414b18201bb81df2db600bb5a/Resources/orig_run_time.png)

![refactored](https://github.com/halesoysters/stock_analysis/blob/b49559c8231fd73414b18201bb81df2db600bb5a/Resources/refactored_run_time.png)

Depending on the audience, however, it is possible that the run time may be less important than providing a tool that is basic enough for employees to use, build and learn from.  In such an example, the original code using nested loops may be preferable - it’s all about what purpose the tool is best suited for!
