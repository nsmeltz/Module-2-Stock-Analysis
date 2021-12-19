# Module-2-Stock-Analysis

## Overview of Project: Explain the purpose of this analysis.
   The purpose of this analysis is to evaluate the stock performance. The client, Steve, is interersted in the performance of 12 specific stocks in 2017 and 2018. He has a    spreadseet of data that gives the ticker name of the stock, date, opening price, high price, low price, closing price, adjusted closing price, and volume (ie number of stocks). This analysis determines the total daily volume for each stock (ticker) and the yearly return(color coded by positive or negative gains).      

## Results & Analysis 
  The following images show the analysis result in Excel derived from the original code (top) and refactored code(bottom) for each year with pop-up windows showing run time for the execution of the code.
  
  - **2017 Stock Performance**

![2017 Original](https://github.com/nsmeltz/Module-2-Stock-Analysis/blob/16553dd62d3dbc7f707c5c20db96313ee8f33b55/Resources/2017_original.png "2017 Original")

![2017 Refactored](https://github.com/nsmeltz/Module-2-Stock-Analysis/blob/e4ebac96b39ec2300f349af462cf71b2827291f3/Resources/2017_refactored.png "2017 Refactored")
 
  - **2018 Stock Performance**

![2018 Original](https://github.com/nsmeltz/Module-2-Stock-Analysis/blob/e4ebac96b39ec2300f349af462cf71b2827291f3/Resources/2018_original.png)  

![2018 Refactored](https://github.com/nsmeltz/Module-2-Stock-Analysis/blob/e4ebac96b39ec2300f349af462cf71b2827291f3/Resources/2018_refactored.png)

## Analysis
  - **Stock Performance 2017 vs 2018**
    From the following graphs I can see two major trends: 1) the positive yearly returns are generally greater in 2017 than in 2018 and 2) the total daily volume of stocks dropped from 2017 to 2018. These trends suggest that the market for the analyzed stocks is falling between the years 2017-2018. However, there are two stocks that do not follow this trend, ENPH and RUN. Both of these stocks had positive returns and higher total daily volume in 2018 than 2017. The only stock to have a better return in 2018 than 2017 is RUN suggesting that this stock is growing and could be a good one to invest in. 

![TotalDailyVolume](https://github.com/nsmeltz/Module-2-Stock-Analysis/blob/6eeb04f64e5f1894c1299b495fd44df1a651f4ef/Resources/TotalDailyVolume.png)
![YearlyReturn](https://github.com/nsmeltz/Module-2-Stock-Analysis/blob/6eeb04f64e5f1894c1299b495fd44df1a651f4ef/Resources/YearlyReturn.png)

  - **Code Performance Original vs Refactored**
    The major change that I made between the original and refactored code was to write the calculated values(ticker, daily volume, and yearly return) into arrays then output the values into the appropriate cells after all loops were complete vs outputting them to the cells after every iteration of the inner for loop. The resulting run times for the original ve refactored code for this data set are not significantly different for either year. If I applied the two ways of outputting the result on a larger data set with more rows of data, it would probably make a much greater difference in run time. One thought I did have about using the timer function is that if you initialize the timer prior to the input box then your time data will include how much time it takes for the user to input the year into the box and the loops.  
  

## Summary 
  - **What are the advantages or disadvantages of refactoring code?**
    The advantage to refactoring code is that you have the abliity to make your code more efficient and faster, by applying new ways of storing the data and manipulating it. Usually the first time you write a script you are just trying to get the code to run and accomplish the task your are trying to complete. From my expereice in Matlab I have frequently written code and gone back through and noticed different ways that I could write it to be more efficient or even less complicated. I think that is is easier to refactor as you learn more tips and tricks on how to code efficiently. I think that the most annnoying thing about refactoring is that you will inevitably break your code that was running while trying to make it better so sometimes it can be time comsuming to figure out how to fix the new mess you made.  
  - **How do these pros and cons apply to refactoring the original VBA script?**
   These pros and cons apply to the VBA script for the stock analysis because I definitely broke the code that I had working when I tried the new output method to the arrays. I had to figure out what the syntax was for initilizing the variables as arrays in order to fix the problem. Since the run time for this data set was not significantly improved I probably would not have bothered changing the way the data was output to the cells. With larger datasets though the array method will definetly help speed up the run time and keep the code more organized
