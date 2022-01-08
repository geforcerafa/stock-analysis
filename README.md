# The VBA of Wall Street

Data analysis project using VBA to analyze companies of the green energy sector.  

## Overview of Project
Our friend Steve graduated from finance and his parents start investing with him. They believe that the fossil fuels are drying up and the future is for the green energy. Steve’s parents invested in DAQO, ticker DQ, but Steve wants to diversify the portfolio in more stocks and start to analyze other companies. 

### Purpose
We are using VBA to make macros that help Steve analyze more stocks in an efficient way. This helps him to analyze more companies in the green energy sector, and find better investment opportunities. 

## Results

DQ was down 62.6% in 2018 so we started to improve the macro to process information of more stocks. 


To look for better investing opportunities, we refactor this macro to analyze more stocks using loops and If-Then to go through the list of tickers and organize the data from the tickers using arrays. 

At the end we improve the results view with some static and conditional formatting, so it is easier to read. 

Refactoring the code helped to make it lighter and easier to execute. There was a clear improvement in the time of execution. In the following images we can see how the time of execution before the refactoring for the analysis of the year 2017 was 1.298828 seconds, and after it was 0.96875 seconds.

#### Performance before refactoring for the year 2017
![Analysis 2017 slow VBA](https://user-images.githubusercontent.com/96758511/148653012-8edeeb8b-0a96-443e-9bac-78d350979181.png)

#### Performance after refactoring for the year 2017
![Analysis 2017 refactored](https://user-images.githubusercontent.com/96758511/148653339-9f30d7be-2ec7-4c1a-90e9-559f2c7d8f75.png)



## Summary

We saw a great improvement in the execution time of the Macro with the refactoring. It was lighter since we change variables from double to single data type. That saves memory in the execution and makes it easier to run.

The greatest improvement in the refactoring was the last if-then added. This one check if the next row’s ticker is different, the tickerIndex changes and doesn’t have to go through all the data checking one by one.

The following are the outputs of the analysis of all stocks for the years 2017 and 2018.

#### Analysis of results for the year 2017
![Analysis 2017 refactored](https://user-images.githubusercontent.com/96758511/148655220-3a2e3424-ebea-492f-9acb-9edccccc1742.png)

#### Analysis of results for the year 2018
![Analysis 2018 refactored](https://user-images.githubusercontent.com/96758511/148655232-c1a2a481-913b-45bb-bbb1-ae71d7bc9943.png)



### General advantages and disadvantages of refactoring code
 We can say that in general:

#### Advantages
•	You don’t start from scratch, but from a solution to a problem that already works. 
•	Get new ideas of how to write that code, learn different ways to solve coding problems. 
•	The pattern of steps to reach that solution is most likely already there, in some place, you just have to find it and adapt it.

#### Disadvantages
•	You have to be very careful with the variables in the code you get, and the ones in your code, they have to be the same.
•	The resulting code may not be working to solve the need but may not be optimized for this function, so there is more work to do in making it more efficient. 


### In this case, the advantages and disadvantages of the original and refactored VBA script

#### Advantages
•	The code runs much faster after the refactoring, as we could see in the 2017-year example.
•	Making the variables single form double data type, saves memory and it is a lighter macro for the computer to run.


#### Disadvantages
•	The data has to be well ordered, because the refactoring didn’t check one by one all the tickers.
•	In this case the advantages were clear over the disadvantages, we can say that may be the time and effort to refactor the code, but it is totally worth it.
