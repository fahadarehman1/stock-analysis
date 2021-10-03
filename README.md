# Stock Analysis with VBA

## Overview of Project

Steve just started in Wall Street and his parents are excited about his new endeavor. They asked Steve to analyze the 'DQ' stock and come up with the analysis. The Analysis include the performance from 2017 and 2018 year. Steve knew that in order to complete this analysis he needs to enable the macro's in Excel so analysis can be automated.

After configuring the VBA macro in Excel he found out that DQ stock did well in year 2017 with a whopping gain of 199% but in 2018 the stock went down 62%. This is concerning and now he want to analyze other stocks in the portfolio to gain some insights.

The other part of his ask was the refactor the code which was initially created and introduce the idea of using indexes as that should speed up the analysis for larger datasets.

## Analysis and Challenges

Selection Criteria:

- Code will be run using VBA macros in Excel
- Data is available for 2017 and 2018 years
- There are 251 days of trading available for both the years
- The code exists and needs be refactored for Indexing
- Make sure there's no filtering on the data enabled otherwise you'll get an error of: **runtime error6**

### Stock runtimes using Indexes:

Here are the runtimes before we changed the code or applied the indexes:

![2017 runtime(before refactor)](/Resources/orig_2017.PNG)                                          ![2018 runtime(before refactor)](/Resources/orig_2018.PNG)



After refactoring the code for best practices and introducing the concept of Indexes, the analysis ran in one sixth of the time, as shows below:

![2017 runtime(after refactor)](/Resources/VBA_Challenge_2017.PNG)                                   ![2017 runtime(after refactor)](/Resources/VBA_Challenge_2018.PNG)

We can clearly see the picture above that refactoring did work and enhanced the performance. Other things to take a note of is that initially I was running the code and some other macros in the same VBA file and I didn't see any changes in the runtime infact in some cases the loads ran longer and the runtime were around .934 seconds(with indexing in place). After removing the unwanted macros and only keeping the one, the code ran efficiently

### Stock Performance:

Lets compare the stock performance from 2017 to 2018:

![2017 Stock Performance](/Resources/2017_Stock_Performance.PNG)                                     ![2018 Stock Performance](/Resources/2018_Stock_Performance.PNG) 

The first obvious point is that stocks overall did poor in 2018 as compare to 2017. The stock DQ which was up roughly 200% in 2017, dropped 63% in 2018. This was clearly a market crash and affected all the stocks, if I was in Steve place I would have bought DQ in 2019 as when the market is down its the right time for institutional investor to jump in and benefit from it.



## Results

VBA Script and refactoring is recommended in this case but a larger dataset would have helped with the final code analysis. I found VBA script and refactoring code to be little cumbersome and does requires great knowledge of indexing, for and if loops. VBA in general is a great tool to know and having the possibility to automate some your work is endless. 
