# Stock Analysis using VBA

## Overview
The pupose of this project was to create a tool within Excel using VBA coding in order to analyze stock volume and performance within a given year's timeframe. Further, it was necessary to refactor the code to ensure the analysis was performed in a more efficient manner.

## Analysis
In order to perform this analysis, I first needed to understand the overall dataset. The data to be analyzed encompassed a collection of 12 stocks over the course of the years 2017 and 2018. Within this data, information was contained reflecting each stocks' performance details organized by ticker and date. Pertinent data to this analysis was the stocks' ticker, open price, close price, and volume.

First, I created a macro called AllStocksAnalysis(). This macro was a first attempt at utilizing VBA to create a table reflecting a listing of all 12 tickers showing their associated total daily volumes along with the return for the year. This macro included a message box prompting the user to enter their desired year (2017 or 2018) for the analysis. Additionally, I included code to start a timer at the initialization of the macro as well as and code to stop the timer once the macro completed. This also contained a message box that would let the user know what the time was. The key item of code used in this macro was an array of all of the tickers. This array is how the macro was able to capture and calculate all of the values correctly. Below is a snapshot of the coded array using tickers(12) as a string.
    
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR" 

Following this macro I proceeded to refactor the code by creating a new macro called AllStocksAnalysisRefactored(). The main change, with the intent to speed up the runtime, was the following. Instead of the code being designed to with a For Loop within a For Loop cycling through the array, I created a new variable called tickerIndex. The purpose of this was for it to act as a counter, essentially counting its way through the array illustrated above. After running through tickers(0) for ticker "AY" it would then be incremented up 1 to then run through the same process for the subsequent ticker "CSIQ".

### Results & Runtime Comparison
![VBA Challenge Time Comparison](/Resources/VBA_Challenge_Time_Comparison.png "VBA Challenge Time Comparison")

## Summary
In summation, it can be seen from the above image that the refactored VBA code was able to speed up the time for the macro to process through the dataset. In general, refactoring code can be a useful process. Especially if the existing code is such that its inefficiency results in the macro getting bogged down. There can be multiple ways of solving the same problem and therefore taking the time to review your code and brainstorm on better ways of breaking down the steps can be very useful. Alternatively, there can be a downside. Not every situation calls for this. If the dataset is not large and the steps necessary to analyze it are not complex, the time taken to refactor might not net a benefit worth cost in time. In this exercise you can see that the runtime for 2017 (0.140625 vs 0.84375) and 2018 (0.1523438 vs 0.8476563) were much faster than the original code. Clearly in this case, the refactored version did make a large difference in runtime to achieve the same result.
