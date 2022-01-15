# Stock-analysis using VBA Refactoring
## Overview
### Purpose of the Project
Steve likes the analysis we already did, but want to extend this analysis to more than 12 stocks. We will be using existing code and make some changes, which is called Refactoring. In that process we will see whether we can make the code more efficient, by making changes in logic or taking fewer steps.
### Background
The initial analysis does the analysis for 12 stocks only. While this might work well for the small number of stocks, code is not scalable for a bigger data set. 
## Results
### Analysis
Looks like  green stocks did better in 2017 than 2018.
In 2017, all Steve's energy stcoks except for one are in green. ![2017_Stocks_Analysis_Data_Table](https://user-images.githubusercontent.com/56806834/149609003-2454ccd9-813b-4452-9961-6a9d85ca841f.png)

Performance of the same stocks in 2018 is not good at all. Only two stocks are in green.
![2018_Stocks_Analysis_Data_Table](https://user-images.githubusercontent.com/56806834/149609052-206fad3d-65bf-4c49-99d2-57f118c8e951.png)

All the stocks we just looked at are DAQO energy stocks. So Steve can suggest to his parents that they needs to diversify their stock portfolio. Now with the refactored VBA script he can add more stocks to the analysis. 

### Code changes
1. Instead of using  Tickers as individual values, Tickers are defined as an array. This aray will allow us to scale it.
![image](https://user-images.githubusercontent.com/56806834/149609208-35af4322-f663-4718-ac2e-ad482202438f.png)
2. TickerStartingPrices and TickerEndingPrices are also arrays, which will also help with scaling.
3. ![image](https://user-images.githubusercontent.com/56806834/149609574-a671635c-9676-4da7-af1d-6cf7b185118a.png)
4. With one loop through the data set we are setting the values for both the output arrays. (Please see image above)
5. And also we were able to drop the condition to check for 'DQ' ticker as we are doing the analysis for all the tickers.
6. Final change is for output. This was also done through arrays and loop instead of one variable.
7. ![image](https://user-images.githubusercontent.com/56806834/149609721-7b2bc016-3a73-4bd2-9799-5134b23987fd.png)

Both Set up and formatting is not changed in the Refactored code
### Summary
#### Advantages of Refactored code
1. cleaner and better readable code
2. Less complex code
3. More efficient code

#### Disadvantages of Refactored code
1. Time to understand the code that needs to be refactored

#### Advantages and Disadvantages of VBA Script
in this Refactored code:
1. We made the code scalable. 
2. Code will run faster. We will see the a bigger time difference if we have more tickers in the dataset. 
