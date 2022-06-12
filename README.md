# Stock-Analyis with VBA 

## Overview of Project
This project dives into Visual Basic for Applications (VBA) in order to help Steve analyze green energy stocks to invest his parents money. A macro was created for a list of stocks provided by Steve in order to determine total daily volume and the annual return for each stock with a specific focus on the DAQO New Energy Corp stocks (DQ). Finally, to do a little more research for his parents, the code was refactored so that Steve can expand his dataset to include the entire stock market over the last few years. 
## Results
https://github.com/Bropell/stock-analysis/blob/main/VBA_Challenge.xlsm
### Stock Performance
The were a couple of things to comment on in terms of stock performance between 2017 and 2018. The most obvious point here was that the annual return percentages were all positive in 2017 with the exception of the "TERP" stock whereas for 2018, almost every stock, with the exception of two, had a negative annual return percentage. This can easily be seen with the color scheme used to label positive percentages (green) and negative percentages (red) displayed in the screenshots below. 

![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/All_Stocks_2017.png)


![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/All_Stocks_2018.png)


This analysis was very insightful for Steve and his parents since they plan on investing all of their resources into only the DQ stock. If they had done so in 2017 they would have made a large return but would have taken a hit the following year. It would likely be unwise to invest so much into a stock that fluctuates between such large ranges of return. Most stocks did however have a higher total daily volume in 2018 compared to 2017 which means there is more money being funneled into the stock but the return is certainly more important.
### Execution Times: Original Script Vs Refactored Script
Aside from the stock performance itself, there was a significant difference between the execution times of the original script and the refactored script. In fact, there was a whole order of magnitude difference, even though the original script did not include a cell formatting function whereas the refactored code did. The time differences can be seen in the screenshots below. 

2017 Original: ![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/Original_Script_2017.png)

2017 Refactored: ![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

2018 Original: ![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/Original_Script_2018.png)

2018 Refactored: ![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

This large difference in execution time comes from refactoring the code by using arrays. A variable called tickerIndex was defined and would be used as the argument for each array to be looped over. Three output arrays for the parameters of interest were defined in size and object type in order to store the values from the loop. A visual summary can be seen in the screenshot below:

![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/For_Loop_1.png)

Finally, loop through the arrays to output each value to its corresponding area in the spreadsheet. The loop and output locations are shown in the image below.

![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/For_Loop_2.png)


## Summary

- What are the advantages or disadvantages of refactoring code?

There are several advantages and disadvantages of refactoring code. Some advantages of refactoring include: 