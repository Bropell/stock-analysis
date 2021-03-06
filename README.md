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

This large difference in execution time comes from refactoring the code by using multiple arrays. A variable called tickerIndex was defined and would be used as the argument for each array to be looped over. Three output arrays for the parameters of interest were defined in size and object type in order to store the values from the loop. A visual summary can be seen in the screenshot below:

![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/For_Loop_1.png)

Finally, loop through the arrays to output each value to its corresponding area in the spreadsheet. The loop and output locations are shown in the image below. Using arrays in this way is much faster than trying to loop through an array, loop through rows and then outputting the data into the spreadsheet with a nested loop. 

![alt text](https://github.com/Bropell/stock-analysis/blob/main/Resources/For_Loop_2.png)


## Summary

- What are the advantages or disadvantages of refactoring code?

There are several advantages and disadvantages of refactoring code. Some advantages of refactoring include reduced execution times and improved software design. As seen from the results of this analysis, after the code was refactored the execution time went down by a full order of magnitude. Over much larger data sets, this time save will eventually start to add up. Improved software design and the ability to refactor depends on a strong understanding of how the code works and it can extend the life of the code by designing it to be easy to understand and maintain. Some major disadvantages of refactoring code, as seen with this challenge, was the increased complexity of writing the code and how time consuming it can be. For the one writing the code, imprecise refactoring can introduce many new bugs that were not present in the original script. It can be very time consuming to find and fix these bugs in addition to trying to write the code in the first place. 

- How do these pros and cons apply to refactoring the original VBA script? 

The advantages and disadvantages previously mentioned most certainly apply to refactoring the original VBA script. As seen in the screenshots of the execution times, the run times of the refactored code is much faster than the original script. This will be a huge advantage to steve as he expands his search to the entire stock market. The refactored code also includes the formatting function for the cells all in one macro whereas the original had a separate macro to do the formatting which takes even more time. Refactoring the original script for Steve also brought its disadvantages, mostly to the person writing the code, in the form of time consumption and presence of new bugs. There were a lot of bugs that needed to be diagnosed and fixed here compared to the original script, which was advantageous in the learning aspect but not ideal for time purposes. Changing the script to loop using multiple arrays required a higher level of understanding and naturally took more time to write and run without errors. 
