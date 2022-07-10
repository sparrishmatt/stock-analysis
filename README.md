# Stock-analysis

## Analysis Overview
The goal of this analysis was to create a report in Excel using VBA which will be able to analyze stock market changes by year. The output sheet for this project includes stock starting prices, ending prices, total daily volume, as well as percent changes by year. 

A secondary goal of this was to refactor already created code in order to speed up our analysis macros in case the dataset was enlarged. Any future changes to the dataset such as adding more stocks or years to be analyzed are accounted for as long as no changes are made to the VBA code. 

## Results

### 2017
<img src="https://github.com/sparrishmatt/stock-analysis/blob/main/Resources/2017%20Report.png" height="280">

Eleven out of the twelve stocks analyzed had positive gains for the year 2017. The three most notable stocks by increase were DQ which increased by nearly 200%, SEDG at 184.5%, and ENPH at 129.5%. The only stock that fell was TERP which suffered a 7.2% decrease. It is worth noting that while DQ had the largest gains, it also had by far the smallest total daily volume which could cause the prices to shift more dramatically as opposed to those which had a much higher daily volume. 

---
### 2018
<img src="https://github.com/sparrishmatt/stock-analysis/blob/main/Resources/2018%20Report.png" height="280">

Ten out of the twelve stocks had negative gains for the year 2018. The two stocks that had positive gains were RUN at 84.0% and ENPH at 81.9%. Three stocks with the largest price drop were DQ at -62.6%, JKS at -60.5%, and SPWR at -44.6%. DQ had about three times as much total daily volume in 2018 compared to 2017, but the price fell dramatically. ENPH had large positive gain for the second year in a row, and appears to be the most consistent stock out of this group of twelve. 

---
### Code changes and execution time

The biggest difference between the original code and the refactored code is the use of arrays. By having the stock starting and ending prices, as well as their total daily volumes stored as arrays we are able to cut out a nested loop from the original code that cause a longer execution time for the code. These arrays are able to store multiple values, so here they are set to store twelve values for the twelve stocks that are being analyzed. 


<img src="https://github.com/sparrishmatt/stock-analysis/blob/main/Resources/Array%20creation.png" height="280">
<img src="https://github.com/sparrishmatt/stock-analysis/blob/main/Resources/Array%20output.png" height="280">

The for loop from the second screenshot allows all of the outputs to be inserted into the spreadsheet in order and quickly. An important question that comes after refactoring this code is about how much time has the refactor actually saved when so few stocks are being analyzed here. 

#### Original and refactored execution time for 2017

<img src="https://github.com/sparrishmatt/stock-analysis/blob/main/Resources/2017_Timer_Original.png" height="280">
<img src="https://github.com/sparrishmatt/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png" height="280">

#### Original and refactored execution time for 2018

<img src="https://github.com/sparrishmatt/stock-analysis/blob/main/Resources/2018_Timer_Original.png" height="280">
<img src="https://github.com/sparrishmatt/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png" height="280">

Saving around .4 seconds when the program is only working with 12 stocks shows the importance of this refactor. The amount of time that could be saved using extremely large datasets has to be great. 

---
### Pros and cons to refactoring code

The success of this refactor is clear when looking at the amount of time saved compared to the original. But is a refactor always worth it? For something like a stock market analysis where a person might want to analyze large datasets multiple times then a refactor is clearly worth it if a great time save can be guaranteed. However, if it couldn't be guaranteed that a refactor would save as much time, or if the code was being used less often or for smaller datasets then it might not be worth it. 

Refactoring does take time, and while this refactor was fairly small and guided that won't always be the case. Consider the case where Person A has to refactor the code of Person B, and Person B wrote much more code than in this VBA macro. Not only will Person A have to take time to understand how Person B wrote their code, they will have to then spend even more time on their refactor that could have been spent working on new projects. So yes, the pros outweigh the cons for this refactor, but should be considered on a case by case basis for other refactors.
