# stocks-ananlysis

## Overview of Project

Steve's parents want to know how actively DQ was traded in 2018. They believe that if a stock is traded often, then the price will accurately reflect the value of the stock. If we sum up all of the daily volume for DQ, we'll have the yearly volume and a rough idea of how often it gets traded.But we don't want to calculate the numbers by hand, we want to have a more efficient way to do it. VBA is a quicker and more accuate way to do it once we set up the code right.So I helped him to set up the VBA code to calculate each stock's volumes and earnings back in 2017 and 2018. I also used VBA code formatting the results. Like make the cell of returns above 0 to be green and below 0 to be red. 

Steve loves the workbook I prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although my code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, I edit, or refactor, the [Module 2 Challenge](https://courses.bootcampspot.com/courses/2638/assignments/45050?module_item_id=835412)solution code to loop through all the data one time in order to collect the same information that I did in this module. I will analysis the advantages and disadventages of refactoring the code. And describe my result about how efficent it is.

## Results

After running the new coding, 2107 file and 2018 file used the same time, 0.289 seconds.This is much quicker than the original one ![This is an image](https://github.com/YawenShao0902/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png)/![This is an image](https://github.com/YawenShao0902/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png).

## Summary

### There is a detailed statement on the advantages and disadvantages of refactoring code in general

* Advantages: Refactoring is intended to improve the design, structure, and/or implementation of the software (its non-functional attributes), while preserving its functionality. Potential advantages of refactoring may include improved code readability and reduced complexity; these can improve the source code's maintainability and create a simpler, cleaner, or more expressive internal architecture or object model to improve extensibility. Another potential goal for refactoring is improved performance.
* Disadvantages: Hard to find any problems if it runs wrong. We need to be careful every line to make sure it runs. We also need to think a better way to do the coding but sometimes it is even harder than creating a code.

### There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script 

* Advantages: Original VBA script has more steps, so it is more understandable. The refactored VBA script is less steps and delete some unnessary, for example I deleted "'1b) Create three output arrays".

* Disadvantages: Original VBA script has more steps, so it has some steps are not necessay, more work to do. But at the same time, the refactored VBA script is more easy to be destoried if I am not be careful. And since I deleted some steps, so maybe it is less understandable to other people.
