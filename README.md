# Stock Analysis--Refactoring VBA Code and Performance Measurement
A Refactoring Project for Code Efficiency

## Overview of Project
Microsoft Excel macros using Visual Basic for Applications (VBA) were created to analyze 2017 and 2018 stock performance for 12 different tickers of alternative energy. The original VBA script was refactored to improve performance.

### Purpose
The purpose of this project was to compare the original VBA script to the refactored script and show improvement in runtime speed. 

#### Data: Original VBA Script (https://github.com/KimberlyCrawford/Stock-Analysis-Refactoring-VBA-Code-and-Performance-Measurement/blob/main/Original_code.txt) versus Refactored VBA Script (https://github.com/KimberlyCrawford/Stock-Analysis-Refactoring-VBA-Code-and-Performance-Measurement/blob/main/Refactored_code.txt)

## Results: 
Below are execution results for original and refactored script:

### Original Scripts for 2017 and 2018
![Stock_2017_Original_Runtime.jpg](https://github.com/KimberlyCrawford/Stock-Analysis-Refactoring-VBA-Code-and-Performance-Measurement/blob/main/Stock_2017_Original_Runtime.jpg) ![Stock_2018_Original_Runtime.jpg](https://github.com/KimberlyCrawford/Stock-Analysis-Refactoring-VBA-Code-and-Performance-Measurement/blob/main/Stock_2018_Original_Runtime.jpg)

### Refactored Scripts for 2017 and 2018
![Stock_2017_Refactored_Runtime.jpg](https://github.com/KimberlyCrawford/Stock-Analysis-Refactoring-VBA-Code-and-Performance-Measurement/blob/main/Stock_2017_Refactored_Runtime.jpg) ![Stock_2018_Refactored_Runtime.jpg](https://github.com/KimberlyCrawford/Stock-Analysis-Refactoring-VBA-Code-and-Performance-Measurement/blob/main/Stock_2018_Refactored_Runtime.jpg)

## Summary

- Advantages and disadvantages of refactoring code: Refactoring code makes the code more efficient by taking fewer steps, using less memory, and improving the logic of the code to make it easier for future users to read. It does not add more functionality and can be risky when working with large datasets. 
- Pro and cons of refactoring code specific to this analysis: Refactoring code for this stock analysis proved beneficial as runtime was decreased by 66% (2017: Original runtime = 1.320313 seconds versus Refactored runtime = 0.4453125; 2018: Original runtime = 1.273438 seconds versus Refactored runtime = 0.3984375). Refactoring for this small set of data worked well; however, larger datasets may cause problems when refactoring.

#### Module 2, Data Analysis & Visualization Certificate Program, UT Austin McCombs School of Business, 2021.
