# VBA Challenge - Refactoring Code

## Overview of Project

This project is for Refactoring of VBA Stock Analysis Code previously Written.

### Purpose

The purpose of this project is to Refactor/Enhance a VBA code. The VBA code was already written to perform analysis on a stock data set. The goal is to improve the runtime/performance efficiency/speed.

The initial data had 3000 observations of 12 stock symbols to perform the analyses on.  Although the initial data didn't take long to compute, adding a larger set of data to compute could take much longer to run.

## Results

After Refactoring, both versions of the code were tested to see their runtime performance for 2017 and 2018 year.

The results of these tests(in seconds) :

|  code          | 2017 | 2018 |
| ---------------| --------------- | --------------- |
| Original | [**0.75**](https://github.com/nidhipandya/stock-analysis/blob/main/OriginalCode_2017.PNG) | [**0.765625**](https://github.com/nidhipandya/stock-analysis/blob/main/OriginalCode_2018.PNG) |
| Refactored | [**0.10937**](https://github.com/nidhipandya/stock-analysis/blob/main/VBA_Challenge_2017.PNG) | [**0.1015625**](https://github.com/nidhipandya/stock-analysis/blob/main/VBA_Challenge_2018.PNG) |

Here the dataset is small and the runtime performance is about 6 times faster with Refactored code. Refactored code runs the 'for loop' one time only so that's the reason for its better performance.

## Summary

### Advantages of Refactoring Code

The Refactored code increases the runtime efficiency.

### Disadvantages of Refactoring Code

Refactored code uses about 4 arrays, so it needs more memory to store the data. while the original code had only 1 array but it had a nested loop so it took a longer time to run the code.
