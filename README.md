# stock-analysis
## Overview of Project

The purpose of this project was to get us more familiar with Visual Basic for Applications (VBA) programming in Excel. We learned about basic programming fundamentals such as notation, as well as becoming more knowledgable about for loops and arrays. Additionally, we studied how to manipulate data through automation, especially in regards to using the same program to analyze data from different worksheets. 

## Results

As you can see in the two pictures below, the remastered script AllStocksAnalysisRefactored was able to finish analyzing 2017's stocks in around 0.195 seconds, and 2018's stocks in 0.172 seconds.

![Remastered Script - 2017 stock analysis](https://github.com/li-emily/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![Remastered Script - 2018 stock analysis](https://github.com/li-emily/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

While not pictured, the original script AllStocksAnalysis finished analyzing 2017's stocks in 1.063 seconds, and 2018's stocks in 0.95 seconds. As you can see, the entire process was made faster by around 0.8 seconds for both.This also does not include visual formatting with colors to indicate if the return was positive or negative, which was put into a separate script after called formatAllStocksAnalysisTable. 

## Summary
- What are the advantages or disadvantages of refactoring code?
  - The main disadvantage of refactoring code is that it will initially take longer to modify the original code than simply just using it, first with writing the new code, and new possible bugs that will each take time to fix. However, the advantages of refactoring are readability, speed and smoothness, and the ability to quickly modify the code to then be more robust in the future.

- How do these pros and cons apply to refactoring the original VBA script?
  - The original VBA script did have notes which mentioned what was being done each step. However, it did not accomplish everything in one script, as we wrote several individual scripts which were ran one after the other to achieve the final result. In the end, both the original script and the refactored script both had the same output, but the new code took a significantly longer time to rewrite and figure out, which is the biggest con. However, by refactoring the code, we were able to combine the several scripts together, and also improve user experience by streamlining the process to be faster as well.
