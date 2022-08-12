# All Stocks Analysis

The initial requirements for this analysis were to help Steve determine which green energy stocks would be the best ones for investment.

After running a report on green energy stocks, Steve requested a more robust analysis to see if there were opportunities in other sectors that would be prove to be worth investigating further.

Once the requirements changed to allow Steve to analyze all stocks and would require the system to run through thousands of stocks, it became apparent that the code needed to be much more efficient.

The goal of this project is to take the original green_stocks analysis and refactor the code to be much more efficient.

---
## Results

The following screen shots show results of running the reports on both 2107 and 2018 data both before and after refactoring.

### 2017 Results Before Refactoring

![2017_Results_After_Refactoring.png](./resources/2017_Results_After_Refactoring.png)

### 2017 Results After Refactoring

(TODO: screenshot of 2017 AllStocksAnalysisRefactored performance results)

### 2018 Results Before Refactoring

(TODO: screenshot of 2018 AllStocksAnalysis performance results)

### 2018 Results After Refactoring

(TODO: screenshot of 2018 AllStocksAnalysisRefactored performance results)

---
## Summary

After an analysis of the code it became apparent that the nested loop could be performed with just one iteration through the arrays. By doing this the number of iterations is cut from N^2 to N. (Note: I do not remember this formula exactly however an exponential number of steps is very costly compared to a linear number.

(TODO: Detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).)

---
## Challenges Encountered

During this challenge there were a few obsticals that made it take longer than it should have:
- In Excel VBA's editor, it will often come up with a generic message stating that the value of an object or variable is out of range.  This can take time getting used to understanding how to find this error. Most likely some variable sent to determine a Cell on a worksheet is out of range either because it was not initialized. Another possibility is that the value went over the array boundaries.  In either case, the program will seem like it stopped running but if the stop button is not pressed the program will try to continue where it stopped and it will continue to fail. Once the stop button is pressed, the program will reset so it runs from the beginning with the updated code. This happened for a while where it was unclear why changes to the program did not take effect and took more time up than it should have.

- At first when I added a new subroutine to format data it was not working but when I went back to it, I noticed I needed to pass in some variable that were no longer being set and that fixed the problem.  By putting the formatting in a separte routine, it not only cleans up the loop that sets the data it also makes updates to formatting easier in future releases.  Since it only adds 1 iteration through the data vs. an exponential number it should not affect performance making code readability more important.
