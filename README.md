# VBA Refactor Challenge 

## Overview of Project

### Purpose:
I conducted an analysis on stock data for Steve on Green Energy companies to help diversify his parent’s portfolio. To provide better service for Steve and his parents, I refactored the VBA script that helps analyze the stock data to allow for more efficient execution times. The refactored VBA script will allow for Steve to look at more stocks while being more efficient.
## Results

### Analysis between Performance of the original script and the refactored script:
The original script code worked to help find the data we needed for our small amount of data we wanted to analyze. When we ran our original script for 2017 the execution time was .5625 seconds while 2018 was .6171 seconds. Our execution times were longer than our refactored code. The main reason for this longer time was due to how our code was executing. In our original code we used a nested for loop which went through 3000 rows of data for each ticker. Since we have 12 indexes, we would have go through our 3000 cells of data 12 times, which is why our execution times were longer in our original code. Our code would also initialize total volume to 0 for each run of an index for that certain ticker. 


![Old code snippet]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/Old%20code%20snippet.png)

![2017 Old Code Time]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/2017%20Old%20Code%20Time.png)
![2018 Old Code Time]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/2018%20Old%20Code%20Time.png)

The refactored script was much more efficient compared to our original script. 

![VBA_Challenge_2017]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)








## Summary

### Advantages and Disadvantages of Refactoring Code:
Refactoring code has many advantages such as improving legibility that improves the comprehensibility of the code for other programmers. Making the code easier to understand helps with maintenance and extendibility of the program in the long run. Refactoring can remove redundancies and duplications that will improve the effectiveness of the code. 
Refactoring has some great advantages but also brings disadvantages with it. Inaccurate refactoring can introduce new bugs and errors into the code, which is why it is important to keep track of changes. Another disadvantage is refactoring can be very time consuming due to running into bugs and new errors. A team member can make a change to the code and if they did not keep track of their addition, can cause an error which will cause the team to debug to fix that error. 
### Pros and cons of refactoring the original VBA Script


