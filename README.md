# VBA Refactor Challenge 

## Overview of Project

### Purpose:
I conducted an analysis on stock data for Steve on Green Energy companies to help diversify his parent’s portfolio. To provide better service for Steve and his parents, I refactored the VBA script that helps analyze the stock data to allow for more efficient execution times. The refactored VBA script will allow for Steve to look at more stocks in the future while being more efficient.
## Results

### Analysis between Performance of the original script and the refactored script:
The original script code successfully pulled the necessary data for analysis. The original subroutine ran the data for 2017 in .5625 seconds and for 2018 in .6171 seconds. The refactored code reduced the time for 2017 and 2018 in .4766 seconds (2017) and .5233 seconds (2018) respectively and included additional formatting code. The main reason for this longer time in the original subroutine was due to how our code was executing. In our original code we used a nested for loop, which went through over 3000 rows of data for each ticker. Since we have 12 indexes, we would have go through our 3000 cells of data each time we ran analysis on a ticker. Our code would also initialize total volume to 0 for each run of an index for that certain ticker.

Code
![Old code snippet]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/Old%20code%20snippet.png)


![2017 Old Code Time]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/2017%20Old%20Code%20Time.png)
![2018 Old Code Time]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/2018%20Old%20Code%20Time.png)

The refactored script improved efficiency over the original code. In the refactored code, the tickerVolumes array was isolated in its own for loop compared with the original, where it was nested. This initializes all ticker volumes to 0 in one cycle through all the rows, rather than a nested for loop to initialize tickerVolume to 0 then finding the data for the indexed ticker. We also added a tickerIndex to our if statements to be able to find our data for our tickers more efficiently. Our new execution time for 2017 was .0859 seconds and for 2018 it was .0937 seconds. This is a great improvement of efficiency between our original script and our refactored script.

![Refactored Code snippet]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/Refactored%20Code%20snippet.PNG)

![VBA_Challenge_2017]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018]( https://github.com/NickFoley47/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary

### Advantages and Disadvantages of Refactoring Code:
Refactoring code has many advantages such as improving legibility that improves the comprehensibility of the code for other programmers. Making the code easier to understand helps with maintenance and extendibility of the program in the long run. Refactoring can remove redundancies and duplications that will improve the effectiveness of the code. 
Refactoring has some great advantages but also brings disadvantages with it. Inaccurate refactoring can introduce new bugs and errors into the code, which is why it is important to keep track of changes. Another disadvantage is refactoring can be very time consuming due to running into bugs and new errors. A team member can make a change to the code and if they did not keep track of their addition, can cause an error which will cause the team to debug to fix that error. 

### Pros and Cons of Refactoring the Original VBA Script:
The pros of refactoring the original code was we improved our execution time for script. In our original script, 2017 took .5625 seconds while in the refactored script it took .0589. This will allow us to work with bigger data sets in the future while being more efficient. The refactored script is much more efficient than compared to our original script. 
The cons were we ran into new bugs and errors. While we were making the script more efficient, new errors were popping up from mistyping, wrong variable, and some other factors. This increased our time to finish the project, but in the end it made our script much faster and efficient. 

## Teammates

### Ashley Rock, Shruti Ramana and I collaborated together for this project. 

