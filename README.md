# stockanalysis
VBA Module 2

## Overview of Project

### Purpose
Steve needs an analysis tool for a list of stocks.  I have created one with VBA that suits his immediate needs however how wants to be able to expand the data set.  I will need to refactor the code to be able to handle the increased scale without any decrease in performance.
## Analysis and Challenges
There are many ways to code for a specific outcome however efficiency and scalability can vary greatly between methods.  The current code is a mix of arrays, for loops and if then statements that can be reorganized to increase performance in the event more data is added for analysis. 

### Challenges and Difficulties Encountered
The largest difficulty was applying the use of a tickerIndex to efficiently execute For Loops through the 4 arrays.  While it was suggested to create a Loop that would modify tickerIndex based on cell data it proved much more efficient to use tickerIndex as the variable for the Loop. (Picture of code located in Resources folder)

 

## Results
The refactoring increased performance of the code specifically by creating a tickerVolumes array and using the tickerIndex to run through the set of arrays. 





## Summary
### Advantages and Disadvantages of refactoring 
The major advantage and general point of refactoring is to have a cleaner, more efficient and more understandable code that can be managed by multiple developers while maintaining maximum efficiency of computing resources. 
The disadvantage is often the time it takes to do so as well as the seemingly endless ways a code can be made “more efficient’ depending on who is looking at it.   
### Advantages and Disadvantages of our original and refactored VBA script
The original and refactored script had marginal difference in performance or functionality.  Given the tickers still needed to be written into the code to be included into the Ticker array Steve would not be able to scale it beyond the current data without our assistance in any case. The addition of the tickerVolume array seemed complicate rather than stream line the analysis given there is no other us for that information.   

