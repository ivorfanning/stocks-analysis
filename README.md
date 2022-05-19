# stocks-analysis

### Overview of Project: 

  *The purpose of this analysis is helping Steve to analysis which green stocks was traded frequently and has a better return in 2017 and 2018, by the end of this analysis, Steve could help his parents to get a big picture to see if the green stocks are worth to invest and if so, which stock has more potentials.*

### Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

  *By visualizing the “All Stocks Analysis” table and return chart for 2017, most of the green stocks have positive annual returns except stock ticker “TERP”, and the stock ticker “DQ” which is the one Steve’s parents have most interested in, its annual return is close to 200%.* 
  
  ![All Stocks Analysis 2017](https://github.com/ivorfanning/stocks-analysis/blob/main/All_Stocks_Analysis_2017.png)
  
  *However, as we can see from the table analysis and return chart for 2018, only two green stocks are showing positive return, the others are negative returns which “DQ” also has the lowest return.*
  
  ![All Stocks Analysis 2018](https://github.com/ivorfanning/stocks-analysis/blob/main/All_Stocks_Analysis_2018.png)
  
  *From the above analysis, we can tell most of the green stocks are very fluctuate and the stock “DQ” is most fluctuated stock in the green stocks selection. therefore, the risk of investing in green stock is high, and it may suit for small percent of investor's portfolio. We did not know what happened in 2018 cause the return negative, in order to make the decision, we need more news and data for green stocks, especially for "DQ".*

  *As per the runtime message box, the run time of original code is around 0.9 seconds, and after refactoring the code, it reduced to 0.1 seconds.That is a big improvment because we used a more efficient method “Array” to store our values.* 
  
  ```
   Dim tickerIndex As Integer
       tickerIndext = 0
   Dim tickerVolumes(12) As Long
    
   Dim tickerStartingPrices(12) As Single
    
   Dim tickerEndingPrices(12) As Single
   
  ```

### Summary: 

  #### What are the advantages or disadvantages of refactoring code?
  
     Advantages of Code Refactoring:
     
        1.	Code Refactoring makes the code more flexible and extensible for adding many other functions to it. 
        2.	After refactoring, the code is easier to read and maintain.
        
     Disadvantages of Code Refactoring:
     
        1.	It is a time-consuming process, and we don’s know how much time it may take to complete the refactoring code
        2.	In case if the code went wrong, we will have to waste more time to debug it due to complexity of the code.

  #### How do these pros and cons apply to refactoring the original VBA script?
  
    Refactoring this stock analysis code is more efficient and understandable by using different arrays 
    to store the total volume, starting price and ending price.
    When we analyze other different stock dataset, the only thing we need to change is the size of arrays
    and the ticker list inside the array, so it may suitable for other different stock analysis in the in future.

