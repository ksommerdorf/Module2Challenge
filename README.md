# Analysis of Green Energy Stock Returns
## Background of Project
The client, Steve, recently graduated college with a finance degree and wants to help his parents choose the best stock options to invest in. His parents have chosen to invest in alternative energy and have already decided to buy DAQO (DQ) stock. Steve has requested for his data on DAQO and other green energy stock to be analyzed. The stock data was provided in Excel format from the client and contains DAQO stock data along with 11 other green energy stocks for the years 2017 and 2018. 
#### Purpose
This project analyzes stock data in order to calculate the yearly stock return (how much your investment increased or decreased) and the volume of stocks sold for a given year. This analysis is performed by using VBA in order to automate tasks that can perform calculations and read and write to cells and worksheets. The original analysis for the 12 stocks has already been provided to the client and is being refactored in this project in order to be more efficient in analyzing bigger data. With this refactored code, Steve will be able to analyze more data in order to provide his parents with the best possible investment portfolio.

## Analysis and Challenges
First, I copied the original code that creates the input box, header row, ticker arrays, and activates the relevant worksheet. 

![Original_code1](https://user-images.githubusercontent.com/57520471/159149787-9ec367c4-013c-4df6-b0c7-1d08f904425c.png)

Then I refactored the code, given the instructions, that is laid out below. The instructions also explain the purpose of the code followed directly below it.

![refactored_code1](https://user-images.githubusercontent.com/57520471/159149793-e10e24e3-9873-4bcd-b09f-ba89d6fc627c.png)

The formatting section of the code is also copied from the original code and is not part of the refactored code. 

![refactored_code2](https://user-images.githubusercontent.com/57520471/159149838-f71e8a85-cb48-4cac-9765-6c7868a6d350.png)

## Results
### Stock Performance for 2017
![VBA_Challenge_2017_Chart](https://user-images.githubusercontent.com/57520471/159150364-f6acc487-35e7-4460-85ee-79e86f1d5964.png)

Overall the stocks had a positive return rate, except for TERP. DQ had the highest return rate of 199.4% but the lowest total daily volume of 35,796,200. SPWR had the highest total daily volume of 784,187,000 with a return rate of 23.1%.  The lowest return rate was TERP with -7.2% at 139,402,800 total daily volume.

The elapsed run time for the original code.

![original_code_time2017](https://user-images.githubusercontent.com/57520471/159150375-cbd2a5d4-1fce-478e-b4c6-bca18f6d9690.png)

The elapsed run time for the refactored code.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/57520471/159150513-d98b908f-3388-42d4-bcd9-7bbc7b0522f1.png)

### Stock Performance for 2018
![VBA_Challenge_2018_Chart](https://user-images.githubusercontent.com/57520471/159150388-5eb41fcb-c259-4d35-8079-0d5c9b2a07d3.png)

Overall the stocks had a negative return rate with the exception of ENPH and RUN. RUN had the highest return rate with 84.0% and DQ had the lowest return rate with -62.6%. ENPH had the highest total daily volume of 607,473,500 and AY had the lowest daily total volume of 83,079,900. 

The elapsed run time for the original code.

![original_code_time2018](https://user-images.githubusercontent.com/57520471/159150392-2710de81-312d-4bfd-83c8-932cf2d9435c.png)

The elapsed run time for the refactored code.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/57520471/159150399-e0912041-33bf-4d99-b51b-24296f2d96e6.png)

Based on the analysis from the 2017 and 2018 green energy stock data, 2017 had a great return rate whereas the return rate in 2018 plummeted. This shows that rates of return can go up and down drastically and more years of data need to be analyzed in order to determine the return of each stock over time. The new refactored code will be more efficient at handling more years of data compared to the original code because it runs about 85% faster. 

## Summary
Refactoring code creates a more efficient and concise code that provides more universal output and usage. Advantages of refactoring code include easier debugging processes, design improvements, faster run times, and less memory usage. Refactored code is more universal because the concise design of the code shapes an easier model for more people to understand and work with. However, the process of refactoring code can be tedious and labor-intensive. Also, the refactored code does not add more functionality to the code and could become error-prone if the data becomes more complex. 
In this project, the biggest advantage of refactoring the code included an almost 85% faster elapsed run time. Also refactoring the code took the functions of three codes and put the code into one concise function. This makes the code easier to read and run. However, a disadvantage of refactoring three functions into one is if an error occurs the whole function will not run versus just one of the functions not running. Overall refactoring this code will allow Steve to efficiently analyze more stock data for his parents and is overall beneficial to the client's request. 
