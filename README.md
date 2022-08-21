# STOCKS ANALYSIS 
## PROJECT OVERVIEW
### Background
Steve and his parents are Investors who want to invest in a green stock energy business. Recently, Steve requested me to analyze the stock data from 2017 and 2018 to reveal the total daily volume and the overall return on investment for each stock. By analyzing the data from the excel worksheet, I designed an interactive workbook using VBA script that works well to show the total daily volume and the total return for each stock per year. The client was impressed with the worksheet, but he later requested a more flexible worksheet that would allow him to analyze more stock by clicking a button, as he wanted to expand his research beyond twelve stock within a short period processing of time.

### Purpose
The goal of this project is to improve the worksheets efficiency( reducing the execution time) by refactoring the original VBA script. To confirm the effieciency, we will compare the processing time beween the original worksheet and the refactored worksheet.

## RESULTS
### Analysis
The original code was executed at 0.6328125 and 0.6796875 seconds for 2017 and 2018 respectively. The refactored code was executed at 0.203125(0.2) for 2017 and 0.2109375(0.2) for 2018.
The workbook can now handle larger sets of data at any given year efficiently.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/109990578/185802993-7ac20f05-7f6c-4baa-9fc3-9c2322a151e8.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/109990578/185802996-798f27fa-0103-436f-8e56-8ff5b9abef64.png)

The original code took longer for its outputs because it utilized a nested loop that ran through every row 12 times in order to collect information for the 12 stocks. Whereas, the refactored perform the same function in a single loop by utilizing a set of arrays. 

The image below displays the original code with the nested loop
![NESTED LOOP](https://user-images.githubusercontent.com/109990578/185803604-435bd0a2-a4c5-442c-bb2e-bffbdb9d7928.png)

The image below displays the refactored code with a single loop
![SINGLE LOOP](https://user-images.githubusercontent.com/109990578/185803611-43da35d2-8ba1-4f56-aa07-81dfc79056f2.png)

## SUMMARY
Refactoring is a disciplined technique for restructurinng an existing body of code, altering its internal structure without changing its external behaviour. It improves internal code structurre without altering its external functionality by transforming functions and and rethinking algorithms.

The advantages of refactoring codes are:
1. TIME EFFICIENCY:
Refactored code has high efficiency due to the decrease in execution time.
2. LEGIBILITY:
Refactored code are clear enough to read and thus makes it easier for optimization by other programmers.
3. DEBUGGING:
Refactored code are easy to debug because of its comprehensive nature which facilitates maintenance and the extendiility. 
4. CONCISION:
Refactored code removes redundancies and duplications thereby improving the effectiveness of the code.

The disadvantages of refactoring code are:
1. Imprecise refactoring could introduce new bugs and errors into the code.
2. There is no clear definition of neat code.
3. An improve code is often difficult for the customer to recognize, since the functionality remain unchange.
4. In the case of larger teams working on refactoring, the coordination effort required could be surprisingly high.

### Advantages of Refactoring the Original VBA Script
Refactoring the original VBA script helps to improve time efficiency from 0.6 seconds to 0.2 seconds for the original and refactored respectively.
It also allows me to reduce reduncdancy by using single loop instead of the nested loop.

### Disadvantages of Refactoring the Original VBA Script
It was time-consuming when trying to maneauver the code to suit the refactoring instructions. 


