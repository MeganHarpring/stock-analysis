# **Stock-Analysis: VBA-Challenge**

## **Overview of Project**
  The purpose of this project was to anaylze different stocks in the years 2017 and 2018 by looking at their Total Daily Volume and Yearly Returns. Analyzing these two different outcomes for the past few years will help us let the client, Steve and his parents, know which stocks would most likely be the best ones to invest in. We also want to compare how well our original code and refactored code were executed, we will look at the scripts for each code and the time they take to run to see what the pros and cons of both are.

### **Background**
  Steve, our client, wants to find the total daily volume and yearly return for each stock over the past few years to help his parents invest their money wisely. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. Steve's parents believe that if a stock is traded often, then the price will accurately reflect the value of the stock. If we sum up all of the daily volume, we'll have the yearly volume and a rough idea of how often a stock gets traded. If you invested in a stock at the beginning of the year and never sold, the yearly return is how much your investment grew or shrunk by the end of the year. Steve's parents wanted to know specifically how the DQ stock performed and Steve would like help analyzing all the stocks in 2017 and 2018 so he can give his parents an even better idea of which stock(s) to invest in. Our original code looks at specifically the DQ Stock Analysis and the 11 others included in this data, it can be ran for 2017 or 2018. Our refactored code, was then written to be more efficient and effective, espicially if we wanted to expand our anaylsis to include more stocks. Below is the analysis of how both codes worked.
  
## **Results**
#### **Analysis of Stock Performance: 2017 vs. 2018**

Below you will see the analysis for 12 different stocks in 2017 and 2018. We can analyze how the stocks performed by looking at the difference between the Total Daily Volume for both 2017 and 2018 and the Return for each stock in 2017 and 2018.

<img src="https://user-images.githubusercontent.com/100392991/159192256-aeeeda08-19b8-4cc9-be76-8a0620503441.png" width="270" height="310"> <img src="https://user-images.githubusercontent.com/100392991/159192788-74952c26-ddca-4abd-a9f3-84486731b5d4.png" width="270" height="310">

You can tell just by looking at the colors in each table that 2017 had a much better Yearly Return than in 2018. Majority of these stocks all had a negative return in 2018, whereas in 2017 only one stock had a negative return and it was farily low (-7.2%). Steve's parents were specifically interested in how the DQ stock did. The DQ stock had a Total Daily Volume of 35,796,200 in 2017 and 107,873,900 in 2018, which means it was traded quite a bit more in 2018 (a difference of 72,077,700 more in 2018). The Return went from 199.4% in 2017 to -62.6% in 2018. So your investment would of grown 199.4% in 2017 in the DQ Stock but then your investment would of shrunk by -62.6% in 2018. I don't think that Steve's parents belief that the more a stock is traded, the more the price will accurately reflect the value of the stock. The significant different between 2017 and 2018 makes me think the DQ Stock is not all that stable and may be a risky investment. 

I don't know much about stocks, investments, the relationship between Total Daily Volume and Yearly Return, but just by looking at these two tables I would stick with the stocks that stayed in the green for both years (ENPH and RUN) or atleast choose a stock that had a low negative return.

#### **Analysis of Code Performance: Original vs. Refactored**
The original code uses two loops; a nested loop to get the date from both 2017 and 2018 worksheets. The refactored code puts the two loops into one loop and references arrays to get the data. The refactored code is defiantly quicker than the original code, it also can be used more efficiently and effectively with expanded data sets. I think it would run much quicker than the original with even more stocks and data, and possibly be used in different types of data sets and work. Below, I included screen shots of the difference between the two codes and the times it took to run using the timer function. 

**Below, in the left column you will see the nested loops in the original code and in the right column you will see the single loop using the arrays:**

<img src="https://user-images.githubusercontent.com/100392991/159195908-b4826e01-d865-4634-9d3a-2cb0a25d4842.png" width="500"/> <img src="https://user-images.githubusercontent.com/100392991/159195928-e7c66eaf-414f-4a97-8e98-6d40602db2e3.png" width="500"/>
<img src="https://user-images.githubusercontent.com/100392991/159195919-976797a3-1c0e-47fa-8e41-9a8c8a27b3a5.png" width="500"/> <img src="https://user-images.githubusercontent.com/100392991/159195935-9fbd85bd-e01a-4092-aa98-cf04c71d94ef.png" width="500"/>

**Below, in the left column you see the time it took to run the original code for both 2017 and 2018, on the right you see the time it took to run the refactored code for both 2017 and 2018:**

<img src="https://user-images.githubusercontent.com/100392991/159195953-431e5ba2-12e4-4cb2-b1b4-41fe39fb9a49.png" width="245"/> <img src="https://user-images.githubusercontent.com/100392991/159195960-164497cc-86fd-4b98-871e-a8a47ddf78f4.png" width="245"/>   <img src="https://user-images.githubusercontent.com/100392991/159195977-90c1e30d-0b3a-4adc-ae29-e2b58e97e67a.png" width="245"/> <img src="https://user-images.githubusercontent.com/100392991/159195982-9173bdb7-dc73-4f89-96fb-12774cba5e16.png" width="245"/>

**As you can see the 'Refactored' code runs A LOT faster.** 

## **Summary**

#### **1) What are the advantages or disadvantages of refactoring code?**
  Some advantages to refactoring code are: more organized, better quality, could simplify or reduce complexity, easier to understand, cleaner code, easier to update, and it could save time.
  Some disadvantages to refactoring code are: it may not be noticable (functionality usually stays the same, only improves for the programmer), may take a lot of time, may introduce new problems, could end up costing more money.
  
#### **2) How do these pros and cons apply to refactoring the original VBA script?**
  Some pros to refactoring our original VBA script are: the code was executed much quicker, the code was somewhat condensed (took away the nested loop), it is nicely organized, the added comments help to understand what is going on, and could be used to process even more data.
  Some cons to refactoring our original VBA script are: it was and is somewhat confusing, and it took some time to debug a few things and test the code. 
  Overall, I like the refactored version, it seems to be helpful and it shows some new ways to use those same functions. 
  





