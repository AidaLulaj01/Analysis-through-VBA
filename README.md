# **Green Stock Analysis Project**
## **Overview of the Project**

The purpose of this project was to assist Steve with calcuating the total daily volume and total return for a large group of green stocks. We developed code using VBA to largely automate the process for an existing data set organized by Year, (2017 and 2018) Ticker, Daily Volume, and Price. With the click of a button, Steve may now use our code to quickly calculate the desired outputs in a clean, easy-to-read format. 

While we were able to complete the initial analysis by taking advantage of "nested Ifs" within our code, the resulting calculations were resource-intensive and took a relatively long time to determine the outputs. As such, in Round 2 of our analysis, we did our best to refactor the code by defining an index to utilize in lieu of the nest ifs, resulting in noticeable improvements to our calculation speed.

## **Results**

While the above-referenced original code did result in the desired output, we found that the use of nested ifs was detrimental to the performance of the code. Specifically, the 2017 analysis took approximately 0.85 seconds and the 2018 analysis took approximately 0.87 seconds.

## **Summary**

### **_Advantages and Disadvantages of Refactoring Code_**

The advantages of refactoring code are inherent in the above-referenced results. Refactored code may improve upon the original design, thereby improving on performance and formatting. 

There are, however, disadvantages to consider as well. There is an old addage saying that "if it isn't broken, don't fix it." While refactoring code may yield marked improvements, you are also editing existing code that works. Any time you do this, you introduce potential new bugs, syntax errors, etc. While refactored code may work for an existing data set, it may introduce problems down the road when applied to new use cases. If code is modified over time, you always run the risk of introducing flaws/errors into the code. 

Additionally, the refactored code may not be as intuitive to understand for the average user.

### **_Refactoring the VBA Script_**

As noted above, our refactored code improved the 2017 analysis from 0.85 seconds (original) to 0.17 seconds (refactored), and our 2018 analysis from 0.87 seconds (original) to 0.18 seconds. While it is difficult to appreciate this improvement tangibly (as both outputs took less than one second to generate), consider that the data set we were working with is not that large. For larger data sets, the improvements should be far more noticeable. These improvements are largely the result of our refactored analysis not having to loop over each ticker line-by-line throughout the entire data set. Rather, the refactored code relies on an index to quickly reference and cluster together the appropriate outputs. As such, this such yield great improvements over a large data set.

It can be argued, however, that the refactored code is not as intuitive to the average user as the original. While it makes sense to go go line by line and aggregate volumes and prices by ticker (this is how the average person's mind would visualize the calculation), using an index is not as easy to conceptualize and to follow.




