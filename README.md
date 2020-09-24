# Stock Analysis

## Overview of Project
- Our client Steve is helping his parents figure out which stock is the best for them to buy. After previously doing work with Steve, he wants to expand his investigation. He no longer wants to just focus on a couple of stocks, but Steve wants us to expand the dataset to include the entire stock market over the last couple of years. 

### Purpose
- Since we have already came up with VBA code for Steve before, we need to refactor the VBA code so that it will fit to the current needs. We additionally need to make a statement on the use of refactoring. 


# Analysis for All Stocks in the Year 2017 & 2018

#### As shown below, most of the stocks had a positive outcome in 2017. Only TERP and VSLR had negative outcomes. 

![AllStock2017Challenge](https://github.com/mckenziekkilburn/stock-analysis/blob/master/AllStock2017Challenge.PNG)

#### In the image below, we can see that in the year 2018 the stocks were not as successful as the year before. Most of them fell over half of a percentage than they were in 2017. There were only two that remained in the positive spectrum. 

![AllStock2018Challenge](https://github.com/mckenziekkilburn/stock-analysis/blob/master/AllStock2018Challenge.PNG)

#### By looking at both stocks, I would suggest going after the stocks ENPH & RUN. Both of these stocks have a steady positive rate for both years. RUN's rate went up a great amount the following year (2018).


# Execution Times for Both Years

## In the image below, is the elapsed run time for the refactored code VS. the original code for the year 2017. 


![VBA_Challenge_2017](https://github.com/mckenziekkilburn/stock-analysis/blob/master/VBA_Challenge_2017.PNG) ![VBA_2017_orginalresult](https://github.com/mckenziekkilburn/stock-analysis/blob/master/VBA_2017_orginalresult.PNG)

#### ***As we can see in the images above,  the refactored code that we came up with gives us a much faster run time than the code we had originally.*** 



## Below, is the elapsed run time for the refactored code VS. the original code for the year 2018.


![VBA_Challenge_2018](https://github.com/mckenziekkilburn/stock-analysis/blob/master/VBA_Challenge_2018.PNG) ![VBA_2018_originalresult](https://github.com/mckenziekkilburn/stock-analysis/blob/master/VBA_2018_originalresult.PNG)

#### ***Simialar to the 2017 comparison, the refactored code is 3x faster than the original code.***

# Summary to our Analysis


## What are the advantages/disadvantages of refactoring our code in general? 
##### - When it comes to refactoring our code, all it simply means is that we have edited our original code. As shown in the elapsed run time images, refactoring our code is a great way to make it faster and more efficient. This is one of the biggest advantages to refactoring. It makes it easier for other users to read in the future. The only disadvantage to refactoring is we can edit to the point of losing ourselves in the code. Meaning, we already have it working so we need to make sure we do not do anything to make it stop working. 

## What are some of the advantages and disadvantages of original VBA Script and the refactored one?
##### - I would say that the VBA script originally was easier to understand, but it was alot slower than the refactored script. More or less it gets the job done, but it takes longer to run and is not as efficient. Like stated previously, refactoring code is more efficent and easier to maintain in the future. I know when I was refactoring I was getting alot of errors because there were some things I had to correct on my own. It is better to correct the errors now so that you dont pay for it later. 



