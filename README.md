# VBA_Challenge
The VBA of Wall Street
## Overview of Project

### Purpose
The purpose of this analysis is to use the code to analyze all stock market data over the last two years. This will be acheived by refactoring the the original solution code in Module 2. We will also analyze the performance of the code by adding timer functionality.

### Results





### Analysis of Outcomes Based on Goals
The data for outcomes based on goals tells us that smaller goals have a better chance for success. In fact, there is an inverse relationship between the goal amount and whether or not it was failed or successful. At the $15,000 to $19,999 level there is an even chance of success or failure. Between $35,000 to $49,999 is a small window where higher goals correspond to a better chance for success over failure but the number of campaigns in this window is very small compared to the lower goals. I would say the best chance for success would be to set the goal somewhere between less than $1,000 to $19,999.

![image](https://user-images.githubusercontent.com/85522326/124329136-eaedf380-db58-11eb-9ec4-5bc75cc19a71.png)

### Challenges and Difficulties Encountered

- Converting the dates was one of the first challenges I encountered
- It took me some time to understand absolute cell references which is definitely a time saver
- While I am a regular user of excel I do not really use charts or graphs in my work. It takes some getting used to in order to look at the data graphically versus just looking at the numbers

## Results

- What are two conclusions you can draw about the Outcomes based on Launch Date?

  - It looks like mid year is the ideal time to launch a campaign
  - The 4th quarter has a higher fail rate relative to successful campaigns than the rest of the year

- What can you conclude about the Outcomes based on Goals? 

  - It can be concluded that the smaller the goal is the more likely it is that it will be successful
  - The number goals set correlates to a higher success rate

- What are some limitations of this dataset?

  -   One of the limitations is the outcome designated as "live." There are 51 data points that really missed because of this as some of these coule also fall into the success or failed categories.  Also, since the goal and pledged data is presented in USD, the currency column is irrelevant unless we convert to currency. Even that would cause issues with our output. I also think that Outcomes Based on Launch Date could be misleading because there is a larger set to look at for some months over others. So while we can look at springtime being ideal for launch we might also consider that it is due to the fact that this is when the most campaigns have been launched previously.

- What are some other possible tables and/or graphs that we could create?
  - We could look at the length of the campaign
  - We could also take a look at the entire data set applying the same analysis as we did for just plays or other smaller sets
