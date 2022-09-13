# stock-analysis

##Overview Of the Project

###Background and Purpose
My friend, Steve, is helping his parents to find profitable stocks in the green energy field to invest their money in. In order to know this, he was interested in finding out the Total Daily Volume and Yearly Return of the stock.  His parents were interested in DAQ stock and in the beginning, I wrote a code to analyze that. However, that stock shrunk by a lot in 2018, and therefore, steve decided to find some other green energy stocks for his parents to invest in. He provided me with an Excel file with the information on green stocks and asked me to analyze the profitability of them as well as provide him with an automated tool to help him visualize how well a specific stock did within a specific year so he could help his parents invest their money in the right stocks.

###Purpose
After finding out DAQ stock shrunk noticeably in 2018, Steve wants me to run the analysis for more green energy stocks and write a code for him that can analyze different green energy stocks. Since this is a big set of data and the analysis needs to be performed over and over I decided to use VBA to run the analysis which is perfect for codes that need to be repeated over and over (for each stock in this case). After writing the initial code in order to run the analysis faster and more efficiently I refactored my code and reduced its run time significantly.


##Results

Based on my analysis of the green stocks for the years 2017 and 2018, the only two stocks that were profitable in both years are "ENPH" and "RUN". ENPH had a 136.3% return in 2017 and 97.9% return in 2018. RUN also had a 9.5% and 85.2% return in 2017 and 2018 respectively.

After writing the code that works for different green stocks rather than only one, I decided to refactor the code so that it can run faster and more efficiently. Refactoring my code reduced the analysis run time significantly for each year. For 2017 it reduced the run time 0.758 second and for 2018 it reduced it by 0.7188 second which can be viewed in the images below. I did this by creating arrays to hold elements and therefore avoiding changing between different worksheets and lowering the steps. I made the code more readable by adding comments.


![](/Resources/VBA_Challenge_2017.png)

![](/Resources/VBA_Challenge_2018.png)


##Summary

1) Refactoring a code is very useful, especially for large datasets and can make the code more efficient. I believe it's really necessary and makes the code more readable and less complicated. Although, there is always a risk that the new code doesn't work, and debugging can be time-consuming.

2) Based on my personal experience, refactoring the original VBA script will reduce the time that a code runs significantly. This will improve the performance of the computer and saves time. However, for me, it was really challenging and time-consuming to refactor the code, as I faced a lot of errors and had to go back and debug the code numerous times. In my opinion, although the refactored code runs faster, the amount of time that I spent debugging the refactored code was much higher than the time that refactored code saved. For very complicated codes it's definitely worth it to refactor the code and a necessary skill to learn for everyone in the programming world.
