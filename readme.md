
   

# Overview of Project
This project was to analyze how 12 different stocks performed over the year 2017 and 2018 to better help Steve help his parents invest their savings.  This analysis was done by looking at the total traded volume for the year, and well as looking at the percent gain or loss for the year by determining the starting and ending prices.

## Results: 

### Stock Performance

The data shows very different performances for the stocks from the two years.  In 2017, there was only 1 stock to have a decline in end price, and that decline was only about an 8% loss.  Four of the stocks gained over 100%, and an additional two gained more than 50%.  2018 was a very different story.  Only two stocks, ENPH and RUN gained on the year, but those stocks gained 98% and 85% respectively.  

As for the premise of this challenge, to help Steve counsel his parents on what stocks to buy, I would say that even though his parents met at Dairy Queen, I would not recommend that stock.  The stock did gain just over 200% in 2017, it did lose over 61% in 2017, meaning a significant amount of the gains were wiped out over the 2 year span.  

The outcomes for each year are shown as 2017_results.PNG and 2018_results.PNG in GitHub


### Original vs Refactored Script

The refactored script ran much faster than the original script.  The original script ran in a little over 1 second for each of the years.  However, the refactored script ran in 0.2773 seconds in 2017, as shown in VBA_Challenge_2017, and 0.2969 seconds in 2018, as shown in VBA_Challenge_2018.  At first I was amazed at how much faster this code ran in, but upon further pondering, determined that it is much more efficient because of how the code interpreted the data.  In the longer run, the code went through all the data in column A for every ticker, or 12 times, and then had to individually output the results to each output cell, then go to the next ticker.  In fact, I could see this process working.  I once ran this program while I had many other programs open on my computer, consuming processing power, and I could see the delay in printing the data in the cells as the program cycled through the next index.

However, once the script was refactored, the efficiency was amazing. We did get lucky and have the data in an order so we could make it a lot faster, but by storing the data in an array, we could more easily go through the data of the first index, then move on to the next index without printing to the cells until the end.  This sped the process up because we only had one For loop to go through the entire sheet, as opposed to 12 before.


## Summary:

###        What are the advantages or disadvantages of refactoring code? 
As stated in the results section for "Original vs Refactored Script," there was a huge time advantage by running the refactored script because the subroutine only had to run through all the rows once.  There are also a few other advantages, one of which I noticed was that the code was easier to read and follow.  It was clear that you were doing the same process to each part, and just indexing which spot in the array you were focusing on.  

The only disadvantage I could really notice was that it was a little hard to initially figure out how to do the process.  I will say once it was completed it seemed more elegant and I liked the refactored code more, but it was a little more of a brain crunch to figure out how to make it happen.


###        How do these pros and cons apply to refactoring the original VBA script?
I think that the pros of refactoring outweigh the cons by a significant margin.  Yes, it is a little trickier to write and think about how to execute the code, but it is much simpler to decode and enhance once the initial structure of code is set.  From this exercise I can say that it would be my goal to always refactor code and make it as efficient as possible, unless it is a quick code that won't be used for much more that a trial.