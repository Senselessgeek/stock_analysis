# stock_analysis

## Project Overview
  ### Why
    Steve was kind enough to give some work to a struggling data analyst.  This first contract I hope satisfies, and encourages more work from Steve in the future.
    
  ### What
     A basic analysis of stocks that Steves parents are very fond of would like to invest in.  The expected out put is to show that some stocks are not very good and others are probably best to put money towards.
     
     A secondary goal was presented by Steve after the first iteration of the report was turned over with the request to speed it up a bit.
     
  ### How
     Rather than make a sheet with formulas that do a bunch of stuff, we will be using VBA to automate alot of lookups and iterate through the whole lot of data to get to the conclusion we need to present.
     
     To hit the secondary goal we refactored our code to be a bit more inline with the expectation that parsing through the code less would lead to significant reduction of overhead.
     
     
## Results
  What we have found in our results is that while in 2017 stonks only go up, but in 2018 stonks go down.  With two exceptions in ENPH and RUN, based on our very limited data, I would suggest they dump everything in those two stonks and ride to eternity, DIAMOND HANDS!!!
  
![2017 Results](resources/2017.png) | ![2018 Results](resources/2018.png)

After reFactoring the code there was a noticable difference in the runtime speed, I absolutely could tell the difference between .3 seconds to run and 0.05 seconds to run, I totally would notice that while working with constant interuptions on a warehouse floor or being micromanaged by a manager.  Maybe when Steve's parents decide to expand their horizons or Steve decides he was wants to look at real historical data, a 83% reduction will be noticable once it takes longer than an ADHD distraction.

## Run Time Comparison

#### 2017 Before Image

![2017 Before](resources/VBA_Challenge_Original_2017.png)

#### 2017 After Image

![2017 After](resources/VBA_Challenge_2017.png)

#### 2018 Before Image

![2018 Before](resources/VBA_Challenge_Original_2018.png)

#### 2018 After Image

![2018 After](resources/VBA_Challenge_2018.png)

## Code Snippet

  Here we used Dim to make an array
  ```
  Dim tickerVolumes(12) As Long
  ```

  This one is the start of a for loop with no end in sight
  ```
  For i = 2 To RowCount
  ```
  
  Here is where I accidentally tried to divide 0 by a number because i forgot what variable i was calling
  ```
  Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
  ```
 
  
## Summary

### Pro/Con of refactoring in general
  There is cost benefit analysis to be had when looking at refactoring code.  On one hand there is generally a significant amount of time that is to be invested to add efficiencies to already made code.  You can determine this by taking an estimated savings in the refactored code.  Take some base calculations of how many calcs you are doing in the old code, vs what it might look like using psuedocode.  Will those savings net a great enough reduction in the time spent running the program vs if the program was to continue in its current form.  How long could it run like that before refactored code provided an ROI.  Sometimes not fixing and deprecation for a better alterntive is the way things go.
  
  Here Refactoring was not a major issue as the code base was small in comparison to what there could be in some instances.  We knew that Steve is looking at using larger data sets and that he will have more clients.  As this is the first iteration of the program, refactoring now would provide more value to Steve over the long term as the expected use case is indefinite.
  
### Pro/Con of Original vs Refactored
  The logic between both of the codes are very similar, however in the refactored version we went with a more inline type approach. rather than run the same code multiple times,  we ran the code once per iteration and stored it to be called back later.
  This provided a very noticable benefit in being much more scalable to a larger data set.  What I noticed while refactoring was that we will still need to be called upon when Steve wants to increase the number of tickers he compares, since those values are still stored as an absolute.  So some refactoring will be needed in the future as it wasn't a 100% hands clean creation.  For the time being Steve will likely be happy with an 83% decrease in run time.
    
