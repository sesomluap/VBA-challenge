# VBA-challenge

Module 2 Challenge

The purpose of this exercise was to synthesize raw daily stock data into a usable annual summary including how much stocks increased/decreased and in what volume, as well as highlighting the biggest annual outliers from the dataset.

The first thing I had to do was create a loop for each worksheet, so my code would repeat across each year in the dataset, as seen in our census pt 1 in-class activity

I then started with the lowest hanging fruit-- creating titles for the summary tables. We worked on this together in a virtual study group on 10/7

Next was determining the last row for each loop, also discussed in study group and pulled from census data pt 1 activity

Setting all my variables was an important step, and I got help from Microsoft Learn as well as askBCS in making sure these were set up correctly, including which variables needed to be set at 0 and which didn't

I set up the volume and Summary Table Row variables per the credit card in class bonus activity, which also provided the initial skeleton of my 2nd loop where I was able to find closing price and total volume, as well as move ticker and total volume into the summary table.

We struggled with how to find the opening value in our 10/7 study group, but after class Eric Johnson tested the solution I ended up using, where we run a separate loop for (i-1, 1) to find it. Initially I had that loop below the larger one, but after talking it over with Eric realized why the order of operations was so important here.

sidenote--I used ws.Cells() as my function rather than just Cells() on a recommendation from ChatGPT. My askBCS tutor told me it was unnecessary, but when I ran the former function my code only ran on a single sheet, and reverting to ws.Cells() enabled it to run across all sheets again.

The parts of this loop that didn't come from the credit card bonus challenge are my own, as they were fairly intuitive. I did have to look up the FormatPercent function, which I believe I found on Stack Overflow. Figuring out it only worked when applied to the summary table entry and not to the initial calculation was my own tiral and error.

I looked up a number of complex conditional formatting functions, but in the end realized this was a simple enough binary that a ForLoop to color the cells would work just fine. My final result resembles the grader in class activity. I chose not to include zeroes as positive or negative values.

I tried using the max and min functions for my next 3 loops, but it was outputting zeroes into my table. Debugging my code through ChatGPT ultimately pointed me towards the ForLoops I ended up using, which more closely resemble what we've been doing in class anyway.

That ends the hard part! The functions and syntax I use from there are all mine as by this point the finishing steps seemed rather intuitive. I experimented with putting the greatest values functions within the i loop; but it made for slow, buggy code and ran much more smoothly as I have it now.
