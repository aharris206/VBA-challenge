# VBA-Challenge
Before doing this project, I thought that I had `64-bit Excel` installed! I got really confused why VBA wasn't letting me declare a LongLong!!
Fortunatly though, I figured it out by reading up on VBA `Data Types` and upon research, realised that I had the `32-bit Excel` installed, rest assured, I
made sure to `Uninstall 32-bit Office`, restart my computer, `Install 64-bit Office`, restart my computer, then open Excel!
`It's a good thing` though! Now I know `for certain` that I have the correct version installed! ;D

# Code Logic
Basiclly, I went about the task by `looping` through every row and recording the `ticker`, `day open`, `day close`, and `day volume`.
`Orignally`, I had my code set up in a way that `only` set the ticker when the next row was different (since it saves the computer from having to do extra work
by setting the ticker to the same value when it is already the same) but on `reading` the `Rubric` I wanted my code to do exactly what is specified, and it is worded
that the code `reads/ stores . . . for each row` so I re-designed it to do that because I didn't want to get dinged.

I calculate the `Year Open` for a stock by declaring a Boolean firstRow to `True` before I loop. That way, my `If statement` at the end of the loop (before moving to the next) can be set up as:
`If` firstRow `Then`
  yearOpen = dayOpen
`Else`
  `'Nothing happens.`
`End If`
firstRow = `False`

Because of the way `If Statements` work, as long as the condition equates to `True` or `False`, it works! This was a programming trick I was taught years ago
since boolean values are the most simple conditions (as `True` and `False` are the only values they `can` equate to!) This is a `clean` and `simple` way to use Booleans in logical statements!

# I made sure to include a bunch of comments in my code as well (:

# The Bonus Section!
I think there may be multiple ways to go about this. As for me, I declared seperate variables `gTotalVolume` for the `Greatest Total Volume`, `gIncrease` for `Greatest % Increase` and `gDecrease` for `Greatest % Decrease` along with their own respective ticker variables. Looping through my code, as I am setting the 3 values for a given stock, I check if `totalVolume` (the volume for that stock) is greater then `gTotalVolume` (the greatest total volume.) If so, I set `gTotalVolume` to `totalVolume` and `ticker` to `gVolumeTicker` (to record the respective ticker), since the volume for this stock is `higher` then the `previous` 1st place so far. That way, every time it comes across a stock with a higher volume, that stock's ticker and volume are recorded as the new highest, until a higher one is found. `In the end`, it is set to the stock with the `greatest total volume` for that year. I do the `Greatest % Increase` and `Greatest % Decrease` much the same, just with checking if the `percentChange` is greater than the current value for `gIncrease` or lower than the current value for `gDecrease`
