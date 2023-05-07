VBA Challenge (stock price data) - coding references

'https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
Despite using a For each/next loop for each worksheet, I was having trouble getting it to cycle through each of the worksheeets.  After Googling 'running a macro on multiple worksheets at the same time' on Google and looking at a number of references I came across this reference which showed that I was missing one more command 'ws.select'.  This is used at the very beginning of my code.

'https://stackoverflow.com/questions/53099470/how-to-round-to-2-decimal-places-in-vba
The first iteration of my code kept rounding the yearly price changes (and percentage changes) to zero, which was causing all kinds of problems with the portion of the code that searches for the biggest gainers/losers and largest total volume traded in each year.  I figured it had something to do with how my variables were defined, but this cleared it up for me.  I ended up defining most of my variables (non-text) as double rather than integer since the data-set is so large.  This was used primarily at the beginning of the sub, but also towards the end when I add a few more variables in the last 2 for/loops.

'https://www.automateexcel.com/vba/sorting/#vba-code-to-do-a-multi-level-sort
I added a couple of lines of code to first clear any sorting on the stock data and then to perform a multi-level-sort on it - first by ticker and then by date.  This was to ensure that the dataset was displayed in ascending order both alphabetically and also by date.  Since the sub uses a conditional loop in order to determine/record the opening and closing price for each stock, it's crucial to make sure that the data is sorted properly.  This was used right after I defined variables.

https://excelchamps.com/vba/find-last-row-column-cell/
I used this code to help me set the value for my variable 'LastRow' at the beginning of my code.  I also used it later on when I defined the variable 'LastRowCounter' and needed to assign a value to it.

http://dmcritchie.mvps.org/excel/colors.htm
I used this resource to make sure I had the right colorindex definitions for my conditional formatting.  This was actually provided as a reference in class this past week.

