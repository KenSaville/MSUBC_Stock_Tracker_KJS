# VBA-Challenge

This is homework for MSU Bootcamp.  The repository contains a vbs file that can be used to analyze an excel file containing information about stock prices. 

The program works on excel files  containing the following information, with the info for each year on a separate worksheet

<ticker>	<date>	<open>	<high>	<low>	<close>	<vol>		

the program loops through each worksheet, doing the following

1.  Scans through  ticker symbols <ticker>, consolidating the same symbols into a new row
2.  For each symbol it calculates: 	Total volume, Yearly change, Percent change and outputs these in the consolidated symbol rows
3.  Highlights positive (green) and negative (red) annual stock changes
4.  Creates a short summary table showing the stocks with highest volume, and highest percent increase and highest percent decrease each year.

two excel sheets are included

1. A shoerter 'tester file' called alphabetical tester
and
2. a longer file (called mutiple year)

used to test the program.

Each of these files had one or two stocks that had zeros for all or most of the categories

When calculating percent change, this cause a "divide by zero error"

I made several attempts to come up with if then conditions to ignore these rows

However, this was complicated becaue we needed to find the beginning and end of each block of ticker symbols and use the first entry of that block for the
open price and the last entry as close price.  SO we needed to include the zero rows in the search to find the breakpoints of the preceding and following sticks.

To get around this I first converted all of the zeros to ones

then using the open and close price for these stocks resulted in a 0% change (which is accurate)

I then reset the ones to zeros at the end of the script.

This worked for these two files, but could be complicated if only the open price was zero.
