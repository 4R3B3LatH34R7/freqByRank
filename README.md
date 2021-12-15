# freqByRank
Excel UDF to get Frequency of value(s) in a range of cells by Rank

## Background
Yes, we can use many formulas/functions in MS Excel to get the number of times a value appears in a range of cells.\
For example, we can use FREQUENCY, MODE, MODE.SNGL, MODE.MULT, COUNTIF, COUNTIFS but there are limitations.\
FREQUENCY and MODE formulae are not working with non-numeric values in general.\
COUNTIF formulae cannot be set to return a certain rank/number of frequencies/occurences.\
There are many array or non-array formulas that work by using long nested formulas to return the N-th number of Frequencies of a value in a range of cells.\
However, it was never easy to just ask a formula/function to return exactly which value(s) appears exactly 2 times in a range.\
Even then, I have seen some several-levels-nested and very complicated formulas that can return a value with certain frequency but they failed to return all the values having the same frequency.\
The Mode.MULT function can return a values with same mode from a multi-modal dataset but it is limited to work with only numeric values.\
This UDF was made to overcome all those above shortcomings of the above builtin formulae/functions.

## How
The logic behind this UDF is very simple, in that, I take in a range of cells.\
Put them into an array and then find out the count of that value inside the range, save the count in a dictionary with the frequency as the key and then replace value with vbNullString in the range.\
Then repeat the same process with the rest of the cells in the same range if they are not vbNullStrings.\
If there are values with the same frequencies, they are appended together inside the dictionary in an array.\
My original plan was assigned all the values with numerical values with corresponding keys (to get the values back) so that I can use the MODE.MULT function on the range.\
But I was afraid that the conversion/matching processes of the whole range into numerical values might be more resource intensive so I just used a customized COUNTIF function that I can use with arrays instead of ranges as it was designed to be used.\
The tricky part is to customize the return values. The return part code was far longer than the calculation code.\
I even go extra lengths to create a ranking function that would rank the frequencies so that when the user ask for a specific rank of frequency, the UDF would return the right an array containing the exact rank.\
I didn't want to sort the arrays/dictionaries so I have create a function to translate the ranks to frequencies which are actually the keys to the processed dicitonary.

## Parameters/Arguments
