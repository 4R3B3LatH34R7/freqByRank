# freqByRank
Excel UDF to get Frequency of value(s) in a range of cells by Rank
![freqByRank_HeatMap_Demo](/images/freqByRank_HeatMap_Demo.gif)
The .gif above was made using Office365 suite thereby eliminating any need to enter the result range prior to entering the formula, in other words, selecting and entering formula into H2 only is enough because of Excel365's Spill feature.\
<b>NB:Not so easy in older Excel versions.</b> 

## 1.Background
Yes, we can use many formulas/functions in MS Excel to get the number of times a value appears in a range of cells.\
For example, we can use FREQUENCY, MODE, MODE.SNGL, MODE.MULT, COUNTIF, COUNTIFS but there are limitations.\
FREQUENCY and MODE formulae are not working with non-numeric values in general.\
COUNTIF formulae cannot be set to return a certain rank/number of frequencies/occurences.\
There are many array or non-array formulas that work by using long nested formulas to return the N-th number of Frequencies of a value in a range of cells.\
However, it was never easy to just ask a formula/function to return exactly which value(s) appears exactly 2 times in a range.\
Even then, I have seen some several-levels-nested and very complicated formulas that can return a value with certain frequency but they failed to return all the values having the same frequency.\
The Mode.MULT function can return a values with same mode from a multi-modal dataset but it is limited to work with only numeric values.\
This UDF was made to overcome all those above shortcomings of the above builtin formulae/functions.

## 2.How
The logic behind this UDF is very simple, in that, the UDF takes in a range of cells first.\
Put them into an array and then find out the count of that value inside the range, save the count in a dictionary with the frequency as the key and then replace value with vbNullString in the range.\
Then repeat the same process with the rest of the cells in the same range if they are not vbNullStrings.\
If there are values with the same frequencies, they are appended together inside the dictionary in an array.\
My original plan was assigned all the values with numerical values with corresponding keys (to get the values back) so that I can use the MODE.MULT function on the range.\
But I was afraid that the conversion/matching processes of the whole range into numerical values might be more resource intensive so I just used a customized COUNTIF function that I can use with arrays instead of ranges as it was designed to be used.\
The tricky part is to customize the return values. The return part code was far longer than the calculation code.\
I even go extra lengths to create a ranking function that would rank the frequencies so that when the user ask for a specific rank of frequency, the UDF would return the right an array containing the exact rank.\
I didn't want to sort the arrays/dictionaries so I have create a function to translate the ranks to frequencies which are actually the keys to the processed dicitonary.

## 3.Parameters/Arguments
### 3.1.target Range - required
A single cell is selectable but there is no point in doing so.\
While the ceiling is not set, since this UDF utilizes Dictionaries in VBA, there could be an issue with memory usage if a very large range were selected.\
So, I'd be a little bit cautious not to use this UDF on large or multiple cell ranges.

### 3.2.Rank - optional - default=1
The UDF could be called minimally as =freqByRank(B2:F16) and it will return the value(s) within the said range which has the highest occurences(frequencies).\
Above is true for uni-modal datasets but for multi-modal data sets, it should be called as an array formula with Ctrl+Alt+Enter resulting in something like {=freqByRank(B2:F16)}.\
So the question is how do we know if the dataset is uni/multi-modal?\
Simple, we don't!, untill we actually call the UDF.\
On Office365, which has spill feature, there is no need for Ctrl+Alt+Enter as entering normally like =freqByRank(B2:F16) alone is enough and it will return an array to the right of the formula cell if there are more than one cell with the same frequency, automagically.\
However, on older Excel versions, users can still select a number of columns to the RIGHT of the formula cell and enter like =freqByRank(B2:F16,2) and enter as an array formula to see how many non-#N/A values are returned to check how many cells to select to be included into the array. The number 2 in this example asks the UDF to return the value(s) with the second highest frequency.\
The second argument represents the Rank in that the users can ask the UDF to return the value(s) with a specific N-th ranked frequency of occurences with Rank=1 being the highest.\
1 was actually set as the default rank so that calling =freqByRank(B2:F16) will only return the value(s) with maximum occurences within the dataset/range.\
The important takeaway here, is that, if a particular rank was set to return, the UDF shall return it (an array if there are more than 1 values with same frequencies) horizontally.\
On Office365, the right side of the formula cell should be clear and if not, it would cause a #Spill error.\
In older Excel versions, the UDF will only show the left/uppermost value only, even if there are more than 1 values with the same frequency.\
If the requested Rank value is higher than the available/possible rank values, the UDF shall return a 0. For example, if the result contains only 4 ranges of frequencies like 1, 2, 3 & 4 and if the user asked for a rank 6, like =freqByRank(B2:F16,6), it will return a 0.\

If the user wishes to have the UDF return frequencies for all the cells, the UDF could be called as =freqByRank(B2:F16,0) as an ARRAY formula but in this example 96 rows in the column containing the formula cell because B2:F16 hold 16x6=96.\
Since returning results as a 1 dimensional array was set to be the default, the UDF could be called like above.\

### 3.3.returnCount - optional - default=FALSE
This argument determines whether to return the values with set Rank of frequencies as an array of values or the count of the values in that array.\
For example, let's suppose calling the UDF like =freqByRank(B2:F16,4) it will return a very long horizontal array of values with frequencies ranked as 4 because these are all unique values with a frequency of 1 and we called with a default value of FALSE for return count, we don't know how many columns to select as the output array, resulting in either multiple #N/As or an incomplete answer.\
The above situation is particularly true for Excel versions prior to Office365. On Excel365, the resultant array will just spill over to the right.\
In the prior scenario, the third parameter, returnCount comes to the rescue by returning only the count instead of that 49-ish column single-row array by calling the UDF as =freqByRank(B2:F16,4,TRUE).

<b>NB: <i>To prevent an array of numbers if the UDF were called with returnCount=TRUE, the UDF is now limited to return the count only if it is not called as an array formula.</i></b>

### 3.4.return1D - optional - default=TRUE
The actual switch/argument for calling UDF as =freqByRank(B2:F16,0) and get the result in a single column, is the 4th argument and it could be turned on/off as =freqByRank(B2:F16,0,,TRUE/FALSE).\
If left out like =freqByRank(B2:F16,0) or =freqByRank(B2:F16,0,,) will yield the same result as return1D was set to TRUE by default.\
Usually, if the range is large and contain many unique values, there will be many lowest ranked frequencies like 1. In this case, the number of rows in the single column 1D array will be pretty large for example, 50-ish based on the sample test dataset containing a 75 cells range. However, this could become 75 rows column based on the dataset which is the reason there is a switch to not show values with frequency 1 a.k.a unique values.\

Another possible option for this argument is to turn it to FALSE resulting in a result array exactly the same in dimensions as the target array which the users can use to compare/map to the original target cell side-by-side.\
A great thing about this feature is that now users are able to see the realtime changes made in the original dataset reflected in the extrapolated heatmap-ish region. Please refer to the following image for further references.
![freqByRank](/images/freqByRank_ConditionalFormat.png)
Please note that H2:L16 must be selected while the formula in H2 was entered as an array formula with Ctrl+Shift+Enter.
<b>NB: The same effect can be obtained with =COUNTIF($B$2:$F$16,H2) and drag right for 4 more columns and drag down 14 more rows and apply conditional formatting.</b>

### 3.5.return1s - optional - default=FALSE
The 5th and last argument (as of 15DEC2021) is whether to return values which appear just once, in other words, with frequencies=1 or again, in other words, unique values.\
These values usually have the lowest rank but higher in count, for e.g., the results for a dataset which has 4 levels of occurences of values, the Rank 1 values have highest frequencies like 5 or 6, let's suppose, the unique values are usually hightest occurences with a Rank of 1 and yes, of course, this is not necessarily always true.\
In such cases where the unique values are most abundant, we can ask the UDF not to return the unique values with a switch like =freqByRank(B2:F16,0,,,TRUE) with a TRUE for the fifth argument and this is for include1s and set to FALSE by default.\
Since the default is set not to return the unique numbers, the users no longer have to select like 75 cells in a column to get the output array to fit, however, the users will probably still have to select columns to the right if ranks other than 0 were set to return but this process should be easier if a prior call were made with =freqByRank(B2:F16,4,TRUE) which would return the count of possible results or on an Office365 system, this would no longer be necessary.

### 3.6.Unintended side effect
The ability of the UDF to return frequencies based on users' choice of ranks, we can set it like =freqByRank(A1:F16,1) to get the UNIQUE values within a dataset.

### 3.7.Combining with other formulae/functions
After the users used freqByRank with return1D set to TRUE (which is default) with the formula cell in H2, they can use =COUNTIF($B$2:$F$16,H2) will return the count of values in that column.\
I'm sure there are other ways to combine this UDF with Excel's default formulae like MATCH and use it with this UDF to check for a cell's frequency value.\
So, happy explorations!

## 4.Releases
There will be 3 types of releases:
1. VBA code for UDF that could be copied from here
2. .bas module
3. .xlsm file

## 5.The Future
This is a proof-of-concept tool that I developed for my own use that was shared to the public.\
Though I tried my best to test the code as much as possible, there might hitherto yet unforeseen bugs and errors might still exist.\
Therefore, the users are responsible for their own usage of my code if they decided to use the code thus shared and it is understood that by sharing this as an open-source code, I shall not be held liable to any mishaps stemming from using the code I shared.\
However, the users are welcome to share their opinion and report bugs using the Discusssions and can also send emails to me at the emailaddress that was shared on my GitHub profile.
