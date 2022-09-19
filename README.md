# VBA of Wall Street
## Overview of Project
The client was given an Excel workbook that came with VBA scripted macros that analyzed all the stock entries compiled in a sheet for 2017 and 2018, depending on which year the client inputs when prompted. The macros then printed out the tickers, their corresponding total volume for the designated year, and their corresponding return at end of that year. Each sheet contains a few thousand entries grouped into twelve stock tickers, and the macros are able to analyze all the entries in little more than one second.

### Purpose
The challenge is to refactor the code script to make the macros more efficient and take less time to complete, especially when the macros are going to be run on longer lists with many more stock tickers. The initial code as-is would likely need a sizable amount of time when the entries number in many more thousands up to almost one million, especially when they come with dozens of stock tickers.

A successful refactor should allow the macros to run much faster without compromising the accuracy of the results; the output should be the same between the original and the refactored macros. The script should still be easy to comprehend with clear comments.

## Results
After a refactored script was created separate from the original script and then tested to work correctly, each one was run for the same year to verify that the output matched. This was done for both the 2017 year and the 2018 year, and the matching outputs are posted below.

![This is the stock performance output for the year 2017.](https://github.com/Owen-Wang1234/stock-analysis/blob/main/Resources/Stock_Performance_2017.png)

![This is the stock performance output for the year 2018.](https://github.com/Owen-Wang1234/stock-analysis/blob/main/Resources/Stock_Performance_2018.png)

In order to measure the difference in run time, each script was run a number of times with the timer records collected.

### Stock Performance from 2017 to 2018
The 2017 analysis gives the appearance that almost all of the 12 companies have performed reasonably well with only TERP being in the negative (return of -7.2%). The best performer in that year was DQ (return of 199.4%) with SEDG not far behind (return of 184.5%); ENPH and FLSR also posted returns above 100% (at least doubling any money that was intially invested at the start of the year). The others also showed some positive return except TERP.

The 2018 analysis however displays a completely different story. The only two which posted positive returns were ENPH (81.9%) and RUN (84.0%); and DQ in fact happened to be the worst performer (return of -62.6%). Unless the client is willing to take a chance on the two positive performers for 2019, it would be more prudent to diversify with other sectors that show a more consistent positive return.

### The Original Code and Performance
The "main engine" of the macro program generally looked like this:

```
    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '2) Initialize the list of all the stock tickers
    Dim tickers(11) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    '3a) Initialize the starting price and the ending price
    Dim startingPrice As Double
    Dim endingPrice As Double

    '3b) Go to the correct data worksheet
    Worksheets(yearValue).Activate

    '3c) Set the starting and ending rows for the loop
    rowStart = 2

    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where data exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    '4) The outer loop goes through all the tickers
    For i = 0 To 11
    
        ticker = tickers(i)
    
        'Initialize the total volume and reset to 0
        totalVolume = 0
    
        '5) The inner loop does the actual analysis through the data sheet
        Worksheets(yearValue).Activate
        For j = rowStart To rowEnd
        
            '5a) Accumulate totalVolume for the current ticker
            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
        
            '5b) Set the starting price of the current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                startingPrice = Cells(j, 6).Value
            
            End If
        
            '5c) Set the ending price of the current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
            
            End If
        
        Next j
    
        '6) Print out the results
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = ticker
        
        Cells(4 + i, 2).Value = totalVolume
        
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
        
    Next i
```

The stock ticker array, the total volume, and the starting and ending prices were initialized. Next, the entire list of entries was examined; if the ticker matches then the total volume is incremented by the recorded amount and the starting price and ending price of the stock are both found by checking if the entry is first and last respectively. Then, the stock ticker, total volume over the year, and annual return (based on the ratio of ending over starting) were all printed out to their designated cells. The variables are reset as the whole list gets re-examined for the next stock ticker in the array.

A sample collection of the run times of the original code is displayed here:

![These are six readings from the original macro time measurement](https://github.com/Owen-Wang1234/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

With this admittedly small sample size of six, some basic observations are:

- MEAN:1.16081
- MEDIAN:1.1582
- MIN:1.121094
- MAX:1.203125

Although this appears reasonably far faster than analyzing over 3,000 stock data entry lines for twelve companies by hand, the run time will grow at a much faster rate if the data sheets were much longer and contained many more thousands if not a million data entry lines for several dozens of companies if not a thousand.

### The Refactored Code and Performance
The revised "main engine" after refactoring now looks like this:

```
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        tickerVolumes(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        'Output the Ticker
        Cells(4 + i, 1).Value = tickers(i)
        
        'Output the Total Daily Volume
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'Output the Return
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i
```

The total volume, the starting price, and the ending price are now declared as arrays like the stock ticker list; this means a quick loop is used to initialize the total volumes at 0. Now, as the entire list is examined, when the stock ticker in the entry matches the current ticker in the list, the volume, the starting price, and the ending price are all entered into the correct index in the respective arrays that correspond to the ticker. When the stock ticker in the entry does not match the current ticker in the list, the next ticker in the list is used. The entire list is gone through just once as opposed to every time for each stock ticker like before, and the output is now printed outside the end of loop as a result. Another loop is set up to print the contents of the output arrays.

A sample collection of the run times of the refactored code is displayed here:

![These are six readings from the refactored macro time measurement](https://github.com/Owen-Wang1234/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Refactored.png)

With this admittedly small sample size of six, some basic observations are:

- MEAN:0.181641
- MEDIAN:0.173828
- MIN:0.1640625
- MAX:0.2109375

Without compromising the intended function, the program was able to do the same job at a much faster pace, taking almost 15% of the time initially required with the original code.

## Summary

1. What are the advantages or disadvantages of refactoring code?
   - Refactoring a program generally involves adjusting the code for improvement without affecting the function. The concept of improvement includes but is not limited to faster performance, better readability, and simpler code. Although the relative difficulty of such a task varies, a successful refactoring means the program will run faster and may be less prone to issues. However, refactoring tends to "optimize" the code in a certain way depending on what form of improvement is desired, and refactoring towards one thing may be at the expense of another. As an example, one primary concern is that refactoring programs for speed might make them dependent on how the input and data are set up and configured, so they may no longer be as robust.
2. How do those pros and cons apply to refactoring the original VBA script?
   - When refactoring VBA macros, doing so to increase performance speed to go through large data sheets with many data entries quickly runs the risk of making the macro less robust as it depends even more on the configuration and arrangement of the data. As an example, the macro involved in this project may be much faster after refactoring to handle much larger volumes of data in much less time, but it is now less robust than before. The data must be grouped by ticker for the refactored macro to work properly; the original macro was not as dependent on this.
