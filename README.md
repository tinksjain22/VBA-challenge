# VBA-challenge

VBA Script which loops through and generated the following:

## Retrieval of Data:

The script loops through one quarter of stock data and reads/ stores all of the following values from each row:

    Ticker symbol
    Volume of stock 
    Open price
    Close price


## Column Creation

Created the following column and populates the same:

    Ticker symbol
    Total stock volume
    Quarterly change
    Percent change 

## Formatting : 

Formatting is applied correctly and appropriately to the quarterly change and  percent change column using interioir color and format cells function. 

## Summary:

All three of the following values are calculated in the output by looping the conditional statements along with Max and Min function:

    Greatest % Increase
    Greatest % Decrease 
    Greatest Total Volume

## Looping Across Worksheet: 

The VBA script can run on all sheets successfully using 'For Each' loop.

## Acknowledgments

I've included a few of resources used for exploring VBA in-built functions.

[Interior Color](https://stackoverflow.com/questions/365125/how-do-i-set-the-background-color-of-excel-cells-using-vba)

[Format Function](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/formatpercent-function)

[Microsoft Advanced Filter](https://learn.microsoft.com/en-us/office/vba/api/excel.range.advancedfilter)
