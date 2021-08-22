# Overview of Project

## Purpose of Analysis

The purpose of this Analysis was to refactor or re-structure the existing 
code in order to improve the efficiency of retrieving the stock data that was previously worked on.
However, this time the stock data for the years of 2017 and 2018 wil be specifically worked on.

## Data results

The final data is shown using two Excel tables that shows stock data for 2017 and 2018.
Additionally, The retrieval time of said data is also shown in the PNG file to demostrated how
fast the data was gathered. The ticker, total daily volume and the return on each stock represent
the headers of the data table and it clearly shows what percentage of return each ticker had.
I started to refactor the code by creating a ticker index that was equal to zero, then created 
the three necessary arrays, loop the data through the ticker volumes, ticker starting prices and 
tickers ending prices. Next, I increased the volume of the ticker and the ticker index itself and finally
I once again looped the arrays to create an output for the ticker, total daily volume and return. 


 '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    Dim TickerVolumes(12) As Long
    Dim TickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
      ' If the next row’s ticker doesn’t match, increase the tickerIndex.
       For i = 0 To 11
            TickerVolumes(i) = 0
            TickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        TickerVolumes(tickerIndex) = TickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            TickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
          End If

            '3d Increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4)Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = TickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / TickerStartingPrices(i) - 1

##Summary 

### Pros and cons of refactoring an existing code

The more obvious advantages of refactoring a code involves a more readable code, it becomes easier to understand 
for other users, data retrieval speed is improved and debugging becomes easier. However, trying to refactor an
 existing code can lead to both wasting too much time on it getting confused trying to re-structure it. This
is why it is important to see if it's even worth it in a long run.

### The result of refactoring the stock Analysis code

The speed of which the stock analysis sub routine performed was improved. The macro run time took less time that before and 
That alone was a great improvement over using the old code.

### PNG Files

![BVA_challenge2017]https://github.com/wilsoncr93/Stock-analysis/blob/main/resources/vba_challenge_2017.PNG

![BVA_challenge2018]https://github.com/wilsoncr93/Stock-analysis/blob/main/resources/vba_challenge_2018.PNG
