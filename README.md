# stock_analysis
# Overview of the Project: VBA Stock Analysis
## Purpose Of the Analysis
In this Project we will assist Steve and his parents to analyze an entire dataset. We will refactor the module 2 solution code to loop through all the data one time and collect the relevant information. With refactoring we will be able to analyze the given data set more efficiently and in as faster manner. In the Green Stocks worksheet we have been provided with the data of the 12 green stocks for years 2017 and 2018 respectively. Therefore, we will loop over these sheets and refactor our code to present faster conclusions.

# Refactored Code: 

 '1a) Create a ticker Index
 
 
![Test Image](/Resources/tickerIndex.png) <br/>
    
    

  '1b) Create three output arrays   
    
  ![Test Image](/Resources/OutputArrays.png) <br/>
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    
    'Initialize ticker volumes to zero
    
     For i = 0 To 11
     tickerVolumes(i) = 0

      Next i
 
      ''2b) Loop over all the rows in the spreadsheet and  
      ' loop over all the rows

       For i = 2 To RowCount
 
       '3a) Increase volume for current ticker
       'Increase volume for current ticker
   
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
    
         If Cells(i - 1, 2).Value <> tickers(tickerIndex) Then
          tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
    
    '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
    
    If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
      '3d Increase the tickerIndex. 
        tickerIndex = tickerIndex + 1
        
    End If

     Next i

     '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
      For i = 0 To 11
  
    Worksheets("All Stocks Analysis").Activate
    tickerIndex = i
    Cells(i + 4, 1).Value = tickers(tickerIndex)
    Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
    Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
    
    Next i
    
## Results 
### Total Daily Volumn and Return for the each year
![Test Image](/Resources/VBA_Challenge_2017.png) <br/>

![Test Image](/Resources/VBA_Challenge_2018.png) <br/>

### Runtime of Original code
  ![Test Image](/Resources/VBA_Challenge_2017_time.png) <br/>
  
  ![Test Image](/Resources/VBA_Challenge_2018_time.png) <br/>
  

### Runtime of refactored code

  ![Test Image](/Resources/VBA_ChallengeRefactored_2017_time1.png) <br/>
  
  ![Test Image](/Resources/VBA_ChallengeRefactored_2018_time1.png) <br/>


# Summary: 
## What are the advantages or disadvantages of refactoring code?
### Advantages
- Makes a code more systematic and structured. <br/>
- Easy to read and Interpret. <br/> 
- Quicker to execute. <br/>
- Basically its more beneficial on the client end as any client will be able to easily analyze it and helps in efficient decision making.<br/>

### Disadvantages
-

## How do these pros and cons apply to refactoring the original VBA script?
