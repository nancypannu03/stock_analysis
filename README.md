# stock_analysis
# Overview of the Project: VBA Stock Analysis
## Purpose Of the Analysis
In this Project we will assist Steve and his parents to analyze an entire dataset. We will refactor the module 2 solution code to loop through all the data one time and collect the relevant information. With refactoring we will be able to analyze the given data set more efficiently and in as faster manner. In the Green Stocks worksheet we have been provided with the data of the 12 green stocks for years 2017 and 2018 respectively. Therefore, we will loop over these sheets and refactor our code to present faster conclusions.

# Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

 '1a) Create a ticker Index
 
 
![Test Image](/Resources/tickerIndex.png) <br/>
    
    

    '1b) Create three output arrays   
    
    ![Test Image](/Resources/OutputArrays.png) <br/>
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    
    ![Test Image](/Resources/2a.png) <br/>
    
        
    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount
     '3a) Increase volume for current ticker
     
     ![Test Image](/Resources/2b_3a.png) <br/>
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            

            '3d Increase the tickerIndex. 
            
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next i
    



# Summary: 
## What are the advantages or disadvantages of refactoring code?
## How do these pros and cons apply to refactoring the original VBA script?
