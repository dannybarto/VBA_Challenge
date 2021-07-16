# VBA_Challenge
The VBA of Wall Street
## Overview of Project

### Purpose
The purpose of this analysis is to use the code to analyze all stock market data over the last two years. This will be acheived by refactoring the the original solution code in Module 2. We will also analyze the performance of the code by adding timer functionality.

### Results

- 2017 was the ideal year for making investments
- ENPH was the top performer over the two years that were analyzed
- RUN provided positive returns for both years
- TERP was the big loser having lost value two years in a row
- Results for both 2017 and 2018 were retured in less than 1 second

###### Screenshots for Reference

2017_Stock_Analysis.png![2017_Stock_Analysis](https://user-images.githubusercontent.com/85522326/125882184-1ab18b05-dd28-471f-909a-c45ff673f889.png)

2018_Stock_Analysis.png![2018_Stock_Analysis](https://user-images.githubusercontent.com/85522326/125882195-3f6b595d-69e6-467b-9a27-3199912b9f60.png)

###### Full Refactored Code

    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
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
    

    '1b) Create three output arrays   
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    
        
    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        
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
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
      End Sub
   
 





## Summary

- ### Detailed statement on the advantages and disadvantages of refactoring code in general

  ### Advantages
  
  - Simpler to read and understand the code
  - Easier to maintain
  
  ### Disadvantages
  
  - Refactoring code is a very time intensive process 
  - Refactoring is not very flexible as you cannot introduce new functionality
   
- ### Detailed statement on the advantages and disadvantages of the original and refactored VBA script
  
  ### Advantages
  
  - We know that the original code already works
  - With original code there is no need to re-allocate resources within the organization
  
  ### Disadvantages
  
  - In the long term the initial investment in refactoring code might pay off due to the increased efficiency and understandability of the code
  - The refactored code is scalable whereas the original code is not. 
