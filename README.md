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







## Results

- What are two conclusions you can draw about the Outcomes based on Launch Date?

  - It looks like mid year is the ideal time to launch a campaign
  - The 4th quarter has a higher fail rate relative to successful campaigns than the rest of the year

- What can you conclude about the Outcomes based on Goals? 

  - It can be concluded that the smaller the goal is the more likely it is that it will be successful
  - The number goals set correlates to a higher success rate

- What are some limitations of this dataset?

  -   One of the limitations is the outcome designated as "live." There are 51 data points that really missed because of this as some of these coule also fall into the success or failed categories.  Also, since the goal and pledged data is presented in USD, the currency column is irrelevant unless we convert to currency. Even that would cause issues with our output. I also think that Outcomes Based on Launch Date could be misleading because there is a larger set to look at for some months over others. So while we can look at springtime being ideal for launch we might also consider that it is due to the fact that this is when the most campaigns have been launched previously.

- What are some other possible tables and/or graphs that we could create?
  - We could look at the length of the campaign
  - We could also take a look at the entire data set applying the same analysis as we did for just plays or other smaller sets
