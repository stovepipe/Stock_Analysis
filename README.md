# Stock Analysis Challenge with Excel VBA

## Overview of Project
    
    The purpose of this project is to refactor Microsoft Excel VBA code to increase processing efficiency. In the original AllStocksAnalysis macro, we wrote a code for the client that successfully analyzes twelve stock key figures for a given dataset and produces the analysis at the click of a button. The client is now seeking to expand this analysis to a much larger dataset that contains the entire stock market. Prior to scaling up the macro to accommodate the larger dataset, the code must be refactored.  

## Results

    The file below, "VBA_Challenge.xlsm", is where the macro was written and the analysis was performed. This file contains the original dataset and additional analysis. While this dataset contains stock market data on twelve stocks, the client was originally only interested in one stock, DAQO (Ticker: DQ). After completing the initial analysis of DAQO stock (see the DQAnalysis macro below), the code was expanded to analyze all stocks contained in the dataset (see the AllStocksAnalysis macro below). Before scaling up the dataset again, the AllStocksAnalysis macro has been refactored to increase code efficiency (see the AllStocksAnalysisRefactored macro below)

[VBA_Challenge.xlsm](https://github.com/stovepipe/stock-analysis/blob/main/VBA_Challenge.xlsm)

    DQAnalysis macro:
    ```
    Sub DQAnalysis()

        Worksheets("DQ Analysis").Activate
    
            'Name the Output worksheet Data for the specific stock
            Range("A1").Value = "DAQO (Ticker: DQ)"

            'Create a header row
            Cells(3, 1).Value = "Year"
            Cells(3, 2).Value = "Total Daily Volume"
            Cells(3, 3).Value = "Return"

        Worksheets("2018").Activate

            'Set initial volume to zero
            totalVolume = 0

            Dim startingPrice As Double
            Dim endingPrice As Double

            'Establish the number of rows to loop over
            rowStart = 2
            rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

            'loop over all the rows
            For i = rowStart To rowEnd

                If Cells(i, 1).Value = "DQ" Then

                    'increase totalVolume by the value in the current row
                    totalVolume = totalVolume + Cells(i, 8).Value

                End If

                If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

                    'Finidng the first close price of the year
                    startingPrice = Cells(i, 6).Value

                End If

                If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

                    'Finding the last close price of the year
                    endingPrice = Cells(i, 6).Value

                End If

            Next i

        'Directing the output to the correct cells
        Worksheets("DQ Analysis").Activate
            Cells(4, 1).Value = 2018
            Cells(4, 2).Value = totalVolume
            Cells(4, 3).Value = (endingPrice / startingPrice) - 1

    End Sub
    ```

    AllStocksAnalysis macro:
    ```
    Sub AllStocksAnalysis()
        'Add timer elements for measuring code performance
        Dim startTime As Single
        Dim endTime  As Single
    
        'Ask for input of which year to run code and assign to variable
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
   
            'Initialize variables for starting price and ending price
            Dim startingPrice As Single
            Dim endingPrice As Single
   
        'Activate data worksheet
        Worksheets(yearValue).Activate
   
        'Get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

        'Loop through tickers
         For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
       
        'loop through rows in the data
            Sheets(yearValue).Activate
                For j = 2 To RowCount
           
                'Get total volume for current ticker
                    If Cells(j, 1).Value = ticker Then

                        totalVolume = totalVolume + Cells(j, 8).Value

                    End If
           
                    'get starting price for current ticker
                    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                        startingPrice = Cells(j, 6).Value

                    End If

                    'get ending price for current ticker
                    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                        endingPrice = Cells(j, 6).Value

                    End If
                Next j
       
            'Output data for current ticker
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

            Next i

            'Formatting
            Worksheets("All Stocks Analysis").Activate
            Range("A3:C3").Font.FontStyle = "Bold"
            Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
            Range("B4:B15").NumberFormat = "#,##0"
            Range("C4:C15").NumberFormat = "0.0#%"
            Columns("B").AutoFit
    
            dataRowStart = 4
            dataRowEnd = 15
            For i = dataRowStart To dataRowEnd

                If Cells(i, 3) > 0 Then

                    'Color the cell green
                    Cells(i, 3).Interior.Color = vbGreen

                ElseIf Cells(i, 3) < 0 Then

                    'Color the cell red
                    Cells(i, 3).Interior.Color = vbRed

                Else

                    'Clear the cell color
                    Cells(i, 3).Interior.Color = xlNone

                End If

            Next i
    
        'end timer elements for measuring code performance
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub
    ```

    AllStocksAnalysisRefactored macro:
    ```
    Sub AllStocksAnalysisRefactored()
    
        'Variables needed for code performance measurement
        Dim startTime As Single
        Dim endTime  As Single

        'Variable from user inpout
        yearValue = InputBox("What year would you like to run the analysis on?")

            'Variable needed to start the timer for code performance measurement
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
        tickerIndex = 0

        '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
        '2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        
        Next i
        
        '2b) Loop over all the rows in the spreadsheet.
        Worksheets(yearValue).Activate
        
        For i = 2 To RowCount
    
            '3a) Increase volume for current ticker
                If Cells(i, 1).Value = tickers(tickerIndex) Then

                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
                End If
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
                If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

                   tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

                End If
        
            '3c) check if the current row is the last row with the selected ticker
            'If the next rows ticker doesnt match, increase the tickerIndex.
                If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

                End If
            '3d Increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                    tickerIndex = tickerIndex + 1
            
                End If
        
            Next i
        
            4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
            For i = 0 To 11
        
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
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
            
                    'Color the cell green
                    Cells(i, 3).Interior.Color = vbGreen
            
                Else
        
                    'Color the cell red
                    Cells(i, 3).Interior.Color = vbRed
            
                End If
        
            Next i
    
        'Variable needed to start the timer for code performance measurement
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub
    ```

### Stock Analysis
    
    The image below shows the stock analysis for each year, 2017 and 2018. The "Return" field reflects the change in price for a specific stock over the course of the year, allowing users to quickly see if a stock increased or decreased in price over a given year.

![Stock_Analysis_Comparison_Extra.png](https://github.com/stovepipe/stock-analysis/blob/main/Stock_Analysis_Comparison_Extra.png)

    From the image, we can see that 11 of the 12 stocks analyzed netted positive returns in 2017. The market flipped in 2018 with only 2 stocks continuing to see positive returns in 2018.

### Refactored Script Performance Analysis

    The original AllStocksAnalysis macro runs consistently between 1.00 and 1.1 seconds in length. After refactoring the code, the updated macro, AllStocksAnalysisRefactored, will run consistently between 0.2 and 0.3 seconds in length. From this information, we can conclude that the refactored code is producing the same output, but at a 75% faster rate. As a result of this refactoring, the client will save significant processing time and computing power once the code is scaled to accommodate a larger dataset.

## Summary

    As with any process, refactoring code has advantages and disadvantages. The pros of refactoring code ultimately point to cleaner, easier to read and quicker to process code. However, the process of refactoring can be very time consuming and create additional problems to troubleshoot as the refactoring process is completed. Nevertheless, if resources were unlimited, refactored code ideally leads to more efficient and sustainable systems.

    In regards to the original VBA script, AllStocksAnalysis, the refactored code produced the same output for the client at a 70-80% faster rate. With the clinet's goal of scalability in mind, this result will have a significant impact on the code performance for larger datasets, mainly, in minimizing processing time and computing power. While refactoring the original VBA script took additional time and code changes, the benefit of such a large decrease in run time will benefit the client and allow the code to remain flexible for future applications