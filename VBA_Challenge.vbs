Sub AllStocksAnalysisRefactored()
    'Set variables for the starting time and ending time to be used with a timer that will measure the processing time of the VBA script
    Dim startTime As Single
    Dim endTime  As Single

    'Create input box prompting user with a message
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Initialize timer from the moment user enters the appropriate year
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'yearValue will show the value initially entered in the InputBox
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers with the number of elements inside parenthesis
    Dim tickers(12) As String
    
    'Assign numbers to each of the elements in the array to allow access each stock ticker name
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
    
    'Activate data worksheet of the year selected in the InputBox
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create a ticker Index and set it to zero before looping over the rows.  Because it's going to count in whole numbers the data type "Integer" is used
    Dim tickerIndex As Integer
        tickerIndex = 0

    'Create three output arrays that will be used alongside the tickerIndex to tie the stored value to the ticker array
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    'Create a for loop to initialize the tickerVolumes of each ticker in the tickers array to zero based on the index position
    For i = 0 To 11
        tickerVolumes(tickerIndex) = 0
        
    Next i
        
    'Loop over all the rows in the data spreadsheet to store the values for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices using the tickerIndex
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
    
        'Check if the current row is the first row with the selected tickerIndex
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
               
        'Check if the current row is the last row with the selected tickerIndex
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            'Increase the tickerIndex to begin calculating the values for the subsequent ticker in the index
            tickerIndex = tickerIndex + 1
            
       End If
                     
    Next i
    
    'Loop through arrays to output the Ticker, Total Daily Volume, and Return to the "All Stocks Analysis" worksheet
    For i = 0 To 11
        
        'Activate output worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Using the iterator (i) identified in this loop, record the values on the output worksheet
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
   
    Next i
    
    'Formatting
    'Select the output worksheet
    Worksheets("All Stocks Analysis").Activate
    'Make the header row bold
    Range("A3:C3").Font.FontStyle = "Bold"
    'Create a line beneath the header row
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'Place a comma as a thousands separator
    Range("B4:B15").NumberFormat = "#,##0"
    'Convert the return column to a percentage
    Range("C4:C15").NumberFormat = "0.0%"
    'Automatically extend the columnwidth of Total Daily Volume to show the full column name and all digits in the value
    Columns("B").AutoFit

    'For only the calculated value cells and not including the header row
    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        'Fill cells with green if return is positive
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        'Fill cells with red if return is negative
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    'Stop timer
    endTime = Timer
    
    'Display message box of run time
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


Sub ClearWorksheet()

    'Clear cells to start with a clean sheet before running the analysis
    Cells.Clear
    
End Sub

