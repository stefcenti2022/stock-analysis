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
    
    tickerCount = 12
    'Initialize list of tickers
    ReDim tickers(tickerCount) As String
   
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

    '1b) Create three output arrays
    ReDim tickerVolumes(tickerCount) As Long
    ReDim tickerStartingPrices(tickerCount) As Single
    ReDim tickerEndingPrices(tickerCount) As Single
    
    'Get ticker volumes and store them in the tickerVolumes array
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To tickerCount
        tickerVolumes(tickerIndex) = 0
    Next tickerIndex
    
    ''2b) Loop over all the rows in the spreadsheet.
    tickerIndex = 0
    For i = 2 To RowCount
            
            ticker = tickers(tickerIndex)
            
            '3a) Increase volume for current ticker
            If Cells(i, 1).Value = ticker Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
            
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
    
            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
    
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                            
        End If
        
    Next i
    
    'Format Header
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    'This will be used as an offset to update the cells with the array data and to format the color of the volume data
    dataRowStart = 4
    dataRowEnd = 15
    volumeCol = 3

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To (tickerCount - 1)
        
        Worksheets("All Stocks Analysis").Activate

        dataRow = i + dataRowStart
        
        'Update the Ticker, Total Daily Volume, and Return data
        Cells(dataRow, 1).Value = tickers(i)
        Cells(dataRow, 2).Value = tickerVolumes(i)
        Cells(dataRow, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        'Format the Volume depending on its value
        If Cells(dataRow, volumeCol) > 0 Then
            Cells(dataRow, volumeCol).Interior.Color = vbGreen
        Else
            Cells(dataRow, volumeCol).Interior.Color = vbRed
        End If
        
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub