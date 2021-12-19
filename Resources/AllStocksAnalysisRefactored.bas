Attribute VB_Name = "Module6"
Sub AllStocksAnalysisRefactored()
    
    Worksheets("All Stocks Analysis").Activate
    Cells.Clear 'clears all cell properties(contents,formats,comments,etc)
    
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
    RowCount = Cells(Rows.Count, "A").End(xlUp).ROW
    
    '1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
        
        For i = 0 To 11 ' loop over the tickers in the array
        
            ticker = tickers(i) 'ticker string value
            tickerVolumes(tickerIndex) = 0 'initialize the tickerVolumes to zero

                '2b) Loop over all the rows in the year spreadsheet.
                For j = 2 To RowCount ' loops over all rows in year spreadsheet
                    
                    '3a) Increase volume for current ticker index
                    
                        If Cells(j, 1).Value = ticker Then
                            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                        End If
                    
                    '3b) Check if the current row is the first row with the selected tickerIndex.
                    
                        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                        End If
                        
                    '3c) check if the current row is the last row with the selected ticker
                    
                        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                            'Exit For 'ends the inner for loop when the if condition is met
                        End If
                    
                Next j
                
                Debug.Print "tickerIndex = "; tickerIndex
                Debug.Print "ticker = "; ticker
                Debug.Print "Daily Total Volume = "; tickerVolumes(tickerIndex)
                Debug.Print "Return = "; ((tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex)) - 1) * 100
                
            tickerIndex = tickerIndex + 1 '3d Increase the tickerIndex.
    
        Next i
    
   '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11 'ticker index
            Worksheets("All Stocks Analysis").Activate ' write to the all stocks analysis worksheet

            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
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

    endTime = Timer 'ends the timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue) 'displays a message box that tells you what the time was

End Sub
