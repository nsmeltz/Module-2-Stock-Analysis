Attribute VB_Name = "Module3"
Sub AllStocksAnalysis()

    Worksheets("All Stocks Analysis").Activate
    Cells.Clear 'clears all cell properties(contents,formats,comments,etc)
    
    Dim startTime As Single 'define timer start & end variables
    Dim endTime  As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer  'starts the timer
    
   '1) Format the output sheet on All Stocks Analysis worksheet
            
            Range("A1").Value = "All Stocks (" + yearValue + ")"
   
        'Create a header row
            Cells(3, 1).Value = "Ticker"
            Cells(3, 2).Value = "Total Daily Volume"
            Cells(3, 3).Value = "Starting Price"
            Cells(3, 4).Value = "Ending Price"
            Cells(3, 5).Value = "Yearly Return"

   '2) Initialize array of all tickers labels
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
        
    '3a) Initialize variables for starting price and ending price
        Dim startingPrice As Single
        Dim endingPrice As Single
        
    '3b) Activate data worksheet and sort by ascending Ticker Name
        Worksheets(yearValue).Activate
        'Range("A1:H3013").Sort Key1:=Range("A1:A3013"), Order1:=xlscending
        
    '3c) Get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).ROW

    '4) Loop through tickers
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0

            '5) loop through rows i
                Worksheets(yearValue).Activate
                
                For j = 2 To RowCount
                
                    '5a) Get total volume for current ticker
                    If Cells(j, 1).Value = ticker Then
                        totalVolume = totalVolume + Cells(j, 8).Value
                    End If
                    
                    '5b) get starting price for current ticker
                    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                        startingPrice = Cells(j, 6).Value
                    End If
            
                    '5c) get ending price for current ticker
                    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                        endingPrice = Cells(j, 6).Value
                    End If
                    
                Next j
            
            '6) Output data for current ticker
                Worksheets("All Stocks Analysis").Activate
                Cells(4 + i, 1).Value = ticker
                Cells(4 + i, 2).Value = totalVolume
                Cells(4 + i, 3).Value = startingPrice
                Cells(4 + i, 4).Value = endingPrice
                Cells(4 + i, 5).Value = (endingPrice / startingPrice) - 1
        
        Next i


'Cell Formatting
    
    Range("3:3").Font.FontStyle = "Bold"  'bold the text in row 3
    
    Range("A3:E3").Borders(xlEdgeBottom).LineStyle = xlContinuous 'line on bottom of header cells
    
    Range("B4:B15").NumberFormat = "#,##0"  'number with no decimals
    Range("C4:C15").NumberFormat = "$#,##0.00" 'currency wiht 2 decimal places
    Range("D4:D15").NumberFormat = "$#,##0.00"
    Range("E4:E15").NumberFormat = "0.0%" 'percentage 1 decimal place
    
    Columns("B").AutoFit 'AutoFit column width
    
    
' Conditional Formatting

    'define the start and end rows
        dataRowStart = 4
        dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd

        If Cells(i, 5).Value > 0 Then

            'Color the cell green
            Cells(i, 5).Interior.Color = vbGreen

        ElseIf Cells(i, 5).Value < 0 Then

            'Color the cell red
            Cells(i, 5).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 5).Interior.Color = xlNone

        End If

    Next i

endTime = Timer 'ends the timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue) ' displays message about run time

End Sub

