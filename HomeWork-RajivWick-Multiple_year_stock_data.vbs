Sub Stocks_Summary()

'---------------------
'Variables
'---------------------
' Ticker Variable
Dim Ticker As String

'Open Price Variable
Dim OpenPriceYear As Double

'Close Price Variable
Dim ClosePriceYear As Double

'Price difference between start and end
Dim PriceDeltaYear As Double

'Percentage value of price difference
Dim PercentageChangePrice As Double

'The Total overall volume of a individual stock
'Type LongLong to allow for large numbers
Dim TotalStockVol As LongLong

'Row cursor for our summary table output
Dim SummaryTableRow As Integer

'Storing the total amount of rows in worksheet
'Type Long used to allow for large number of data
Dim LastRow As Long


'---------------------
'Bonus
'---------------------

'Storing highest percentage from summary table
Dim HighestPer As Double
'Storing lowest percentage from summary table
Dim LowestPer As Double
'Storing highest volume from summary table
Dim HighestVol As LongLong


'---------------------
'For Loops
'---------------------
Dim i As Long
Dim j As Integer





    '---------------------
    'For Each WorkSheet
    '---------------------
    For Each ws In Worksheets
    
    'Store total number of rows inside the worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set the summary table row curse to 2, allowing it to start at cell row position 2 (Position 1 cotains all the titles)
    SummaryTableRow = 2
    
    '---------------------
    'Creating our Summary Table Titles
    '---------------------
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    '---------------------
    'Creating our Bonus Table Titles
    '---------------------
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Intializing our TotalStockVol variable as 0
    TotalStockVol = 0
    
    
        '---------------------
        'Summary Table - Start to search the data rows from row 2 until the last row
        '---------------------
        For i = 2 To LastRow
        
            'Test the ticker value in the row we are comparing is not the same as the ticker value of the row below, we have reached a new stock, now do this:
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Added extra condition incase of stock entries that contain only a single data row
                If TotalStockVol = 0 Then
            
                OpenPriceYear = ws.Cells(i, 3).Value
                                                   
                End If


            'Assign the Ticker value found in the row we are currently in to our Ticker variable
            Ticker = ws.Cells(i, 1).Value
            
            'As we have reached the last row of the particular stock, and we know the data is in chronological date order we can assuming the last row of a stock contains the final close price in that year.
            'Assign the Stock Close value in the current row to our ClosePriceYear Variable
            ClosePriceYear = ws.Cells(i, 6).Value
            
            'Calculate the Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
            PriceDeltaYear = ClosePriceYear - OpenPriceYear
            
            'Output values to summary table
            ws.Cells(SummaryTableRow, 9).Value = Ticker
            ws.Cells(SummaryTableRow, 10).Value = PriceDeltaYear
                     
              
                '---------------------
                'Summary Table Visual Formatting - Green/Red - PriceDelaYear
                '---------------------
                
                'When the value is positive then the cell will be filled in green
                If PriceDeltaYear > 0 Then
                
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                
                Else
                
                'Otherwise the value is negative and the cell will be filled in red
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
                
                End If
                
            'Calculate and output the percentage of the difference from the stock value at the start and end of the year within data set
            ws.Cells(SummaryTableRow, 11).Value = (ClosePriceYear / OpenPriceYear) - 1
            'Output the total amount volume
            ws.Cells(SummaryTableRow, 12).Value = TotalStockVol      
            '---------------------
            'Summary Table Visual Formatting - % on price difference
            '---------------------
            ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
           
                        
            '---------------------
            'Now that we have come to the final data entry of a particular stock, we will reset counter variables
            '---------------------
            TotalStockVol = 0
            
            '---------------------
            'We will increase the Summary Row counter by 1 to allow a new row of outputs within the summary table as we output new stock data
            '---------------------
            SummaryTableRow = SummaryTableRow + 1
        
                    
            '---------------------
            ' If the ticker value we are comparing in the current row is the same as the ticker value in the row below, we want to do the following:
            '---------------------
            Else
                
                'This condition will only be met once we come into a new stock noted by the TotalStockVol resetting to "O"
                'We know due to the chronological order of the data set, the first stock entry will be earliest entry of the year, therefor we can store the openprice value from this row
                If TotalStockVol = 0 Then
            
                OpenPriceYear = ws.Cells(i, 3).Value
                                                   
                End If
            
                                                               
            'Add the volume found in the row to a the total volume variable
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
            
            End If
                     
                
        Next i
        
        
        
        '---------------------
        'Bonus Table
        '---------------------
        
        'Set our values to "0" before we enter our For Loop
        HighestPer = 0
        LowestPer = 0
        HighestVol = 0
       
        'We are searching through the newly created summary table, we know the No. of rows of the summary table as we used a counter/curser to create it.
        For j = 2 To SummaryTableRow
        
        
            'We are comparing Percentage value of the cell against our stored HighestPer Variable
            'If the cell value is higher than the stored variable value then do this
            If ws.Cells(j, 11).Value > HighestPer Then
                
                'New highest value equals current row
                HighestPer = ws.Cells(j, 11).Value
                'Output the data - Ticker Value from current Row
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
                'Output the data - Highest Percentage Value from Row
                ws.Cells(2, 17).Value = HighestPer
                'Apply percentage formatting to the cell
                ws.Cells(2, 17).NumberFormat = "0.00%"
       
            End If
            
            'We are comparing Percentage value of the cell against our stored LowestPer Variable
            'If the cell value is lower than the stored variable value then do this
            If ws.Cells(j, 11).Value < LowestPer Then
                
                'New Lowest value equals current row
                LowestPer = ws.Cells(j, 11).Value
                'Output the data - Ticker Value from current Row
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
                'Output the data - Lowest Percentage Value from Row
                ws.Cells(3, 17).Value = LowestPer
                'Apply percentage formatting to the cell
                ws.Cells(3, 17).NumberFormat = "0.00%"
                
            End If
            
            
            'We are comparing Volume value of the cell against our stored HighestVol Variable
            'If the cell value is higher than the stored variable value then do this
            If ws.Cells(j, 12).Value > HighestVol Then
        
                'New Higest value equals current row
                HighestVol = ws.Cells(j, 12).Value
                'Output the data - Ticker Value from current Row
                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
                'Output the data - Highest Volume Value from Row
                ws.Cells(4, 17).Value = HighestVol
            End If
        
        Next j
        
       'Summary Table Visual Formatting  - Autofit Cells
        ws.Columns("I:Q").AutoFit
    
    Next ws
        
End Sub


