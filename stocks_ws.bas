Attribute VB_Name = "stocks_ws"
Sub stocks_ws()
    
    'Enables the code to run in all the worksheets of the file using a for loop
    For Each ws In Worksheets
    
        'Variable for holding the ticker name
        Dim ticker As String
        
        'Initial variable for holding the yearly change from the opening price at the beginning of a given year to the closing price at the end of said year
        Dim yearly_change As Double
        yearly_change = 0
        
        'Variable for holding the opening day value for each ticker, will be used to calculate yearly_change and therefore, percent change
        Dim open_val As Double
        
        'Variable for holding the closing day value for each ticker, will be used to calculate yearly_change and percent change with open_val
        Dim close_val As Double
        
        'Initial variable for holding the percentage change from the opening price at the beginning of a given year to the closing price at the end of said year
        Dim percent_change As Double
        percent_change = 0
        
        'Initial variable for holding the total stock volume
        Dim total_stock_vol As Double
        total_stock_vol = 0
        
        'Row for each ticker in summary table
        Dim ticker_table_row As Integer
        ticker_table_row = 2
            
         'Creating the column headers
         ws.Range("I1") = "Ticker"
         ws.Range("J1") = "Yearly Change"
         ws.Range("K1") = "Percent Change"
         ws.Range("L1") = "Total Stock Volume"
         
         'For loop to cycle through each date for one year
         For i = 2 To 753001
         
            'Setting up the yearly and percent change calculations by storing the opening value for each ticker. Similar to the logic for checking if the tickers are different, but using i - 1 to ensure capture of the first row's data for each ticker
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Want row i and column 3 to get the <open> value for each ticker
                open_val = ws.Cells(i, 3).Value
                
           'Ended the if statement because I wanted to ensure capture of the open value and wasn't sure how to do it in the main if statement below
            End If
             
             'Need to determine if the next ticker is the same or different. Checking if the ticker is different first
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             
                 'If the above if statement is true then the ticker value is equal to the first row's ticker
                 ticker = ws.Cells(i, 1).Value
                 
                 'If the above if statement is true, we are looking at the last row of data for each ticker, therefore we can capture the close_value from this row. The closing value can be found in column 6
                 close_val = ws.Cells(i, 6).Value
                 
                 'To calculate the yearly change, I am subtracting the opening value stored in the first if statement from the closing value captured just now
                 yearly_change = close_val - open_val
                 
                 'To calculate the percent change, I am dividing the yearly change by the opening value
                 percent_change = yearly_change / open_val
                
                 'Calculating total stock volume for a given ticker
                 total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
                 
                 'Pulling in the ticker name to the summary ticker table
                 ws.Range("I" & ticker_table_row).Value = ticker
                 
                 'Pulling in total stock volume to the summary ticker table
                 ws.Range("L" & ticker_table_row).Value = total_stock_vol
                 
                 'Pulling in the yearly change to the summary ticker table
                 ws.Range("J" & ticker_table_row).Value = yearly_change
                 
                 'Pulling in percent change to the summary ticker table
                 ws.Range("K" & ticker_table_row).Value = percent_change
                 
                 'Moving to the next row in the ticker table
                 ticker_table_row = ticker_table_row + 1
                 
                 'Resetting total stock volume to zero so the number doesn't keep building
                 total_stock_vol = 0
                 
             'If the ticker following is the same
             Else
                 
                 'Calculating total stock volume for a given ticker
                 total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
                 
             End If
             
         Next i
    
     'For loop for the conditional formatting of the summary ticker table
    For i = 2 To 3001
     
         'Making the cells in column K percentages
         ws.Cells(i, 11).NumberFormat = "0.00%"
         
         'If statement stating that if the values in column J are less than 0, make them red (3 on the VBA color index) or if they are not less than 0, we assume they are greater than 0 and made green (4 on the VBA color index)
         If ws.Cells(i, 10).Value < 0 Then
         
             ws.Cells(i, 10).Interior.ColorIndex = 3
             
         Else
         
             ws.Cells(i, 10).Interior.ColorIndex = 4
             
         End If
         
     Next i

Next ws

'Pop up message box, saying that the entire macro is done running, all the ticker summary tables have been created
MsgBox "Ticker tables complete."

End Sub
