Attribute VB_Name = "greatest_ws"
Sub greatest()
'Subroutine for determining the greatest percent increase, decrease and total volume. To be run after the stocks_ws macro because we need the summary ticker tables to supply the data

    For Each ws In Worksheets

        'Variable holding the greatest percent increase
        Dim greatest_p_increase As Double
        
        'Variable holding the greatest percent decrease
        Dim greatest_p_decrease As Double
        
        'Variable holding the greatest total volume
        Dim greatest_total_vol As Double
        
        'Creating row and column headers
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
            'Calculating the greatest percent increase using the maximum function
            greatest_p_increase = WorksheetFunction.Max(ws.Range("K:K").Value)
            
            'Pushing greatest_p_increase to the table
            ws.Range("Q2").Value = greatest_p_increase
            
            'Format number into a percentage
            ws.Range("Q2").NumberFormat = "0.00%"
            
            'Calculating the greatest percent decrease using the minimum function
            greatest_p_decrease = WorksheetFunction.Min(ws.Range("K:K").Value)
            
            'Pushing greatest_p_decrease to the table
            ws.Range("Q3").Value = greatest_p_decrease
            
            'Format number into a percentage
            ws.Range("Q3").NumberFormat = "0.00%"
            
            'Calculating the greatest total volume using the maximum function
            greatest_total_vol = WorksheetFunction.Max(ws.Range("L:L").Value)
            
            'Pushing greatest_total_vol to the table
            ws.Range("Q4").Value = greatest_total_vol
            
            'Pulling corresponding ticker names using a for loop
            For i = 2 To 3001
            
                'If the percent change in column J is the greatest increase or decrease, pull the corresponding ticker name from the same row but column I and push it into the table
                If ws.Cells(i, 11).Value = greatest_p_increase Then
                
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                ElseIf ws.Cells(i, 11).Value = greatest_p_decrease Then
                
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    
               'If the total volume in column L is the greatest, pull the ticker name from column I and push it into the table
                ElseIf ws.Cells(i, 12).Value = greatest_total_vol Then
                
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                End If
                
            Next i
            
    Next ws

    'Message box alerting to the completion of the greatest tables on each sheet
    MsgBox "Greatest calculations complete"

End Sub
