Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()

    'Declare variables
    Dim total_volume As Double  'running total volume for ticker symbols
    Dim row As Long             'used in For loop to iterate through all rows in worksheet from 2 to last row
    Dim summary_row As Long     'used as a counter to keep track of row on in summary table
    Dim last_row As Long        'used to find last row in worksheet
    Dim ws As Worksheet
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim starting_row As Double
    

    
    For Each ws In Worksheets
   'Set values for each worksheet
       
   
   'Initialize variables
    'Row is initialized in the For loop
    summary_row = 2     'summary_row starts at 2 to not overwrite header
    total_volume = 0    'starts at 0 by default
    open_price = 0
    close_price = 0
    yearly_change = 0
    percent_change = 0
    starting_row = 2
    
    
    'Find the last row of worksheet
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).row   'identifies last row of worksheet
    
    
    'Create headers for summary_row
    ws.Range("I1").Value = "Ticker"
    ws.Range("I1").Font.Bold = True
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("J1").Font.Bold = True
    ws.Range("K1").Value = "Percent Change"
    ws.Range("K1").Font.Bold = True
    ws.Range("L1").Value = "Total Volume"
    ws.Range("L1").Font.Bold = True
    
        
    'Iterate through all rows in worksheet
    For row = 2 To last_row
    
        'If ticker changes
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then 'using not equal <>
                    
            'Add total_volume
            total_volume = total_volume + ws.Cells(row, 7).Value
            
            'Print ticker results
            ws.Range("I" & summary_row).Value = ws.Cells(row, 1).Value
            ws.Range("L" & summary_row).Value = total_volume
            
            'Reset total_volume and summary_row
            total_volume = 0
            
            'Calculate Open Price and Close Price
            open_price = ws.Cells(starting_row, 3).Value
            close_price = ws.Cells(row, 6)
        
           'Calculate yearly_change
            yearly_change = close_price - open_price
           
           'Print yearly_change in table
           ws.Range("J" & summary_row).Value = yearly_change
           
                    
                'Positive or Negative Change
                If yearly_change >= 0 Then
                    
                    'If Positive change to green
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4
                    
                    Else
                    
                    'If Negative change to red
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3
                    
                End If
            
                If open_price <> 0 Then
                                
                'Add a 0 to summary_row
                'ws.Range("J" & summary_row).Value = "N/A"
                'ws.Range("J" & summary_row).Interior.ColorIndex = 46
                'ws.Range("K" & summary_row).Value = "N/A"
                'ws.Range("L" & summary_row).Value = "N/A"
                     
                    'Else
                        'Percent Change from year opening price to year closing price
                         percent_change = (yearly_change / open_price)
                         
                         'Display percent_change in summary_row
                         ws.Range("K" & summary_row).Value = percent_change
                         
                         'Format percent_change
                         ws.Range("K" & summary_row).NumberFormat = "0.00%"
                    'End If
                 
                 
           End If
           
           'Setting new ticker starting row
            starting_row = (row + 1)
            
                  
           'Add 1 to summary_row
            summary_row = summary_row + 1
        
            'Reset percent_change to 0
            percent_change = 0
            
            'Add total_volume
            'total_volume = total_volume + ws.Cells(row, 7).Value
            
            
            
            End If
       
    Next row

Next ws
End Sub

