Attribute VB_Name = "Module1"
Sub StockMarketData()

   
   
        'Set variables for dataset
        Dim last_row As Double
        last_row = Cells(Rows.Count, "A").End(xlUp).Row
        
        'Set variables for output data
        Dim output_row As Double
        output_row = 1
        
        'Insert column headings
        Range("i1") = "Ticker"
        Range("j1") = "Yearly Change"
        Range("k1") = "Percent Change"
        Range("l1") = "Total Stock Volume"
        
        'Find Tickers: searches for when ticker in current cell is different from ticker in previous cell
        For i = 2 To last_row
        
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            
                'Move down one row in output table to create new row for next ticker
                output_row = output_row + 1
            
                'Insert ticker in new row
                Cells(output_row, 9).Value = Cells(i, 1).Value
                    
                    'Nested loop to find when ticker in next cell is different from ticker in current cell: this is for locating close price
                    For j = i To last_row
                        
                        'Find last row (j) of given ticker
                        If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
                        
                            Exit For
                                
                        End If
                    'Calculate the difference between open and close price of stock and place it in output table... Not sure why j + 1 is used for close price, this does not make logical sense to me...
                    Cells(output_row, 10).Value = Cells(j + 1, 6).Value - Cells(i, 3).Value
                        
                        'Format background fill of price change cells for +ve, -ve, and zero (green, red, blue)
                        If Cells(output_row, 10).Value < 0 Then
                            
                            Cells(output_row, 10).Interior.ColorIndex = 3
                            
                        ElseIf Cells(output_row, 10).Value > 0 Then
                            
                            Cells(output_row, 10).Interior.ColorIndex = 4
                                
                        Else: Cells(output_row, 10).Interior.ColorIndex = 8
                                
                                
                        End If
                            
                            'Calculate percentage change in price of stock and format percentage change output cells to percent
                            Cells(output_row, 11).Value = FormatPercent(Cells((j + 1), 6).Value / Cells(i, 3).Value - 1, 2, vbTrue)
                            
                    Next j
                        
            End If
        
        Next i
        
    
        'calculate stock trading value and show in output table
        Dim vol_count As Double
        vol_count = 0
        'reset output row to 1 to start filling volume data from top of output table
        output_row = 1
    
        For k = 2 To last_row
        
            'Loop to add vol traded cells until a different ticker is recognised
            vol_count = vol_count + Cells(k, 7)
            
                If Cells(k, 1).Value <> Cells(k + 1, 1).Value Then
                
                    'shift down one row to establish new row in output table
                    output_row = output_row + 1
                    
                    'input volume sum for stock into new row
                    Cells(output_row, 12) = vol_count
                    
                    'reset volume sum to 0 in prep for summing of next stock trading volume
                     vol_count = 0
                    
                End If
        
        Next k
        
    
        'Greatest percentage increase
        Cells(2, 14).Value = "Greatest % increase"
        ' insert ticker: Cells(2,15).value =
        Cells(2, 16).Value = FormatPercent(Application.WorksheetFunction.Max(Range("k:k")))
        
        'Greatest percentage decrease
        Cells(3, 14).Value = "Greatest % decrease"
        'insert ticker: Cells(3,15).value =
        Cells(3, 16).Value = FormatPercent(Application.WorksheetFunction.Min(Range("k:k")))
        
        'Greatest total volume
        Cells(4, 14).Value = "Greatest total volume"
        'insert ticker: Cells(4,15).value =
        Cells(4, 16).Value = Application.WorksheetFunction.Max(Range("L:L"))
        
        
        


End Sub

