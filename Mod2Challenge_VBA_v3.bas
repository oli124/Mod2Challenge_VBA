Attribute VB_Name = "Module11"
Sub StockMarketData()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        'Set variables for dataset
        Dim last_row As Double
        last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Set variables for output data
        Dim output_row As Double
        output_row = 1
        
        'Insert column headings
        ws.Range("i1,o1") = "Ticker"
        ws.Range("j1") = "Yearly Change"
        ws.Range("k1") = "Percent Change"
        ws.Range("l1") = "Total Stock Volume"
        ws.Range("p1") = "Value"
        
        'Find Tickers: searches for when ticker in current cell is different from ticker in previous cell
        For i = 2 To last_row
        
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
                'Move down one row in output table to create new row for next ticker
                output_row = output_row + 1
            
                'Insert ticker in new row
                ws.Cells(output_row, 9).Value = ws.Cells(i, 1).Value
                    
                    'Nested loop to find when ticker in next cell is different from ticker in current cell: this is for locating close price
                    For j = i To last_row
                        
                        'Find last row (j) of given ticker
                        If ws.Cells(j, 1).Value <> ws.Cells(j + 1, 1).Value Then
                        
                            Exit For
                                
                        End If
                    'Calculate the difference between open and close price of stock and place it in output table... Not sure why j + 1 is used for close price, this does not make logical sense to me...
                    ws.Cells(output_row, 10).Value = ws.Cells(j + 1, 6).Value - ws.Cells(i, 3).Value
                        
                        'Format background fill of price change cells for +ve, -ve, and zero (green, red, blue)
                        If ws.Cells(output_row, 10).Value < 0 Then
                            
                            ws.Cells(output_row, 10).Interior.ColorIndex = 3
                            
                        ElseIf ws.Cells(output_row, 10).Value > 0 Then
                            
                            ws.Cells(output_row, 10).Interior.ColorIndex = 4
                                
                        Else: ws.Cells(output_row, 10).Interior.ColorIndex = 8
                                
                                
                        End If
                            
                            'Calculate percentage change in price of stock and format percentage change output cells to percent
                            ws.Cells(output_row, 11).Value = ws.Cells((j + 1), 6).Value / ws.Cells(i, 3).Value - 1
                            
                            'Format as %
                            ws.Cells(output_row, 11) = FormatPercent(ws.Cells(output_row, 11), 2)
                            
                            
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
            vol_count = vol_count + ws.Cells(k, 7)
            
                If ws.Cells(k, 1).Value <> ws.Cells(k + 1, 1).Value Then
                
                    'shift down one row to establish new row in output table
                    output_row = output_row + 1
                    
                    'input volume sum for stock into new row
                    ws.Cells(output_row, 12) = vol_count
                    
                    'reset volume sum to 0 in prep for summing of next stock trading volume
                     vol_count = 0
                    
                End If
        
        Next k
        
        Dim last_row_output As Double
        last_row_output = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        'Greatest percentage increase
        ws.Cells(2, 14).Value = "Greatest % increase"
         
        ws.Cells(2, 16).Value = FormatPercent(Application.WorksheetFunction.Max(ws.Range("k:k")))
            'Loop to match max % value to ticker row
            For m = 2 To last_row_output
                
                If ws.Cells(2, 16).Value = ws.Cells(m, 11).Value Then
                
                    Exit For
                    
                End If
                
            Next m
        'insert ticker for max %
         ws.Cells(2, 15).Value = ws.Cells(m, 9).Value
         
        'Greatest percentage decrease
        ws.Cells(3, 14).Value = "Greatest % decrease"
        
        ws.Cells(3, 16).Value = FormatPercent(Application.WorksheetFunction.Min(ws.Range("k:k")))
            
            'Loop to match min % value to ticker row
            For n = 2 To last_row_output
                
                If ws.Cells(3, 16).Value = ws.Cells(n, 11).Value Then
                
                    Exit For
                    
                End If
                
            Next n
        'insert ticker for min %
        ws.Cells(3, 15).Value = ws.Cells(n, 9).Value
        
        'Greatest total volume
        ws.Cells(4, 14).Value = "Greatest total volume"
        'insert ticker: Cells(4,15).value =
        ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
        
            'Loop to match max vol value to ticker row
            For p = 2 To last_row_output
                
                If ws.Cells(4, 16).Value = ws.Cells(p, 12).Value Then
                
                    Exit For
                    
                End If
                
            Next p
        'insert ticker for max vol
        ws.Cells(4, 15).Value = ws.Cells(p, 9).Value
        
        'Autofit cells to contents
        ws.Cells.EntireColumn.AutoFit
    
    Next ws

End Sub

