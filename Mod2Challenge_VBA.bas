Attribute VB_Name = "Module1"
Sub StockMarketData()
    
    'Set variables for dataset
    Dim ticker_data_col As Double
    Dim open_px As Double
    Dim close_px As Double
    Dim vol As Double
    Dim total_rows As Double
        
    ticker_data_col = 1
    open_px = 3
    close_px = 6
    vol = 7
    total_rows = 753001
    
    'Set variables for output data
    Dim output_row As Double
    Dim ticker_output_col As Double
    Dim yearly_change_output_col As Double
    Dim percent_change_output_col As Double
    Dim volume_output_col As Double
        
    output_row = 1
    ticker_output_col = 9
    yearly_change_output_col = 10
    percent_change_output_col = 11
    volume_output_col = 12
    
    'Insert column headings
    Range("i1") = "Ticker"
    Range("j1") = "Yearly Change"
    Range("k1") = "Percent Change"
    Range("l1") = "Total Stock Volume"
    
    'Find Tickers: searches for when next cell is different from the current cell
    For i = 2 To total_rows
    
        If Cells(i, ticker_data_col).Value <> Cells(i - 1, ticker_data_col).Value Then
        
            output_row = output_row + 1
        
            Cells(output_row, ticker_output_col).Value = Cells(i, ticker_data_col).Value
            
                For j = i To total_rows
            
                    If Cells(j, ticker_data_col).Value <> Cells(j + 1, ticker_data_col).Value Then
                    
                        Exit For
                            
                    End If
                                                                                 'Not sure why this vvv is j + 1, this does not make logical sense to me
                Cells(output_row, yearly_change_output_col).Value = Cells(j + 1, close_px).Value - Cells(i, open_px).Value
                
                    If Cells(output_row, yearly_change_output_col).Value < 0 Then
                        
                        Cells(output_row, yearly_change_output_col).Interior.ColorIndex = 3
                        
                    ElseIf Cells(output_row, yearly_change_output_col).Value > 0 Then
                        
                        Cells(output_row, yearly_change_output_col).Interior.ColorIndex = 4
                            
                    Else: Cells(output_row, yearly_change_output_col).Interior.ColorIndex = 8
                            
                            
                    End If
                        
                        Cells(output_row, percent_change_output_col).Value = FormatPercent(Cells(j + 1, close_px).Value / Cells(i, open_px).Value - 1, 2, vbTrue)
                        
                Next j
                    
        End If
    
    Next i
    

'stock trading value column
Dim vol_count As Double
vol_count = 0
output_row = 1

    For k = 2 To total_rows
    
        vol_count = vol_count + Cells(k, 7)
        
            If Cells(k, ticker_data_col).Value <> Cells(k + 1, ticker_data_col).Value Then
            
                output_row = output_row + 1
            
                Cells(output_row, volume_output_col) = vol_count
                
                 vol_count = 0
                
            End If
    
    Next k
    

'Greatest percentage increase
Cells(2, 14).Value = "Greatest % increase"
Cells(2, 15).Value = FormatPercent(Application.WorksheetFunction.Max(Range("k:k")))

'Greatest percentage decrease
Cells(3, 14).Value = "Greatest % decrease"
Cells(3, 15).Value = FormatPercent(Application.WorksheetFunction.Min(Range("k:k")))

'Greatest total volume
Cells(4, 14).Value = "Greatest total volume"
Cells(4, 15).Value = Application.WorksheetFunction.Max(Range("L:L"))
    
'Autofit columns to contents
Worksheets("2018").Cells.EntireColumn.AutoFit

    
End Sub

