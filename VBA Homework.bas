Attribute VB_Name = "Module1"
Sub StockAnalysis()
    
    Dim ticker As String
    
    Dim open_price As Double
    
    Dim close_price As Double
    
    Dim volume As Double
    
    Dim i As Double
    
    Dim summary_table_row As Double
    
    Dim opening_row As Double
    
    Dim last_row As Double
    
    Dim percent_change As Double
    
    'initalizing variables
    volume = 0
    
    opening_row = 2
    
    summary_table_row = 2
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    Debug.Print (last_row)
    
    'looping begins here
    
    For i = 2 To last_row
    
        ticker = Cells(i, 1).Value
        
        volume = volume + Cells(i, 7).Value
        
        If ticker <> Cells(i + 1, 1).Value Then
        
            'write ticker to summary table
            
            Cells(summary_table_row, 9).Value = ticker
            
            Cells(summary_table_row, 12).Value = volume
            
            open_price = Cells(opening_row, 3).Value
            
            close_price = Cells(i, 6).Value
            
            yearly_change = close_price - open_price
            
            Cells(summary_table_row, 10).Value = yearly_change
            
            'conditional for Yearly_Change
            
            If yearly_change > 0 Then
        
                Cells(summary_table_row, 10).Interior.ColorIndex = 4
            
            Else
            
                Cells(summary_table_row, 10).Interior.ColorIndex = 3
            
            End If
            
            'formula for percent_change
            
            percent_change = yearly_change / open_price
            
            Cells(summary_table_row, 11).Value = percent_change
            
            
            
            'reset volume to zero
            
            volume = 0
            
            summary_table_row = summary_table_row + 1
            
            opening_row = i + 1
            
    
        End If
        
    
    Next i
    
    
    Range("K:K").NumberFormat = "0.00%"
    
    Cells(1, 9).Value = "Ticker"
    
    Cells(1, 10).Value = "Yearly Change"
    
    Cells(1, 11).Value = "Percent Change"
    
    Cells(1, 12).Value = "Total Stock Value"
    
    
    










End Sub
