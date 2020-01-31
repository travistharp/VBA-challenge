Attribute VB_Name = "Module1"
Sub stocks()


'worksheet loop setup
Dim xsheet As Worksheet
For Each xsheet In ThisWorkbook.Worksheets
xsheet.Select

    'variables
    Dim ticker As String
    Dim op_price As Double
    Dim cl_price As Double
    Dim day_change As Double
    Dim change As Double
    Dim vol As Double
    Dim volume_total As Double
    Dim perchange As Double
    Dim last_row As Long
    Dim last_col As Long
    Dim sum_table_row As Integer




    'find number of rows
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'summary table setup
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Change"
    Range("K1").Value = "% Change"
    Range("L1").Value = "Total Stock Volume"
    
    sum_table_row = 2
    op_price = 0
    cl_price = 0
    
    For i = 2 To last_row
    
        'find yearly opening price
        If Cells(i, 2).Value = 20160101 Then
        op_price = Cells(i, 3).Value
        
        End If
        
        'find the change in ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            'set ticker symbols
            ticker = Cells(i, 1).Value
            
            'find year end close price
            cl_price = Cells(i, 6).Value
            'Cells(sum_table_row, 14).Value = cl_price
            'Cells(sum_table_row, 13).Value = op_price
            
            'find change in open/close
            change = cl_price - op_price
            
            'find yearly % change
            If op_price <> 0 Then
            percentage = (cl_price - op_price) / op_price
            End If
            
            
            'Print % change to summary
            Cells(sum_table_row, 11).Value = percentage
            Cells(sum_table_row, 11).NumberFormat = "0.00%"
            
            
            'find volume total
            volume_total = volume_total + Cells(i, 7).Value
            
            'Print ticker in summary
            Cells(sum_table_row, 9).Value = ticker
            
            'Print yearly change in summary and format
            Cells(sum_table_row, 10).Value = change
                
                If change < 1 Then
                Cells(sum_table_row, 10).Interior.ColorIndex = 3
                Else: Cells(sum_table_row, 10).Interior.ColorIndex = 4
                
                End If
            
            'Print total volume in summary
            Cells(sum_table_row, 12).Value = volume_total
            
            'Next Summary Row
            sum_table_row = sum_table_row + 1
            
            'reset total volume for next ticker
            volume_total = 0
            
            'reset stock price change
            change = 0
            
            'reset price change %
            percentage = 0
            
            'reset opening price
            oprice = 0
            cprice = 0
            
            
        Else
            'Add the volume total
            volume_total = volume_total + Cells(i, 7).Value
            
    
    
            
        End If
        
    
    
    Next i
    
Next xsheet


End Sub

