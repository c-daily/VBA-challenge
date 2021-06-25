Attribute VB_Name = "Module1"
Sub test()

Dim last_row As Long
Dim summary_table_row As Integer
Dim stock_volume As Double
Dim open_price As Double
Dim close_price As Double
Dim year_change As Double
Dim percent_change As Double

For Each ws In Worksheets
    
    ws.Activate
    
    ws.Range("K1").Value = "Ticker"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percent Change"
    ws.Range("N1").Value = "Total Stock Volume"
    
    year_change = 0
    open_price = 0
    percent_change = 0
    stock_volume = 0
    summary_table_row = 0
    
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

        For i = 2 To last_row
            
            If open_price = 0 Then
                open_price = Cells(i, 3).Value
            End If
            
            stock_volume = stock_volume + Cells(i, 7).Value
            
            stock_ticker = Cells(i, 1).Value
    
            If Cells(i + 1, 1).Value <> stock_ticker Then
                summary_table_row = summary_table_row + 1
                Cells(summary_table_row + 1, 11) = stock_ticker
                
                close_price = Cells(i, 6).Value
                
                year_change = close_price - open_price
                
                Cells(summary_table_row + 1, 12).Value = year_change
                
            If year_change > 0 Then
                Cells(summary_table_row + 1, 12).Interior.ColorIndex = 4
            ElseIf year_change < 0 Then
                Cells(summary_table_row + 1, 12).Interior.ColorIndex = 3
            Else: Cells(summary_table_row + 1, 12).Interior.ColorIndex = 6
            End If
            
            If open_price = 0 Then
                percent_change = 0
            Else: percent_change = (year_change / open_price)
            End If
            
            Cells(summary_table_row + 1, 13).Value = Format(percent_change, "Percent")
            
            open_price = 0
            
            Cells(summary_table_row + 1, 14).Value = stock_volume
            
            stock_volume = 0

            End If

        Next i

Next ws

End Sub
