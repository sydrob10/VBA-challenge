Sub VBAHW()
    For Each ws In Worksheets
        ws.Activate
        Call stockdata
    Next ws
End Sub
Sub stockdata()

    Dim current_ticker As String
    Dim next_ticker As String
    Dim last_row As Double
    Dim summary_row As Integer
    Dim opening_price As Double
    Dim closing_price As Double
    Dim stock_volume As Double
                
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    summary_row = 2
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    opening_price = Cells(2, 3).Value
    stock_volume = Cells(2, 7).Value
                   
    For current_row = 2 To last_row
        current_ticker = Cells(current_row, 1).Value
        next_ticker = Cells(current_row + 1, 1).Value
        stock_volume = stock_volume + Cells(current_row, 7).Value
                
        If current_ticker <> next_ticker Then
            Cells(summary_row, 9).Value = current_ticker
            
            closing_price = Cells(current_row, 6)
            yearly_change = closing_price - opening_price
            Cells(summary_row, 10).Value = yearly_change
            If yearly_change >= 0 Then
                Cells(summary_row, 10).Interior.ColorIndex = 4
                Else
                Cells(summary_row, 10).Interior.ColorIndex = 3
            End If
            
            If opening_price = 0 Then
                opening_price = Cells(current_row + 1, 3).Value
            End If
            
            percent_change = (yearly_change / opening_price)
            Cells(summary_row, 11).Value = percent_change
            
            Cells(summary_row, 12).Value = stock_volume
            
            summary_row = summary_row + 1
            opening_price = Cells(current_row + 1, 3).Value
            stock_volume = 0
        
        End If
                       
    Next current_row
    
    Range("I:L").Columns.AutoFit
    Columns("K:K").NumberFormat = "0.00%"
        
End Sub

