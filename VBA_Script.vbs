Sub StockData()
For Each ws In Worksheets
    ' Create all the required variables for calculation
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As Double
    Dim output_row As Integer
    Dim price_calculation_row As Long
    
    
    ' Set output information header values
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Get the last row of the worksheet
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set the output and price calculation row to print the output information
    output_row = 2
    price_calculation_row = 2
    
    ' Loop through the whole sheet to get the required output information
    For i = 2 To last_row:
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ' Calculation for getting the output variables
            ticker = ws.Cells(i, 1).Value
            open_price = ws.Cells(price_calculation_row, 3).Value
            close_price = ws.Cells(i, 6).Value
            yearly_change = close_price - open_price
            percent_change = yearly_change / ws.Cells(price_calculation_row, 3).Value
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            ' Set the output information cells
            ws.Range("I" & output_row).Value = ticker
            ws.Range("J" & output_row).Value = yearly_change
            ws.Range("K" & output_row).Value = percent_change
            ws.Range("L" & output_row).Value = total_stock_volume
            
            ' Set the colour index for yearly change
            If ws.Range("J" & output_row).Value > 0 Then
                ws.Range("J" & output_row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & output_row).Interior.ColorIndex = 3
            End If
            
            ' Set the format for percent change
            ws.Range("K" & output_row).NumberFormat = "0.00%"
            
            ' Set the output and price calculation row value
            output_row = output_row + 1
            price_calculation_row = i + 1
            
            ' Reset the total stock volume variable to zero for next ticker
            total_stock_volume = 0
        Else
            ' Add the volume for each row
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        End If
    Next i
    
    ' Calculation for greatest increase, decrease and total volume
    ' Created the headers
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ' Create the required variables
    Dim ticker_increase As String
    Dim ticker_decrease As String
    Dim ticker_total_volume As String
    Dim increase_value As Double
    Dim decrease_value As Double
    Dim total_volume_value As Double
    
    ' Set the variable values as the values for the first ticker
    ticker_increase = ws.Cells(2, 9).Value
    ticker_decrease = ws.Cells(2, 9).Value
    ticker_total_volume = ws.Cells(2, 9).Value
    increase_value = ws.Cells(2, 11).Value
    decrease_value = ws.Cells(2, 11).Value
    total_volume_value = ws.Cells(2, 12).Value
    
    ' Get the last row of the worksheet
    last_row_value = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Loop through all the output data to get the greatest values
    For j = 2 To last_row_value:
        If ws.Cells(j + 1, 11).Value > increase_value Then
            ticker_increase = ws.Cells(j + 1, 9).Value
            increase_value = ws.Cells(j + 1, 11).Value
        ElseIf ws.Cells(j + 1, 11).Value < decrease_value Then
            ticker_decrease = ws.Cells(j + 1, 9).Value
            decrease_value = ws.Cells(j + 1, 11).Value
        End If
        If ws.Cells(j + 1, 12).Value > total_volume_value Then
            ticker_total_volume = ws.Cells(j + 1, 9).Value
            total_volume_value = ws.Cells(j + 1, 12).Value
        End If
    Next j
    
    ' Set the output information cells
    ws.Range("P2").Value = ticker_increase
    ws.Range("P3").Value = ticker_decrease
    ws.Range("P4").Value = ticker_total_volume
    ws.Range("Q2").Value = increase_value
    ws.Range("Q3").Value = decrease_value
    ws.Range("Q4").Value = total_volume_value
    
    ' Set the format for percent increase and percent decrease column
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
Next ws
End Sub