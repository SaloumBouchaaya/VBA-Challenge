Sub quarterly_data()

Dim i As Double
Dim ticker As String
Dim lastrow As Double
Dim summary_table_row As Double
Dim j As Double
Dim quarterly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim ws As Worksheet
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double
Dim opening_price As Double
Dim closing_price As Double



For Each ws In ThisWorkbook.Worksheets


    total_stock_volume = 0
    j = 2
    summary_table_row = 2

    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Quarterly_Change"
    ws.Range("L1").Value = "Percent_Change"
    ws.Range("M1").Value = "Total_Stock_Volume"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"




lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row


For i = 2 To lastrow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value


    ws.Range("J" & summary_table_row).Value = ticker
    summary_table_row = summary_table_row + 1

    ws.Range("M" & summary_table_row - 1).Value = total_stock_volume

    closing_price = Cells(i, 6)
    opening_price = Cells(j, 3)

    quarterly_change = closing_price - opening_price
    ws.Range("K" & summary_table_row - 1).Value = quarterly_change

If opening_price <> 0 Then
    percent_change = ((closing_price - opening_price) / opening_price) * 100
ElseIf closing_price <> 0 Then
    percent_change = ((closing_price - opening_price) / opening_price) * 100
ElseIf opening_price = 0 Then
    percent_change = 0
ElseIf closing_price = 0 Then
    percent_change = 0
End If
 
    ws.Range("L" & summary_table_row - 1).Value = percent_change
 
    

    total_stock_volume = 0
    
    
    j = i + 1

End If

If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value


End If


If ws.Cells(i, 11).Value > 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
ElseIf ws.Cells(i, 11).Value < 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 3
Else: ws.Cells(i, 11).Interior.ColorIndex = 2

End If


Next i

Next ws

End Sub

