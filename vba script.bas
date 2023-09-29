Attribute VB_Name = "Module1"
Sub stock_analysis()
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As Double
Dim opening_price As Double
Dim closing_price As Double

total_volume = 0

Dim summary_table_row As Integer
summary_table_row = 2

For i = 2 To 753001

    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
        opening_price = Cells(i, 3).Value
        
    End If
    
    total_volume = total_volume + Cells(i, 7).Value
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        
        ticker = Cells(i, 1).Value
       ' opening_price = Cells(i, 3).Value
        closing_price = Cells(i, 6).Value
        yearly_change = closing_price - opening_price
        
        If opening_price <> 0 Then
        
            percent_change = (yearly_change / opening_price) * 100
        Else
            percent_change = 0
    
        End If
        
        Range("I" & summary_table_row).Value = ticker
        Range("J" & summary_table_row).Value = yearly_change
        Range("K" & summary_table_row).Value = percent_change
        Range("L" & summary_table_row).Value = total_volume
        summary_table_row = summary_table_row + 1
        total_volume = 0
    Else
        
       ' yearly_change = closing_price - opening_price
      '  percent_change = (yearly_change / opening_price) * 100
      '  total_volume = total_volume + Cells(i, 7).Value
    
        
        
    End If
    
Next i


        
End Sub
