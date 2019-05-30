Attribute VB_Name = "Module1"
Sub Stocks():
    Dim ticker As String 'variable for ticker symbol
    Dim vol As Double 'variable for stock volume additions
    Dim lastrow As Long 'variable to find value of last row
    Dim sht As Worksheet
    
    Set sht = ActiveSheet
    
    vol = 0

For Each ws In Worksheets
    lastrow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
    ws.Cells(1, 9) = "Ticker Symbol"
    ws.Cells(1, 10) = "Total Stock Volume"
    
    
    Dim sum_table_row As Integer
        sum_table_row = 2
        
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker = ws.Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value
            
            ws.Range("I" & sum_table_row).Value = ticker
            ws.Range("J" & sum_table_row).Value = vol
            
            sum_table_row = sum_table_row + 1
            
            vol = 0
        Else
            vol = vol + ws.Cells(i, 7).Value
        End If
    Next i
Next ws

End Sub
