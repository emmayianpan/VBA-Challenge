Option Explicit

Sub Basic()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim year_open As Double
    Dim year_high As Double
    Dim year_low As Double
    Dim year_close As Double
    Dim Total_Stock_Volume As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim TableRow As Integer
    Dim LastRow As Long
    Dim i As Long
    
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        'ws.Range("M1").Value = "Year Open"
        'ws.Range("N1").Value = "Year Close"
        TableRow = 2
        year_open = ws.Cells(2, 3)
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Total_Stock_Volume = 0
       
    
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                Ticker = ws.Cells(i, 1)
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
                year_close = ws.Cells(i, 6)
                Yearly_Change = year_close - year_open
                            
                            
                If year_open = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = (Yearly_Change / year_open) * 100
        
                End If
                    ws.Range("I" & TableRow).Value = Ticker
                    ws.Range("J" & TableRow).Value = Yearly_Change
                    ws.Range("K" & TableRow).Value = Percent_Change & "%"
                    ws.Range("L" & TableRow).Value = Total_Stock_Volume
                    'ws.Range("M" & TableRow).Value = year_open
                    'ws.Range("N" & TableRow).Value = year_close
                
            year_open = ws.Cells(i + 1, 3)
            Total_Stock_Volume = 0
            
            
                If ws.Range("J" & TableRow).Value >= 0 Then
                    ws.Range("J" & TableRow).Interior.Color = vbGreen
            
                Else
                    ws.Range("J" & TableRow).Interior.Color = vbRed
                End If
                    TableRow = TableRow + 1
            
            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7)
            
            End If
        
        
    
        Next i
    Next ws
    

End Sub