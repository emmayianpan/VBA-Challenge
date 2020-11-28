Option Explicit
Sub Greatest()
    Dim ws As Worksheet
    Dim max_increase As Double
    Dim min_increase As Double
    Dim max_volume As Double
    Dim LastRow As Long
    Dim i As Integer

    For Each ws In Worksheets
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        max_increase = ws.Application.Max(ws.Columns(11))
        ws.Range("Q2") = max_increase * 100 & "%"
        
        min_increase = ws.Application.Min(ws.Columns(11))
        ws.Range("Q3") = min_increase * 100 & "%"
        
        max_volume = ws.Application.Max(ws.Columns(12))
        ws.Range("Q4") = max_volume
        
        LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        
        For i = 2 To LastRow
            If ws.Cells(i, 11) = max_increase Then
                ws.Range("P2") = ws.Cells(i, 9)
            End If
            
            If ws.Cells(i, 11) = min_increase Then
                ws.Range("P3") = ws.Cells(i, 9)
            End If
            
            If ws.Cells(i, 12) = max_volume Then
                ws.Range("p4") = ws.Cells(i, 9)
            End If
            
            
        Next i
    Next ws
    
End Sub
