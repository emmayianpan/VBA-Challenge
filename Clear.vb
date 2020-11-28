Sub Clear()
    For Each ws In Worksheets
        ws.Range("I:Q").Clear
    Next ws

End Sub