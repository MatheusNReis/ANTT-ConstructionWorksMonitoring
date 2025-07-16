Sub MergeCellsInPairs()

    'Merge cells in pairs according to the first line, last line and columns defined in arrays

    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets(1)


    Dim FirstGroupLine As Long, LastGroupLine As Long
    Dim Columns As Variant
    
    FirstGroupLine = 219
    LastGroupLine = 224
    Columns = Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L")
    
    
    For j = LBound(Columns) To UBound(Columns)

        For i = FirstGroupLine To LastGroupLine Step 2
            ws.Range(Columns(j) & i & ":" & Columns(j) & i + 1).Merge
        Next i
    
    Next j
    
End Sub
