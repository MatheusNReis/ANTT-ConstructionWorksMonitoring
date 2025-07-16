Sub InsertOneEmptyLineBetweenEachLine()

   'Insert one empty line between each line of a group of lines
    
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets(1)

    Dim FirstGroupLine As Long, LastGroupLine As Long
    FirstGroupLine = 29
    LastGroupLine = 37


    For i = FirstGroupLine To LastGroupLine Step 1
    
        ws.Rows(i + (i - FirstGroupLine)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        
    Next i
    
End Sub
