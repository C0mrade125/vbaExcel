Function FindData(ByVal rngCell As Range, ByVal Grade As Variant, ByVal Monthly As Variant) As Range
    Dim cell As Range
    Dim rngTarget As Range
    Dim i As Long
    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Sheets("Sheet2")
    i = 4 ' Initialize the row counter

    For Each cell In rngCell
        ' Check if the Grade and Monthly criteria match
        If cell.Offset(0, 5).Value = Grade And cell.Offset(0, 68).Value = Monthly Then
            ' Copy the matching row to the target worksheet
            wsTarget.Cells(i, 1).Resize(1, 7).Value = rngCell(cell.Row, 1).Resize(1, 7).Value
            wsTarget.Cells(i, 8).Resize(1, 38).Value = rngCell(cell.Row, 39).Resize(1, 38).Value
            
            i = i + 1 ' Increment the row counter
        End If
    Next cell

End Function




Sub Show()
    Dim wsSource As Worksheet
    Dim rngCell As Range
    Dim Grade As Variant
    Dim Monthly As Variant
    Dim lr As Long
    
    ' Set the source and target worksheets
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    
    ' Set the criteria
    Grade = "7B"
    Monthly = wsSource.Range("BQ4").Value
    
    ' Find the last row in the source worksheet
    lr = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    ' Set the range to search within
    Set rngCell = wsSource.Range("A1:RQ" & lr)
    
    ' Call the FindData function to copy matching rows
    FindData rngCell, Grade, Monthly

End Sub

