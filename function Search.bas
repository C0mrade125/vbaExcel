Function FindData(wsSource As Worksheet, wsTarget As Worksheet, ByVal Grade As String, ByVal MonthlyValue As Variant) As Long
    Dim lrSource As Long
    Dim i As Long
    Dim rowNum As Long
    Dim matchesFound As Long
    
    ' Find the last row in the source sheet
    lrSource = wsSource.Cells(wsSource.Rows.Count, 2).End(xlUp).row
    
    ' Initialize variables
    i = 4 ' Starting row for output in wsTarget
    matchesFound = 0 ' Counter for found matches
    
    ' Loop through each row in the source range
    For rowNum = 1 To lrSource
        If wsSource.Cells(rowNum, 6).Value = Grade And wsSource.Cells(rowNum, 38).Value = MonthlyValue Then
            ' Copy matching row to target sheet
            wsTarget.Cells(i, 1).Resize(1, 38).Value = wsSource.Cells(rowNum, 1).Resize(1, 38).Value
            i = i + 1 ' Move to the next row in the target sheet
            matchesFound = matchesFound + 1 ' Increment match count
        End If
    Next rowNum
    
    ' Return the count of matches found
   'FindData = matchesFound
End Function


Sub TestFindData()
    Dim countMatches As Long
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    
    ' Set source and target sheets
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsTarget = ThisWorkbook.Sheets("Sheet2")
    
    ' Call the function with specified worksheets
    countMatches = FindData(wsSource, wsTarget, "7A", wsSource.Range("AL4").Value)
    
    ' Display the result
    MsgBox countMatches & " rows copied to " & wsTarget.Name
End Sub

