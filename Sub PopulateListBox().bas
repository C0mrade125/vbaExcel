Sub PopulateListBox()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("List") ' Update with the correct sheet name
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' Clear existing items in the ListBox
    Me.ListBox1.Clear
    
    ' Loop through the table and add items to the ListBox
    For i = 1 To lastRow
        Me.ListBox1.AddItem ws.Cells(i, 1).Value
       
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = ws.Cells(i, 2).Value
        
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = ws.Cells(i, 3).Value
        
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = ws.Cells(i, 4).Value
        
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = ws.Cells(i, 5).Value
        
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = ws.Cells(i, 6).Value
        
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = ws.Cells(i, 7).Value
        
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 7) = ws.Cells(i, 8).Value
        
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 8) = ws.Cells(i, 9).Value
        
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 9) = ws.Cells(i, 10).Value
        
        
     
        
        ' Add the rest of the columns similarly
       
    Next i
    
End Sub
Private Sub cmdAdd_Click()
    Dim ws As Worksheet
    Dim newRow As Long
    
    Set ws = ThisWorkbook.Sheets("List")
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    
    ws.Cells(newRow, 1).Value = Me.txtID.Value
    ws.Cells(newRow, 2).Value = Me.txtName.Value
    ws.Cells(newRow, 3).Value = Me.cmbGender.Value
    ws.Cells(newRow, 4).Value = Me.cmbGrade.Value
    ' Continue assigning other text boxes to relevant worksheet columns
    
    ' Refresh the ListBox after adding
    Call PopulateListBox
End Sub
Private Sub cmdSearch_Click()
    Dim ws As Worksheet
    Dim i As Long
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets("List")
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        If ws.Cells(i, 1).Value = Me.txtID.Value Or ws.Cells(i, 2).Value = Me.txtName.Value Then
            Me.txtID.Value = ws.Cells(i, 1).Value
            Me.txtName.Value = ws.Cells(i, 2).Value
            Me.cmbGender.Value = ws.Cells(i, 3).Value
            Me.cmbGrade.Value = ws.Cells(i, 4).Value
            ' Load other fields similarly
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "Record not found"
    End If
End Sub

Private Sub cmdUpdate_Click()
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("List")
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        If ws.Cells(i, 1).Value = Me.txtID.Value Then
            ws.Cells(i, 2).Value = Me.txtName.Value
            ws.Cells(i, 3).Value = Me.cmbGender.Value
            ws.Cells(i, 4).Value = Me.cmbGrade.Value
            ' Continue updating other fields
            
            ' Refresh the ListBox after updating
            Call PopulateListBox
            Exit For
        End If
    Next i
End Sub
Private Sub cmdDelete_Click()
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("List")
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        If ws.Cells(i, 1).Value = Me.txtID.Value Or ws.Cells(i, 2) = Me.txtName.Value Then
            ws.Rows(i).Delete
            
            ' Refresh the ListBox after deleting
            Call PopulateListBox
            Exit For
        End If
    Next i
End Sub
Private Sub cmdClear_Click()
    Me.txtID.Value = ""
    Me.txtName.Value = ""
    Me.cmbGender.Value = ""
    Me.cmbGrade.Value = ""
    ' Clear other fields similarly
End Sub


Private Sub UserForm_Initialize()
    Call PopulateListBox
    
    ' Add items to ComboBoxes
    Me.cmbGender.AddItem "Male"
    Me.cmbGender.AddItem "Female"
    
    ' Add items for Grade
    Me.cmbGrade.AddItem "7A"
    Me.cmbGrade.AddItem "9A"
    ' Add more grades as needed
End Sub

