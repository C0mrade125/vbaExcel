Sub AddDataValidationList()

  Dim ws As Worksheet
  Dim rng As Range
  Dim dataRng As Range
  Dim dataStr As String

  ' Set the worksheet
  Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name

  ' Set the range where you want the data validation list
  Set rng = ws.Range("A1") ' Change "A1" to your desired cell or range

  ' Set the range containing the data for the list
  Set dataRng = ws.Range("B1:B20") ' Change "B1:B5" to your data range

  ' Create a comma-separated string from the data range
  dataStr = Join(Application.Transpose(dataRng.Value), ",")

  ' Add the data validation list
  With rng.Validation
    .Delete ' Remove any existing validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=dataStr
    .IgnoreBlank = True ' Allow blank cells
    .InCellDropdown = True ' Show the dropdown arrow
  End With

dataStr = Join(Application.Transpose(dataRng.Value), ",")
End Sub
