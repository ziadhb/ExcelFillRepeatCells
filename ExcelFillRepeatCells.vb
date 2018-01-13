'====================================================================================
'This function will fill the blanks in a given vertical range of one column.
'The blanks will be filled with the same value as the last cell above that is not blank
Sub RepeatValueInBlanks()
  Dim Rng As Range
  Set Rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
  TmpValue = ""
  For i = 1 To Rng.Rows.Count
    If Rng.Cells(i, 1).Value = "" Then
      Rng.Cells(i, 1).Value = TmpValue
    Else
      TmpValue = Rng.Cells(i, 1).Value
    End If
  Next i
End Sub
