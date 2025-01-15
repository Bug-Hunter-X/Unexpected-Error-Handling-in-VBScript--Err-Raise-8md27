Function MyFunction(param1, param2)
  On Error Resume Next ' Enable error handling
  ' Some code here
  If param1 = "" Then
    Err.Raise 13, , "param1 cannot be empty"
  End If
  ' More code here
  If Err.Number <> 0 Then ' Check if an error occurred
    MsgBox "Error: " & Err.Description, vbCritical
    Err.Clear ' Clear the error object
    Exit Function ' Exit the function gracefully
  End If
  On Error GoTo 0 ' Disable error handling
End Function