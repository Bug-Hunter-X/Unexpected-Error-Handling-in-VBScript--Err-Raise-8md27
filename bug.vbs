Function MyFunction(param1, param2)
  ' Some code here
  If param1 = "" Then
    Err.Raise 13, , "param1 cannot be empty"
  End If
  ' More code here
End Function