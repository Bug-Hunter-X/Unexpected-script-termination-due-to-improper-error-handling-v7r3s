Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise vbObjectError + 1, , "Parameters cannot be empty"
  End If
  ' ... rest of the function code ...
End Function