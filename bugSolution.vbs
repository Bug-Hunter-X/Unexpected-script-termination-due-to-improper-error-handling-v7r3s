Function MyFunction(param1, param2)
  On Error Resume Next
  If IsEmpty(param1) Or IsEmpty(param2) Then
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    Err.Clear
    ' Handle the error gracefully
    MsgBox "Error: Parameters cannot be empty. Error number: " & ErrNumber & ". Description: " & ErrDescription, vbCritical
    ' Return a default value or handle the situation appropriately
    MyFunction = Null
    Exit Function
  End If
  On Error GoTo 0
  ' ... rest of the function code ...
End Function