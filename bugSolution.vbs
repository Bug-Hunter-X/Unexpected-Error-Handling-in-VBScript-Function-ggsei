Function MyFunc(param1)
  If IsEmpty(param1) Then
    ' Handle the empty parameter gracefully
    MyFunc = ""
    Exit Function
  End If
  ' ... rest of the function ...
End Function