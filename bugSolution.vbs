Function f(a)
  If a = 1 Then
    f = 1  ' Explicitly return a value
    Exit Function
  Else
    f = a * f(a - 1) 'Correct recursive call with return value
  End If
end function
MsgBox f(5)