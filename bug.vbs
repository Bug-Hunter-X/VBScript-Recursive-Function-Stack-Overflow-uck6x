Function f(a)
  If a = 1 Then Exit Function 
  f(a-1)
end function
MsgBox f(5)