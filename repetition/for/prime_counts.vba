Function prime(n As Integer) As Boolean  'try sub later
Dim flag As Boolean, L As Double, i As Integer
flag = True
L = Int(n ^ 0.5)
For i = 2 To L
    If n Mod i = 0 Then
    flag = False
    Exit For
    End If
Next i
prime = flag
End Function

Function countprime(n1 As Integer, n2 As Integer) As Integer
Dim i As Integer, c As Integer
For i = n1 To n2
    If prime(i) = True Then c = c + 1
Next i
countprime = c
End Function
