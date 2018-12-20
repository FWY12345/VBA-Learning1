Function tank(R As Double, H As Double, d As Double)
'R = InputBox("Please enter the radius of tank:")
'd = InputBox("Please enter the depth of tank:")
'H = InputBox("Please enter the hight of tank:")
If d <= R Then
    tank = WorksheetFunction.Pi() * d ^ 2 / 3 * (3 * R - d)
ElseIf R < d And d <= H - R Then 'you need to use 'and' instead of R<d<=H-R
    tank = WorksheetFunction.Pi() * R ^ 3 / 3 * 2 + WorksheetFunction.Pi() * R ^ 2 * (d - R)
Else  'take attention here
    If H - R < d And d <= H Then
    tank = WorksheetFunction.Pi() * R ^ 3 / 3 * 4 + WorksheetFunction.Pi() * R ^ 2 * (H - 2 * R) - WorksheetFunction.Pi() * (H - d) ^ 2 / 3 * (3 * R - H + d)
    End If
End If
End Function
