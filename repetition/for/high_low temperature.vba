Sub HighTemp()
Dim i As Integer, nr As Integer, HT As Double, c As Integer
Call Reset
Range("A2").Select 'Ensure that the macro will choose the entire data
HT = InputBox("Please enter the temperature you want to find the days whose temperature higher:")
Selection.CurrentRegion.Select
nr = Selection.Rows.Count
For i = 2 To nr
        'debug.assert(cells(i,4)<HT) This helps to find the first required temperature quickly
        If Cells(i, 4) > HT Then
        c = c + 1
        If c = 1 Then
        Range("G1") = "Date"
        Range("H1") = "Temperature"
        End If
        Range("G" & c + 1) = Range("B" & i) & "/" & Range("C" & i) & "/" & Range("A" & i)
        Range("H" & c + 1) = Range("D" & i)
    End If
Next i
Range("A1").Select
End Sub

Sub Reset()
Columns("G:H").Clear
End Sub
