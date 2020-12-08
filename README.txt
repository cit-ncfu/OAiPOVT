Private M(99) As Double
' This is a comment beginning at the left edge of the screen.
Private Sub Command1_Click()
Dim i As Integer
List1.Clear
For i = 0 To 99
    M(i) = 100 * Rnd
    List1.AddItem M(i)
    Next i
End Sub

Private Sub Command2_Click()
Dim i As Integer, Sum As Double
Sum = 0
i = 99
Do While i >= 0
    Sum = Sum + M(i)
    i = i - 1
Loop
MsgBox Sum, vbExclamation + vbOKOnly, "Not found"  ' This is an inline comment.

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub List1_Click()

End Sub
