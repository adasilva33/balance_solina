Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check if the changed cell is C2
    If Not Intersect(Target, Me.Range("C2")) Is Nothing Then
        ' Ensure it's not blank and is numeric
        If IsNumeric(Target.Value) And Target.Value > 0 Then
            MsgBox "La valeur dans C2 est plus grande que 0.", vbInformation, "Value Check"
        End If
    End If
End Sub