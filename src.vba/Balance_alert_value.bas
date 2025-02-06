Public Sub CheckAndDisplayMessage(cell_check As Range, cell_output As Range)
    MsgBox "Checking:", cell_check.Address, "Value:", cell_check.Value
    If Not IsEmpty(cell_check.Value) Then
        If cell_check.Value = True Or LCase(cell_check.Value) = "true" Then
            MsgBox cell_output.Value, vbInformation, "Message Alerte"
        End If
    End If
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    Call CheckAllCells
End Sub

Sub CheckAllCells()
    'moyenne courte inf QN
    Call CheckAndDisplayMessage(Sheets("test").Range("C9"), Sheets("pop_up").Range("H3"))
    'moyenne courte inf LCI
    Call CheckAndDisplayMessage(Sheets("test").Range("D9"), Sheets("pop_up").Range("I3"))
    'moyenne courte sup LCS
    Call CheckAndDisplayMessage(Sheets("test").Range("E9"), Sheets("pop_up").Range("J3"))
    'TU1
    Call CheckAndDisplayMessage(Sheets("test").Range("F9"), Sheets("pop_up").Range("K3"))
    'TU2
    Call CheckAndDisplayMessage(Sheets("test").Range("G9"), Sheets("pop_up").Range("L3"))
    'moyenne longue inf QN
    Call CheckAndDisplayMessage(Sheets("test").Range("H9"), Sheets("pop_up").Range("M3"))
End Sub

