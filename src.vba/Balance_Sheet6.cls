Public Sub CheckAndDisplayMessage(cell_check As Range, cell_output As Range)
    ' Ensure the cell_check is not empty or invalid
    If Not IsEmpty(cell_check.Value) Then
        ' Check if the value in cell_check is True
        If cell_check.Value = True Then
            ' Display the value from cell_output
            MsgBox cell_output.Value, vbInformation, "Message Alerte"
        'Else
        '    MsgBox "The value in the check cell is not True.", vbExclamation, "Invalid Check"
        End If
    Else
        MsgBox "The check cell is empty.", vbExclamation, "Empty Cell"
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    'moyenne courte inf QN
    If Not Intersect(Target, Sheets("test").Range("C9")) Is Nothing Then
        Call CheckAndDisplayMessage(Sheets("test").Range("C9"), Sheets("pop_up").Range("H3"))
    End If
    
    'moyenne courte inf LCI
    If Not Intersect(Target, Sheets("test").Range("D9")) Is Nothing Then
        Call CheckAndDisplayMessage(Sheets("test").Range("D9"), Sheets("pop_up").Range("I3"))
    End If
    
    'moyenne courte sup LCS
    If Not Intersect(Target, Sheets("test").Range("E9")) Is Nothing Then
        Call CheckAndDisplayMessage(Sheets("test").Range("E9"), Sheets("pop_up").Range("J3"))
    End If
    
    'TU1
    If Not Intersect(Target, Sheets("test").Range("F9")) Is Nothing Then
        Call CheckAndDisplayMessage(Sheets("test").Range("F9"), Sheets("pop_up").Range("K3"))
    End If
    
    'TU2
    If Not Intersect(Target, Sheets("test").Range("G9")) Is Nothing Then
        Call CheckAndDisplayMessage(Sheets("test").Range("G9"), Sheets("pop_up").Range("L3"))
    End If
    
    'moyenne longue inf QN
    If Not Intersect(Target, Sheets("test").Range("H9")) Is Nothing Then
        Call CheckAndDisplayMessage(Sheets("test").Range("H9"), Sheets("pop_up").Range("M3"))
    End If
End Sub
