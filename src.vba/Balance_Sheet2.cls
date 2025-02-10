Public Sub CheckAndDisplayMessage(cell_check As Range, cell_output As Range)
    Debug.Print "Checking:", cell_check.Address, "Value:", cell_check.Value
    If Not IsEmpty(cell_check.Value) Then
        If cell_check.Value = True Or LCase(cell_check.Value) = "true" Or cell_check.Value = Vrai Or LCase(cell_check.Value) = "vrai" Then
            MsgBox cell_output.Value, vbInformation, "Message Alerte"
        End If
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Call CheckAllCells
End Sub

Private Sub Worksheet_Calculate()
    Call CheckAllCells
End Sub

Sub CheckAllCells()

    If Sheets("calculs_intermediaires").Range("AT2").Value = True Then
        'MsgBox "CheckAllCells is running - AT2 is TRUE!", vbInformation
        
        'Checks alertes unitaires
        If Sheets("calculs_intermediaires").Range("BA4").Value = True Then
            'TU1
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("BA4"), Sheets("pop_up").Range("K3"))
        End If
        
        If Sheets("calculs_intermediaires").Range("BJ4").Value = True Then
          'TU2
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("BJ4"), Sheets("pop_up").Range("L3"))
        End If
        
        'Checks alertes globales
        If Sheets("calculs_intermediaires").Range("AR4").Value = True Then
            
            'MsgBox "CheckAllCells is running - AR4 is TRUE!", vbInformation
            
            ' moyenne courte inf QN
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("AR13"), Sheets("pop_up").Range("H3"))
            
            ' moyenne courte inf LCI
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("AS13"), Sheets("pop_up").Range("I3"))
            
            ' moyenne courte sup LCS
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("AT13"), Sheets("pop_up").Range("J3"))
            
            ' moyenne longue inf QN
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("AU13"), Sheets("pop_up").Range("M3"))
    
        End If
    End If
    
End Sub


