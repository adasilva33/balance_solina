Public Sub CheckAndDisplayMessage(cell_check As Range, cell_output As Range)
    Debug.Print "Checking:", cell_check.Address, "Value:", cell_check.Value
    If Not IsEmpty(cell_check.Value) Then
        If cell_check.Value = True Then
            MsgBox cell_output.Value, vbInformation, "Message Alerte"
        End If
    End If
End Sub

Private Sub Worksheet_Calculate()
    Call CheckAllCells
End Sub

Sub CheckAllCells()

    If Sheets("calculs_intermediaires").Range("AT2").Value = True Then
        
        'Checks alertes unitaires
        If Sheets("calculs_intermediaires").Range("BA4").Value = True Then
            'TU1
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("BA4"), Sheets("pop_up").Range("K3"))
            Debug.Print "TU1"
        End If
        
        If Sheets("calculs_intermediaires").Range("BJ4").Value = True Then
          'TU2
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("BJ4"), Sheets("pop_up").Range("L3"))
            Debug.Print "TU2"
        End If
        
        'Checks alertes globales
        If Sheets("calculs_intermediaires").Range("AR4").Value = True Then
            
            'MsgBox "CheckAllCells is running - AR4 is TRUE!", vbInformation
            
            ' moyenne courte inf QN
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("AR13"), Sheets("pop_up").Range("H3"))
            Debug.Print "moyenne_courte < QN"
            
            ' moyenne courte inf LCI
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("AS13"), Sheets("pop_up").Range("I3"))
            Debug.Print "moyenne_courte < LCI"
            
            ' moyenne courte sup LCS
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("AT13"), Sheets("pop_up").Range("J3"))
            Debug.Print "moyenne_courte > LCS"
            
            ' moyenne longue inf QN
            Call CheckAndDisplayMessage(Sheets("calculs_intermediaires").Range("AU13"), Sheets("pop_up").Range("M3"))
            Debug.Print "moyenne_courte < QN"
    
        End If
    End If
    
End Sub


