Public Sub CheckAndDisplayMessage(cell_check As Range, cell_output As Range)
    Debug.Print "Checking:", cell_check.Address, "Value:", cell_check.Value
    If Not IsEmpty(cell_check.Value) Then
        If cell_check.Value = True Then
            MsgBox cell_output.Value, vbInformation, "Message Alerte"
        End If
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim watchCells As Range
    Dim cell As Range

    ' Set reference to the correct worksheet
    Set ws = Sheets("calculs_intermediaires")

    ' Use a workaround in case there's only one cell to monitor
    On Error Resume Next
    Set watchCells = Application.Union(ws.Range("AT2"), ws.Range("BA4"), ws.Range("BJ4"), _
                                       ws.Range("AR4"), ws.Range("AR13"), ws.Range("AS13"), _
                                       ws.Range("AT13"), ws.Range("AU13"))
    On Error GoTo 0

    ' Ensure that watchCells is not empty before proceeding
    If Not watchCells Is Nothing Then
        ' Check if the changed cell is within the watch list
        If Not Intersect(Target, watchCells) Is Nothing Then
            Call CheckAllCells
        End If
    End If
End Sub

Sub CheckAllCells()
    Dim wsCalc As Worksheet
    Dim wsPopUp As Worksheet
    
    ' Set references for better performance
    Set wsCalc = Sheets("calculs_intermediaires")
    Set wsPopUp = Sheets("pop_up")

    ' Check only if AT2 is TRUE before running further checks
    If wsCalc.Range("AT2").Value = True Then
        
        ' Check alertes unitaires
        If wsCalc.Range("BA4").Value = True Then
            Call CheckAndDisplayMessage(wsCalc.Range("BA4"), wsPopUp.Range("K3"))
            Debug.Print "TU1"
        End If
        
        If wsCalc.Range("BJ4").Value = True Then
            Call CheckAndDisplayMessage(wsCalc.Range("BJ4"), wsPopUp.Range("L3"))
            Debug.Print "TU2"
        End If
        
        ' Check alertes globales
        If wsCalc.Range("AR4").Value = True Then
            Call CheckAndDisplayMessage(wsCalc.Range("AR13"), wsPopUp.Range("H3"))
            Debug.Print "moyenne_courte < QN"
            
            Call CheckAndDisplayMessage(wsCalc.Range("AS13"), wsPopUp.Range("I3"))
            Debug.Print "moyenne_courte < LCI"
            
            Call CheckAndDisplayMessage(wsCalc.Range("AT13"), wsPopUp.Range("J3"))
            Debug.Print "moyenne_courte > LCS"
            
            Call CheckAndDisplayMessage(wsCalc.Range("AU13"), wsPopUp.Range("M3"))
            Debug.Print "moyenne_courte < QN"
        End If
    End If
End Sub


