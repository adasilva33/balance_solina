Private Sub Worksheet_Calculate()
    Dim watchCells As Range
    Dim ws As Worksheet
    Set ws = Sheets("calculs_intermediaires")

    ' Define the cells we want to monitor
    Set watchCells = Application.Union(ws.Range("AT2"), ws.Range("BA4"), ws.Range("BJ4"), _
                                       ws.Range("AR4"), ws.Range("AR13"), ws.Range("AS13"), _
                                       ws.Range("AT13"), ws.Range("AU13"))

    ' Check if recalculated cells intersect with our watchCells
    If Not Intersect(Me.UsedRange, watchCells) Is Nothing Then
        Call CheckAllCells
    End If
End Sub

Public Sub CheckAndDisplayMessage(cell_check As Range, cell_output As Range)
    If Not IsEmpty(cell_check.Value) And cell_check.Value = True Then
        MsgBox cell_output.Value, vbInformation, "Message Alerte"
    End If
End Sub

Sub CheckAllCells()
    Dim wsCalc As Worksheet, wsPopup As Worksheet
    Set wsCalc = Sheets("calculs_intermediaires")
    Set wsPopup = Sheets("pop_up")

    If wsCalc.Range("AT2").Value = True Then
        If wsCalc.Range("BA4").Value = True Then Call CheckAndDisplayMessage(wsCalc.Range("BA4"), wsPopup.Range("K3"))
        If wsCalc.Range("BJ4").Value = True Then Call CheckAndDisplayMessage(wsCalc.Range("BJ4"), wsPopup.Range("L3"))
        If wsCalc.Range("AR4").Value = True Then
            Call CheckAndDisplayMessage(wsCalc.Range("AR13"), wsPopup.Range("H3"))
            Call CheckAndDisplayMessage(wsCalc.Range("AS13"), wsPopup.Range("I3"))
            Call CheckAndDisplayMessage(wsCalc.Range("AT13"), wsPopup.Range("J3"))
            Call CheckAndDisplayMessage(wsCalc.Range("AU13"), wsPopup.Range("M3"))
        End If
    End If
End Sub
