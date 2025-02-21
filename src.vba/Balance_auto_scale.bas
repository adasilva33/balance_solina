Sub AdjustYAxis(minY As Double, maxY As Double)
    Dim chartObj As ChartObject
    Dim ws As Worksheet
    Dim sheetPassword As String

    ' Définir explicitement la feuille de travail
    Set ws = ThisWorkbook.Sheets("interface") ' Remplacez par le bon nom

    ' Déprotéger la feuille
    On Error Resume Next
    ws.Unprotect password:=sheetPassword
    On Error GoTo 0

    ' Vérifier s'il y a des graphiques
    If ws.ChartObjects.Count = 0 Then
        MsgBox "Aucun graphique trouvé sur la feuille " & ws.Name, vbExclamation, "Erreur"
        ws.Protect password:=sheetPassword, UserInterfaceOnly:=True ' Reprotéger avant de quitter
        Exit Sub
    End If

    ' Vérifier si les valeurs sont valides
    If Not IsNumeric(minY) Or Not IsNumeric(maxY) Then
        MsgBox "Les valeurs min et max ne sont pas valides.", vbExclamation, "Erreur"
        ws.Protect password:=sheetPassword, UserInterfaceOnly:=True ' Reprotéger avant de quitter
        Exit Sub
    End If

    ' Ajuster les axes des graphiques
    On Error Resume Next
    For Each chartObj In ws.ChartObjects
        With chartObj.chart.Axes(xlValue)
            .MinimumScale = minY
            .MaximumScale = maxY
        End With
    Next chartObj
    On Error GoTo 0

    ' Réactiver la protection après modifications
    ws.Protect password:=sheetPassword, UserInterfaceOnly:=True

    MsgBox "Ajustement terminé avec succès.", vbInformation, "Succès"
    Call MoveCursorToPopUpLastRow
End Sub
Sub CallVueGlobale()
    Dim minValue As Double
    Dim maxValue As Double
    
    ' Lire les valeurs depuis la feuille "calculs_intermediaires"
    minValue = Sheets("calculs_intermediaires").Range("BY8").Value
    maxValue = Sheets("calculs_intermediaires").Range("BY9").Value

    ' Appel de la fonction principale
    AdjustYAxis minValue, maxValue
End Sub
Sub CallRecentrer()
    Dim minValue As Double
    Dim maxValue As Double
    
    ' Lire les valeurs depuis la feuille "calculs_intermediaires"
    minValue = Sheets("calculs_intermediaires").Range("BX8").Value
    maxValue = Sheets("calculs_intermediaires").Range("BX9").Value

    ' Appel de la fonction principale
    AdjustYAxis minValue, maxValue
End Sub

' Fonction pour vérifier l'existence d'une feuille
Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function
