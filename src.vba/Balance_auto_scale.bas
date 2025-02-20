Sub AdjustYAxis()
    Dim chartObj As ChartObject
    Dim ws As Worksheet
    Dim minY As Variant, maxY As Variant
    Dim sheetPassword As String

    ' Définir explicitement la feuille de travail
    Set ws = ThisWorkbook.Sheets("interface") ' Remplacez par le bon nom
    sheetPassword = "gaetan" ' Remplacez par votre mot de passe

    ' Déprotéger la feuille
    On Error Resume Next
    ws.Unprotect Password:=sheetPassword
    On Error GoTo 0

    ' Vérifier s'il y a des graphiques
    If ws.ChartObjects.Count = 0 Then
        MsgBox "Aucun graphique trouvé sur la feuille " & ws.Name, vbExclamation, "Erreur"
        ws.Protect Password:=sheetPassword, UserInterfaceOnly:=True ' Reprotéger avant de quitter
        Exit Sub
    End If

    ' Vérifier la feuille contenant les valeurs min/max
    If Not WorksheetExists("calculs_intermediaires") Then
        MsgBox "La feuille 'calculs_intermediaires' n'existe pas.", vbExclamation, "Erreur"
        ws.Protect Password:=sheetPassword, UserInterfaceOnly:=True ' Reprotéger avant de quitter
        Exit Sub
    End If

    ' Récupérer les valeurs min/max
    minY = Sheets("calculs_intermediaires").Range("BX8").Value
    maxY = Sheets("calculs_intermediaires").Range("BX9").Value

    ' Vérifier si ce sont bien des nombres
    If Not IsNumeric(minY) Or Not IsNumeric(maxY) Then
        MsgBox "Les valeurs des cellules BU8 et BU9 ne sont pas valides.", vbExclamation, "Erreur"
        ws.Protect Password:=sheetPassword, UserInterfaceOnly:=True ' Reprotéger avant de quitter
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
    ws.Protect Password:=sheetPassword, UserInterfaceOnly:=True

    MsgBox "Ajustement terminé avec succès.", vbInformation, "Succes"
    Call MoveCursorToPopUpLastRow
End Sub

' Fonction pour vérifier l'existence d'une feuille
Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function
