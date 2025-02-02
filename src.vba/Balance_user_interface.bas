Function PopUpAndInputWithConfirmation(sheetName As String, messageCell As String, promptCell As String) As Variant
    Dim ws As Worksheet
    Dim messageText As String
    Dim promptText As String
    Dim userInput As Variant
    Dim confirmInput As VbMsgBoxResult
    
    ' Vérifier si la feuille existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "La feuille '" & sheetName & "' n'existe pas.", vbExclamation, "Erreur"
        PopUpAndInputWithConfirmation = CVErr(xlErrRef)
        Exit Function
    End If
    
    ' Récupérer les messages depuis la feuille spécifiée
    On Error Resume Next
    messageText = ws.Range(messageCell).Value
    promptText = ws.Range(promptCell).Value
    On Error GoTo 0
    
    If messageText = "" Or promptText = "" Then
        MsgBox "Les cellules spécifiées ne contiennent pas de texte valide.", vbExclamation, "Erreur"
        PopUpAndInputWithConfirmation = CVErr(xlErrValue)
        Exit Function
    End If
    
    MsgBox messageText, vbInformation, "Message"
    
    Do
        userInput = Application.InputBox(promptText, "Saisie numérique", Type:=1)
        
        If userInput = False Then
            MsgBox "Saisie annulée.", vbExclamation, "Annulé"
            PopUpAndInputWithConfirmation = CVErr(xlErrValue)
            Exit Function
        End If
    Loop While Not IsNumeric(userInput)
    
    confirmInput = MsgBox("Confirmez-vous la valeur saisie : " & userInput & " ?", vbYesNo + vbQuestion, "Confirmation")
    
    If confirmInput = vbYes Then
        PopUpAndInputWithConfirmation = userInput
        MsgBox "Tare MAJ", vbInformation, "Enregistré"
    Else
        MsgBox "Modification annulée. Aucune valeur n'a été enregistrée.", vbExclamation, "Annulé"
        PopUpAndInputWithConfirmation = CVErr(xlErrValue)
    End If
End Function

Sub DisplayCellTextWithConfirmation(sheetName As String, cellAddress As String)
    Dim ws As Worksheet
    Dim cellText As String
    Dim userResponse As VbMsgBoxResult
    
    ' Vérifier si la feuille existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "La feuille '" & sheetName & "' n'existe pas.", vbExclamation, "Erreur"
        Exit Sub
    End If
    
    ' Récupérer le texte de la cellule
    On Error Resume Next
    cellText = ws.Range(cellAddress).Value
    On Error GoTo 0
    
    If cellText = "" Then
        MsgBox "La cellule " & cellAddress & " est vide ou n'existe pas.", vbExclamation, "Erreur"
        Exit Sub
    End If
    
    ' Afficher la boîte de message avec le texte
    userResponse = MsgBox(cellText & vbCrLf & vbCrLf & "Confirmez-vous pour continuer ?", _
                          vbYesNo + vbQuestion, "Confirmation")
    
    ' Vérifier la réponse de l'utilisateur
    If userResponse = vbYes Then
        MsgBox "Confirmation reçue. Vous pouvez continuer.", vbInformation, "Validé"
    Else
        MsgBox "Action annulée par l'utilisateur.", vbExclamation, "Annulé"
        Exit Sub
    End If
End Sub

Sub DisplayFinOF()
    Call DisplayCellTextWithConfirmation("pop_up", "D3")
End Sub

Sub DisplayFinLot()
    Call DisplayCellTextWithConfirmation("pop_up", "E3")
End Sub

Sub DisplayDebutOf()
    Call DisplayCellTextWithConfirmation("pop_up", "C3")
End Sub
Sub DisplayDebutEquipe()
    Call DisplayCellTextWithConfirmation("pop_up", "F3")
    Call DisplayCellTextWithConfirmation("pop_up", "F4")
    Call DisplayCellTextWithConfirmation("pop_up", "F5")
    Call DisplayCellTextWithConfirmation("pop_up", "F6")
    
    Sheets("calculs_intermediaires").Cells(7, "N").Value = PopUpAndInputWithConfirmation("pop_up", "F7", "F8")

    Call DisplayCellTextWithConfirmation("pop_up", "F9")
    Call DisplayCellTextWithConfirmation("pop_up", "F10")
End Sub
Sub DisplayFinEquipe()
    Call DisplayCellTextWithConfirmation("pop_up", "G3")
End Sub



