Function PopUpAndInputWithConfirmation(sheetName As String, messageCell As String, promptCell As String, inputType As Integer) As Variant
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
        If inputType = 0 Then
            userInput = Application.InputBox(promptText, "Saisie texte", Type:=2)
        ElseIf inputType = 1 Then
            userInput = Application.InputBox(promptText, "Saisie numérique", Type:=1)
        Else
            MsgBox "Type d'entrée non valide. Utilisez 0 pour texte ou 1 pour nombre.", vbExclamation, "Erreur"
            PopUpAndInputWithConfirmation = CVErr(xlErrValue)
            Exit Function
        End If

        If userInput = False Then
            MsgBox "Saisie annulée.", vbExclamation, "Annulé"
            PopUpAndInputWithConfirmation = CVErr(xlErrValue)
            Exit Function
        End If

        If inputType = 1 And Not IsNumeric(userInput) Then
            MsgBox "Veuillez entrer une valeur numérique.", vbExclamation, "Erreur"
        End If
    Loop While inputType = 1 And Not IsNumeric(userInput)

    confirmInput = MsgBox("Confirmez-vous la valeur saisie : " & userInput & " ?", vbYesNo + vbQuestion, "Confirmation")

    If confirmInput = vbYes Then
        PopUpAndInputWithConfirmation = userInput
        MsgBox "Valeur MAJ", vbInformation, "Enregistré"
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
    Call DisplayCellTextWithConfirmation("pop_up", "E3")
    Call MoveCursorToPopUpLastRow
End Sub

Sub DisplayDebutOf()
    Sheets("interface").Cells(3, "C").Value = PopUpAndInputWithConfirmation("pop_up", "C3", "C4", 0)
    Sheets("interface").Cells(4, "C").Value = PopUpAndInputWithConfirmation("pop_up", "C5", "C6", 0)
    Sheets("interface").Cells(5, "C").Value = PopUpAndInputWithConfirmation("pop_up", "C7", "C8", 1)
    Call DisplayCellTextWithConfirmation("pop_up", "C9")
    Call MoveCursorToPopUpLastRow
End Sub

Sub DisplayDebutEquipe()
    Call DisplayCellTextWithConfirmation("pop_up", "F3")
    Call DisplayCellTextWithConfirmation("pop_up", "F4")
    Call DisplayCellTextWithConfirmation("pop_up", "F5")
    Call DisplayCellTextWithConfirmation("pop_up", "F6")
    
    Sheets("calculs_intermediaires").Cells(7, "N").Value = PopUpAndInputWithConfirmation("pop_up", "F7", "F8", 1)

    Call DisplayCellTextWithConfirmation("pop_up", "F9")
    Call DisplayCellTextWithConfirmation("pop_up", "F10")
    Call MoveCursorToPopUpLastRow
End Sub
Sub DisplayFinEquipe()

    Call DisplayCellTextWithConfirmation("pop_up", "G3")
    Call MoveCursorToPopUpLastRow
End Sub
Sub CursorToLastRow()
Call MoveCursorToPopUpLastRow
End Sub

Function MoveCursorToPopUpLastRow()
    Dim ws As Worksheet
    Dim activeSheet As Worksheet
    Dim win As Window
    Dim popUpWindow As Window
    Dim lastRow As Long
    Dim targetCell As Range
    
    ' Store the current active sheet to restore later
    Set activeSheet = activeSheet
    
    ' Try to get the "pop_up" sheet
    On Error Resume Next
    Set ws = Sheets("data_brute")
    On Error GoTo 0

    ' If the sheet doesn't exist, exit the function
    If ws Is Nothing Then
        MsgBox "Sheet 'data_brute' does not exist!", vbExclamation
        Exit Function
    End If

    ' Find last used row in Column B and move one row down
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    Set targetCell = ws.Range("B" & lastRow)

    ' Identify the window displaying "pop_up" without changing focus
    For Each win In Application.Windows
        If win.Visible Then
            On Error Resume Next ' Prevent errors if no sheet is selected
            If win.SelectedSheets(1).Name = "data_brute" Then
                Set popUpWindow = win
                Exit For
            End If
            On Error GoTo 0
        End If
    Next win

    ' Only attempt to activate if a valid window was found
    If Not popUpWindow Is Nothing Then
        popUpWindow.Activate
        targetCell.Select
        activeSheet.Activate ' Restore focus to the original sheet in the first window
    Else
        MsgBox "The 'data_brute' sheet is not currently visible in another window.", vbExclamation
    End If
End Function

