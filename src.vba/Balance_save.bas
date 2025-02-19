Sub SaveCopyWithCustomName()
    Dim wb As Workbook
    Dim savePath As String
    Dim fileName As String
    Dim timeStamp As String
    Dim folderPath As String

    ' Définir le fichier actif
    Set wb = ActiveWorkbook

    ' Récupérer les valeurs de la feuille "interface"
    Dim projectName As String, versionName As String
    On Error Resume Next
    projectName = Sheets("interface").Range("C3").Value
    versionName = Sheets("interface").Range("C4").Value
    On Error GoTo 0

    ' Vérifier si les cellules sont remplies
    If projectName = "" Then
        MsgBox "Erreur : La cellule B3 (nom du projet) est vide.", vbExclamation
        Exit Sub
    End If
    If versionName = "" Then
        MsgBox "Erreur : La cellule B4 (version) est vide.", vbExclamation
        Exit Sub
    End If

    ' Nettoyer le nom du fichier pour éviter les caractères invalides
    projectName = Replace(projectName, "/", "-")
    versionName = Replace(versionName, "/", "-")

    ' Format du timestamp (jour_heure_minute_seconde)
    timeStamp = Format(Now, "yyyymmdd")

    ' Construire le nom du fichier
    fileName = projectName & "_" & versionName & "_" & timeStamp & ".xlsm"

    ' Définir un dossier par défaut (même dossier que le fichier original)
    folderPath = wb.Path

    ' Vérifier si le chemin du dossier est valide
    If folderPath = "" Then
        folderPath = Application.DefaultFilePath ' Utiliser le répertoire par défaut d'Excel
    End If

    ' Construire le chemin complet
    savePath = folderPath & "\" & Sheets("interface").Range("C6").Value & "\" & fileName

    ' Vérifier si le chemin d'accès est accessible
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Erreur : Le dossier cible n'existe pas ou est inaccessible.", vbExclamation
        Exit Sub
    End If

    ' Sauvegarder la copie du fichier
    On Error Resume Next
    wb.SaveCopyAs savePath
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'enregistrement du fichier. Vérifiez l'accès au dossier.", vbCritical
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Confirmation
    MsgBox "Une copie du fichier a été enregistrée sous : " & savePath, vbInformation
    
    Call MoveCursorToPopUpLastRow
End Sub

Sub SaveOrCopyWorkbook()
    Dim wb As Workbook
    Dim savePath As String
    Dim fileName As String
    Dim timeStamp As String
    Dim folderPath As String
    Dim saveOption As Boolean
    
    ' Définir le fichier actif
    Set wb = ActiveWorkbook
    
    ' Récupérer les valeurs de la feuille "interface"
    Dim targetPath As String, projectName As String, versionName As String
    On Error Resume Next
    targetPath = Sheets("calculs_intermediaires").Range("O4").Value
    saveOption = Sheets("calculs_intermediaires").Range("O7").Value ' Vrai/Faux
    projectName = Sheets("interface").Range("C3").Value ' Nom du projet
    versionName = Sheets("interface").Range("C4").Value ' Version
    On Error GoTo 0
    
    ' Vérifier si les cellules sont remplies
    If targetPath = "" Then
        MsgBox "Erreur : Le chemin de sauvegarde est vide.", vbExclamation
        Exit Sub
    End If
    If projectName = "" Then
        MsgBox "Erreur : Le nom du projet est vide.", vbExclamation
        Exit Sub
    End If
    If versionName = "" Then
        MsgBox "Erreur : La version est vide.", vbExclamation
        Exit Sub
    End If
    
    ' Format du timestamp (jour_heure_minute_seconde)
    timeStamp = Format(Now, "yyyymmdd")
    
    ' Construire le nom du fichier
    fileName = projectName & "_" & versionName & "_" & timeStamp & ".xlsm"
    savePath = targetPath & "\" & fileName
    
    ' Vérifier si le dossier cible existe
    If Dir(targetPath, vbDirectory) = "" Then
        MsgBox "Erreur : Le dossier cible n'existe pas ou est inaccessible.", vbExclamation
        Exit Sub
    End If
    
    ' Sauvegarder le fichier
    On Error Resume Next
    If saveOption Then
        ' Enregistrement sous un nouveau nom
        wb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Else
        ' Enregistrement d'une copie
        wb.SaveCopyAs savePath
    End If
    
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'enregistrement du fichier. Vérifiez l'accès au dossier.", vbCritical
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Confirmation
    If saveOption Then
        MsgBox "Le fichier du projet " & projectName & " (version " & versionName & ") a été enregistré sous un nouveau nom : " & savePath, vbInformation
    Else
        MsgBox "Une copie du fichier du projet " & projectName & " (version " & versionName & ") a été enregistrée sous : " & savePath, vbInformation
    End If
    
End Sub
