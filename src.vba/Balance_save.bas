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
    projectName = Sheets("interface").Range("B3").Value
    versionName = Sheets("interface").Range("B4").Value
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
    timeStamp = Format(Now, "yyyy-mm-dd_HH-MM-SS")

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
End Sub