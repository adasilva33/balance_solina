Sub ACTUALISATION_DATA_Bilan_Plannif()
        
    semaines_colonne = Array("G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH")
    lignes_machines = Array(4, 9, 14, 19, 24, 29, 34, 39, 44, 49, 54, 59, 64, 69, 74, 79, 84, 89, 94)

    Dim semaine_nom_plannif As Variant
    Dim cellule_recherche As Range
    Dim ligne As Integer
    Dim colonne As Integer
    
    For Each machine In lignes_machines
    
        ligne = ligne_machine(Range("D" & machine))
        
        If ligne < 77 Then
            ligne = ligne + 5
        Else
            ligne = ligne + 4
        End If
        
        For Each semaine In semaines_colonne
            colonne = 0
            semaine_nom_plannif = "Semaine considérée : " & Range(semaine & "3")
            colonne = colonne_semaine_consideree_plannification(semaine_nom_plannif) + 6
            Sheets("Bilan plannif").Range(semaine & machine).Value = Worksheets("Plannification").Cells(ligne, colonne)
        Next semaine
        
    Next machine
End Sub


Sub masquer_colonnes()
    Dim semaines_colonne As Variant
    semaines_colonne = Array("G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE")
    Dim i As Long
    Dim rep As Boolean
    rep = False
    i = 0
    
    While rep = False And i < UBound(semaines_colonne) + 1
        If ObtenirNumeroSemaine(Sheets("Bilan Plannif").Range(semaines_colonne(i) & "3")) > ObtenirNumeroSemaine(Sheets("Bilan Plannif").Range("B3")) Then
            rep = True
        End If
        i = i + 1
    Wend
    
    ' Masquer les colonnes de H ö la colonne pr_c_dente
    If rep = True Then
        MasquerColonnes "Bilan Plannif", "G:" & semaines_colonne(i - 3)
    End If
End Sub

Sub MasquerColonnes(nomFeuille As String, plageColonnes As String)
    Dim feuille As Worksheet
    Set feuille = ThisWorkbook.Sheets(nomFeuille)
    feuille.Columns.Hidden = False
    feuille.Columns(plageColonnes).Hidden = True
End Sub
Sub Finale_Bilan()
    Call masquer_colonnes
    Call ACTUALISATION_DATA_Bilan_Plannif

End Sub
Function colonne_semaine_consideree_plannification(semaine_consideree As Variant) As Integer
    If Not IsError(Application.Match(semaine_consideree, Sheets("Plannification").Range("4:4"), 0)) Then
        sortie = Application.Match(semaine_consideree, Sheets("Plannification").Range("4:4"), 0)
    End If
    colonne_semaine_consideree_plannification = sortie
End Function

Function ligne_machine(machine_consideree As String) As Integer
    If Not IsError(Application.Match(machine_consideree, Sheets("Plannification").Range("E:E"), 0)) Then
        sortie = Application.Match(machine_consideree, Sheets("Plannification").Range("E:E"), 0)
    End If
    ligne_machine = sortie
End Function

Function ObtenirNumeroSemaine(ByVal dateSemaine As String) As Integer
    Dim tableauDate As Variant
    tableauDate = Split(dateSemaine, "-")
    
    Dim numeroSemaine As Integer
    numeroSemaine = CInt(tableauDate(1))
    
    ObtenirNumeroSemaine = numeroSemaine
End Function

