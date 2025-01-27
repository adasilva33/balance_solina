Public colonnes_jour_semaine As Variant
Public lignes_capacite_journaliere As Variant

Sub Remplissage_DATA_Semaine_1()
   'Definition des variables globales
    colonnes_jour_semaine = Array("N", "O", "P", "Q", "R", "S", "T")
    lignes_capacite_journaliere = Array(18, 27, 36, 45, 54, 63, 72, 81, 89, 97, 105, 113, 121, 129, 137, 145, 153, 161, 169)
    'Definition du tableau structure
    Dim ws_dataset As Worksheet
    Dim tbl As ListObject
    
    Set ws_dataset = ThisWorkbook.Sheets("Lot commande")
    Set tbl = ws_dataset.ListObjects("Données_ERP")
    
    For Each ligne In lignes_capacite_journaliere
        For Each colonne In colonnes_jour_semaine
        
            'Ecart du nb d'equipe entre dispo et prevu
            Range("I" & ligne).Formula = "=-H" & ligne & "*G" & ligne & "+K" & ligne & "+L" & ligne + 1 & ""
            
            ' Mise en forme conditionnelle sur la cellule I & ligne
            Dim plage As Range
            Dim condition As FormatCondition

            ' Sp_cifiez la plage de cellules sur laquelle appliquer la mise en forme conditionnelle
            Set plage = Range("I" & ligne)

            ' Supprimez toutes les rgles de mise en forme conditionnelle existantes dans la plagee
            plage.FormatConditions.Delete

            ' Ajoutez une nouvelle rgle de mise en forme conditionnelle pour la plagee
            'Set condition = plage.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="-0,5", Formula2:="0,5")

            ' D_finissez le format de mise en forme pour la condition verte (valeur entre 0,5 et -0,5)
            'condition.Interior.Color = RGB(180, 255, 180)  ' Vert

            'D_finit MEFC
            Set condition = plage.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            condition.Interior.Color = RGB(255, 0, 0) ' Rouge

            Set condition = plage.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            condition.Interior.Color = RGB(180, 255, 180) ' Vert
            
            'Copie du nombre de poste/ jour
            Range(colonne & ligne).Formula = "=G" & ligne
            'Copie dispo semaine h
            Range(colonne & ligne - 1).Formula = "=I" & ligne
            'Copie dispo Nb heure theorique
            Range(colonne & ligne - 2).Formula = "=F" & ligne
            'Initialisation Nb heure theorique, poste et jour travaille
            Range("F" & ligne).Value = 6.75
            Range("G" & ligne).Value = 2
            Range("H" & ligne).Value = 5
            
            'Formule pour la prevision
            Range(colonne & ligne + 1).Formula = "=SUMIFS(Données_ERP[[1xXX]:[1xXX]], Données_ERP[[Poste de charge principal]:[Poste de charge principal]], $E$" & ligne & ", Données_ERP[[N° semaine planif]:[N° semaine planif]], " & colonne & "6, Données_ERP[[Jour semaine]:[Jour semaine]], " & colonne & "7, Données_ERP[[Statut]:[Statut]], ""Prévu"") / " & colonne & ligne - 2 & ""

             'Redistribution
            Range(colonne & ligne + 3).Formula = ""
           
             'Redistribution totale
            Range(colonne & ligne + 4).Formula = ""
            
            'Sous total
            If ligne > 77 Then
                Range(colonne & ligne + 4).Formula = "=SUM(" & colonne & ligne + 1 & ":" & colonne & ligne + 3 & ")"
            Else
                Range(colonne & ligne + 5).Formula = "=SUM(" & colonne & ligne + 1 & ":" & colonne & ligne + 4 & ")"
            End If
            
        Next colonne
    Next ligne

    
      'Total equipe
    For Each ligne In lignes_capacite_journaliere
        Range("L" & ligne).Value = ""
        Range("L" & ligne + 1).Formula = "=SUM(N" & ligne + 1 & ":T" & ligne + 1 & ")"
        Range("L" & ligne + 2).Formula = "=SUM(N" & ligne + 2 & ":T" & ligne + 2 & ")"
        Range("L" & ligne + 3).Formula = "=SUM(N" & ligne + 3 & ":T" & ligne + 3 & ")"
        Range("L" & ligne + 4).Formula = "=SUM(N" & ligne + 4 & ":T" & ligne + 4 & ")"
        
      
        If ligne < 77 Then
            Rows(ligne + 4).Hidden = False
            Range("M" & ligne + 3).Value = "Redistribution - Prévu"
            Range("M" & ligne + 4).Value = "Redistribution - S-1"
             
        Else
            Rows(ligne + 3).Hidden = False
            Range("M" & ligne + 2).Value = "Redistribution - Prévu"
            Range("M" & ligne + 3).Value = "Redistribution - S-1"
        End If
    Next ligne
          
End Sub
Sub formule_dispo_prevu_2_e_feuille()
    'Avant d'etirer il faut faire la S1 puis la S2 avec a, tout etirer et faire remplissage_Feuilles_22
    
    lignes_capacite_journaliere = Array(18, 27, 36, 45, 54, 63, 72, 81, 89, 97, 105, 113, 121, 129, 137, 145, 153, 161, 169)
    'Definition du tableau structure
    
    
    For Each ligne In lignes_capacite_journaliere
    
    'Ecart du nb d'equipe entre dispo et prevu
        Range("Y" & ligne).Formula = "=-X" & ligne & "*W" & ligne & "+AA" & ligne & "+AB" & ligne + 1 & ""
    Next ligne
    
End Sub
Sub Remplissage_Feuilles_2()

    Dim colonne As Integer
    Dim pas As Integer
    Dim i As Integer
    
    lignes_capacite_journaliere = Array(18, 27, 36, 45, 54, 63, 72, 81, 89, 97, 105, 113, 121, 129, 137, 145, 153, 161, 169)

    colonne = 11 ' D_finit la colonne de d_part pour Redistribution S-1
    pas = 16 ' D_finit le pas d'incr_mentation
    i = 0
    
    Do While i < 54
    
        'Redistribution S-1 : Probleme de decalage quand on etire
        Cells(7, colonne).Value = "Redistribution S-1"

        For Each ligne In lignes_capacite_journaliere
        
            'Donnees texte : meme probleme quand on etire
            If ligne < 77 Then
                Cells(ligne + 4, colonne + 2).Value = "Redistribution S-1"
                Cells(ligne + 2, colonne + 2).Value = "Redistribution - PC0"
                
                'Formule pour avoir ce qui n'a pas ete fait S-1
                 If i <> 0 Then
                    Cells(ligne, colonne).Formula = "=IF(" & Cells(ligne, colonne - pas - 2).Address & "<0, 0, " & Cells(ligne, colonne - pas - 2).Address & "-" & Cells(ligne + 4, colonne - pas + 1).Address & ")"
                End If
                
            Else
                Cells(ligne + 3, colonne + 2).Value = "Redistribution S-1"
                
                If i <> 0 Then
                    Cells(ligne, colonne).Formula = "=IF(" & Cells(ligne, colonne - pas - 2).Address & "<0, 0, " & Cells(ligne, colonne - pas - 2).Address & "-" & Cells(ligne + 3, colonne - pas + 1).Address & ")"
                End If
            End If
            
            'Valeurs
            Cells(ligne, colonne - 5).Value = 6.75
            Cells(ligne, colonne - 4).Value = 2
            Cells(ligne, colonne - 3).Value = 5
            
            'Cas particulier DOSETTES
            
            Cells(11, colonne - 5).Value = 6.75
            Cells(11, colonne - 4).Value = 2
            Cells(11, colonne - 3).Value = 5
            
        Next ligne
              
        colonne = colonne + pas
        i = i + 1
    
    Loop
End Sub

Sub ajuster()
    ActiveSheet.UsedRange.Columns.AutoFit
End Sub

Sub Masquer()

Dim lignes_capacite_journaliere As Variant
lignes_capacite_journaliere = Array(18, 27, 36, 45, 54, 63, 72, 81, 89, 97, 105, 113, 121, 129, 137, 145, 153, 161, 169)

Dim L1() As Variant
Dim L2() As Variant

ReDim L1(LBound(lignes_capacite_journaliere) To UBound(lignes_capacite_journaliere))
ReDim L2(LBound(lignes_capacite_journaliere) To UBound(lignes_capacite_journaliere))

Dim i As Long
For i = LBound(lignes_capacite_journaliere) To UBound(lignes_capacite_journaliere)
    L1(i) = lignes_capacite_journaliere(i) - 1
    L2(i) = lignes_capacite_journaliere(i) - 2
Next i

    For Each elmt In L1
        Rows(elmt).EntireRow.Hidden = True
    Next elmt

    For Each elmt In L2
        Rows(elmt).EntireRow.Hidden = True
    Next elmt
End Sub

Sub SupprimerColonnes()
    Dim ws As Worksheet
    Dim colonne As Range
    Dim colonneV As Range
    Dim colonneAGK As Range

    ' Sp_cifiez la feuille de calcul concern_e
    Set ws = ThisWorkbook.Sheets("Plannification")

    ' Sp_cifiez les colonnes V et AGK
    Set colonneV = ws.Columns("V")
    Set colonneAGK = ws.Columns("AGK")

    ' Supprime les colonnes de V ö AGK
    Application.DisplayAlerts = False ' D_sactive les alertes de suppression
    Range(colonneV, colonneAGK).Delete Shift:=xlToLeft ' Supprime les colonnes et d_cale les colonnes restantes vers la gauche
    Application.DisplayAlerts = True ' R_active les alertes

    MsgBox "Supprimées"
End Sub

Sub Etendre_vers_2()
    Range("F4:U173").Select
    Selection.AutoFill Destination:=Range("F4:AK173"), Type:=xlFillDefault
    Range("F4:AK173").Select
End Sub
Sub Etendre_vers_final()
    Range("V4:AK173").Select
    Selection.AutoFill Destination:=Range("V4:AGK173"), Type:=xlFillDefault
    Range("V4:AGK173").Select
End Sub

Sub total()

    Call Remplissage_DATA_Semaine_1
    Call Etendre_vers_2
    Call formule_dispo_prevu_2_e_feuille
    Call Etendre_vers_final
    Call Remplissage_Feuilles_2
    Call ajuster
    MsgBox "FINI"
End Sub

Sub RamenerSemaine()
    Dim numeroSemaine As Integer
    Dim celluleSemaine As Range
    Dim plageSemaines As Range
    
    ' R_cup_rer le num_ro de semaine de la cellule I2
    numeroSemaine = Range("I2").Value
    
    ' D_finir la plage contenant les cellules avec les semaines
    Set plageSemaines = Rows(4).Find("Semaine considérée").Offset(0, 1).EntireRow
    
    ' V_rifier si le num_ro de semaine existe dans la plage des semaines
    Set celluleSemaine = plageSemaines.Find("2023-" & numeroSemaine, LookIn:=xlValues)
    
    ' Si la cellule avec la semaine est trouv_e, s_lectionner la cellule
    If Not celluleSemaine Is Nothing Then
        celluleSemaine.Select
    Else
        MsgBox "Semaine non trouvée."
    End If
End Sub

