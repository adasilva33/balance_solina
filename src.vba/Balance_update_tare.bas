Sub update_tar()
    

    Call DisplayCellTextWithConfirmation("pop_up", "B3")
    Call DisplayCellTextWithConfirmation("pop_up", "B4")
    Call DisplayCellTextWithConfirmation("pop_up", "B5")
    
    Call MoveCursorToPopUpLastRow2
End Sub

Function MoveCursorToPopUpLastRow2()
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