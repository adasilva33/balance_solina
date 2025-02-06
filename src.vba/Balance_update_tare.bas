Sub update_tar()
    Dim userInput As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetCell As Range
    
    ' Set the worksheet to the desired sheet
    Set ws = ThisWorkbook.Sheets("calculs_intermediaires")
    
    ' Call the function and store the result
    Call DisplayCellTextWithConfirmation("pop_up", "B3")
    userInput = PopUpAndInputWithConfirmation("pop_up", "B4", "B5", 1)
    
    ' Check if the user canceled
    If IsError(userInput) Then Exit Sub
    
    ' Determine the last row of the table starting from M6
    lastRow = ws.Range("M6").End(xlDown).Row
    
    ' bobine (next row in column M)
    Set targetCell = ws.Cells(lastRow + 1, "M")
    targetCell.FormulaR1C1 = "=R[-1]C+1"
    
    ' tare (input from user)
    targetCell.Offset(0, 1).Value = userInput
    
    ' ligne debut
    targetCell.Offset(0, 2).FormulaR1C1 = "=R[-1]C[1]+1"
    targetCell.Offset(-1, 2).Copy
    targetCell.Offset(-1, 2).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' ligne fin
    targetCell.Offset(-1, 3).Copy
    targetCell.Offset(-1, 3).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    targetCell.Offset(0, 3).FormulaR1C1 = "=RC[-1]+R4C17"
    
    ' nb
    targetCell.Offset(0, 4).FormulaR1C1 = "=RC[-1]-RC[-2]+1"
    
    ' update consideration
    targetCell.Offset(0, 5).Value = "True"
    targetCell.Offset(-1, 5).Value = "False"
    
    ' Save the workbook
    ActiveWorkbook.save
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