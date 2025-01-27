Sub update_tar()
    Dim userInput As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetCell As Range
    
    ' Set the worksheet to the desired sheet
    Set ws = ThisWorkbook.Sheets("calculs_intermediaires")
    
    ' Call the function and store the result
    userInput = PopUpAndInputWithConfirmation()
    
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
    targetCell.Offset(0, 3).FormulaR1C1 = "=RC[-1]+R4C17"
    
    ' nb
    targetCell.Offset(0, 4).FormulaR1C1 = "=RC[-1]-RC[-2]+1"
    
    ' update consideration
    targetCell.Offset(0, 5).Value = "True"
    targetCell.Offset(-1, 5).Value = "False"
    
    ' Save the workbook
    ActiveWorkbook.Save
End Sub