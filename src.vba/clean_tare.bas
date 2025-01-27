Sub clean_tare()

    Range("M8").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("P7").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]+R[-3]C[-2]-1"
    Range("R7").Select
    ActiveCell.FormulaR1C1 = "TRUE"
    Range("P8").Select
End Sub
'