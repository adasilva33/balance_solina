Sub UpdateChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets("data_brute") ' Adjust the sheet name

    ' Find the last row with data
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Update Chart Series
    Set chartObj = ws.ChartObjects("Chart 1") ' Adjust chart name
    With chartObj.Chart.SeriesCollection(1)
        .XValues = ws.Range("A2:A" & lastRow) ' Update X-axis
        .Values = ws.Range("B2:B" & lastRow) ' Update Y-axis
    End With
End Sub