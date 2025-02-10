Sub AdjustYAxis()
    Dim chartObj As ChartObject
    Dim ws As Worksheet
    Dim minY As Double, maxY As Double
    Dim dataRange As Range
    Dim cell As Range

    ' Set your worksheet
    Set ws = activeSheet
    
    ' Define your data range (Adjust the columns to match your dataset)
    Set dataRange = Sheets("data_brute").Range("M:R") ' Modify based on your columns
    
    ' Get the minimum and maximum values
    minY = Application.WorksheetFunction.Min(dataRange)
    maxY = Application.WorksheetFunction.Max(dataRange)

    ' Loop through all charts in the worksheet
    For Each chartObj In ws.ChartObjects
        With chartObj.Chart.Axes(xlValue) ' Y-Axis
            .MinimumScale = minY
            .MaximumScale = maxY
        End With
    Next chartObj
End Sub