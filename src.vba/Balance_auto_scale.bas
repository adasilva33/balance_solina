Sub AutoScale()
    Dim wsChart As Worksheet, wsData As Worksheet
    Dim cht As ChartObject
    Dim minY As Double, maxY As Double

    ' Define the correct sheets
    Set wsChart = Sheets("interface") ' Sheet where the chart is located
    Set wsData = Sheets("calculs_intermediaires") ' Sheet where min/max values are stored

    ' Define the chart object correctly
    Set cht = wsChart.ChartObjects("Chart 2") ' Replace with actual chart name

    ' Retrieve min and max values safely
    minY = wsData.Range("BU8").Value
    maxY = wsData.Range("BU9").Value

    ' Ensure valid min/max values before applying them
    If IsNumeric(minY) And IsNumeric(maxY) And minY < maxY Then
        With cht.chart.Axes(xlValue)
            .MinimumScale = minY
            .MaximumScale = maxY
        End With
    Else
        MsgBox "Invalid min/max values in BU8/BU9!", vbExclamation, "Error"
    End If
End Sub