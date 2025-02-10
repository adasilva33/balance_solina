Sub UpdateYAxis()
    Dim chartObj As ChartObject
    Dim wsChart As Worksheet, wsData As Worksheet
    Dim chart As chart
    Dim minY As Double, maxY As Double
    Dim dataRange As Range
    Dim cell As Range
    Dim valuesArray As Variant
    Dim tempValue As Double

    ' Define the worksheets
    Set wsChart = Sheets("interface") ' Sheet where the chart is located
    Set wsData = Sheets("data_brute") ' Sheet where the data is stored

    ' Define the data range (M:R) in "data_brute"
    Set dataRange = wsData.Range("M:R") ' Adjust if needed to a specific row range, e.g., M1:R100

    ' Initialize min and max with extreme values
    minY = 1E+30 ' Large initial value for comparison
    maxY = -1E+30 ' Small initial value for comparison

    ' Loop through each cell in the range to find min and max
    For Each cell In dataRange
        If IsNumeric(cell.Value) Then
            If Not IsError(cell.Value) Then ' Ignore errors like #N/A
                tempValue = cell.Value
                If tempValue < minY Then minY = tempValue
                If tempValue > maxY Then maxY = tempValue
            End If
        End If
    Next cell

    ' Ensure valid min/max before applying
    If minY < maxY Then
        ' Find the chart object in the "interface" sheet
        For Each chartObj In wsChart.ChartObjects
            Set chart = chartObj.chart
            
            ' Update Y-axis
            With chart.Axes(xlValue)
                .MinimumScale = minY
                .MaximumScale = maxY
            End With
        Next chartObj
    Else
        MsgBox "No valid numerical data found in M:R range!", vbExclamation, "Y-Axis Update Failed"
    End If
End Sub