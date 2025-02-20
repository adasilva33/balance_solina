Sub clean_data_brute()
    Dim ws As Worksheet
    Dim currentSheet As Worksheet

    ' Store the currently active sheet
    Set currentSheet = activeSheet

    ' Reference the "data_brute" sheet without activating it
    Set ws = ThisWorkbook.Sheets("data_brute")

    ' Clear contents of columns B:H from row 2 down to the last row
    ws.Range("B2:H" & ws.Rows.Count).ClearContents

    ' Restore the user's previously active sheet
    currentSheet.Activate
End Sub