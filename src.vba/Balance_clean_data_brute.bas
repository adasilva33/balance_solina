Sub clean_data_brute()
    ' Select columns B:H starting from row 2 and clear their contents
    Range("B2:H" & Rows.Count).ClearContents
    ' Move the selection back to cell A1
    Range("A1").Select
End Sub