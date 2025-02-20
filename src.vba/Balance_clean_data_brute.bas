Sub clean_data_brute()
    ' Select columns B:H starting from row 2 and clear their contents
    Sheets("data_brute").Activate
    Range("B2:H" & Rows.Count).ClearContents
    
End Sub