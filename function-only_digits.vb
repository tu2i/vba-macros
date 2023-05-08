Function OnlyDigits(tx As String, Optional nms As String = ":+0123456789")
  
Application.Volatile
Dim cleartx As String
'nms = ":+0123456789"

    For i = 1 To Len(tx)
        If InStr(1, nms, Mid(tx, i, 1), vbTextCompare) Then cleartx = cleartx & Mid(tx, i, 1)
    Next i

OnlyDigits = cleartx

End Function
