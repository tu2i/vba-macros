Sub merge_cells_in_range()
'
' merge_cells_in_range Macro
' merge_cells_in_range
'
' Keyboard Shortcut: Ctrl+m
'
Dim rng As Range
Dim txt As String
Dim wrp As Boolean
Set rng = Selection
wrp = rng.Cells(1, 1).WrapText
    
    For Each cell In rng
        If txt <> "" Then txt = txt & vbCrLf & cell.Value Else txt = cell.Value
        cell.Clear
    Next cell

rng.Cells(1, 1).Value = txt
rng.Cells(1, 1).Select
rng.Cells(1, 1).WrapText = wrp

'Call fix_line_breaks
'Application.OnKey "^+m", "fix_line_breaks"  '^crtl, +shift, %alt

End Sub

Sub fix_line_breaks()

valid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ÖöÜüÐðÞþÇçÝý ,"

Set rng = Selection

For Each cell In rng
    txt = cell.Value
    txt2 = txt
        
    txt = Replace(txt, Chr(10), Chr(135))
    txt = Replace(txt, Chr(13), Chr(135))
    txt = Replace(txt, Chr(135) & Chr(135) & Chr(135) & Chr(135), vbCrLf)
    txt = Replace(txt, Chr(135) & Chr(135) & Chr(135), vbCrLf)
    txt = Replace(txt, Chr(135) & Chr(135), vbCrLf)
    txt = Replace(txt, Chr(135), vbCrLf)
    txt = Replace(txt, Chr(13) & Chr(10) & Chr(13) & Chr(10), vbCrLf)
    txt = Replace(txt, Chr(13) & Chr(10) & Chr(13) & Chr(10), vbCrLf)
    txt = Replace(txt, Chr(13) & Chr(10) & Chr(13) & Chr(10), vbCrLf)
    txt = Replace(txt, Chr(13) & Chr(10) & Chr(13) & Chr(10), vbCrLf)
        
    
    i = 2
    Do While i < Len(txt)  'properly format lines
        If Mid(txt, i, 2) = Chr(13) & Chr(10) And InStr(valid, Mid(txt, i - 1, 1)) > 0 And InStr(valid, Mid(txt, i + 2, 1)) > 0 Then
            txt = RTrim(Left(txt, i - 1)) & " " & LTrim(Mid(txt, i + 2, Len(txt)))
        Else
            i = i + 1
        End If

    Loop
         
    
    If txt <> txt2 Then cell.Value = txt
    cell.WrapText = True
    
Next cell

End Sub
