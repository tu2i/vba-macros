Sub paste_values()
'
' paste_values Macro
' Paste_values_only
'
' Keyboard Shortcut: Ctrl+q
'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub num_format()
'
' number_formatt Macro
'
' Keyboard Shortcut: Ctrl+t
'
    If Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 " Then
    Selection.NumberFormat = "#,##0_ ;[Red]-#,##0 "
    Else
    Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    End If
End Sub
Sub leave_values()
'
' leave_valuess Macro
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Sub CenterAcrossSelection()
'
' CenterAcrossSelection Macro
'

'
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .MergeCells = False
    End With
End Sub
