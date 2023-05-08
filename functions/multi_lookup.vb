Function MultiLookup(FindText As Range, CompareTable As Range, ColumnOffset As Integer, Optional Delimiter As String = vbCrLf)
Application.Volatile
result = ""
LookFor = FindText.Value

'clean the input. Sometimes excel incorrectly handles line breaks which can be VbCrLf or Chr(10)&Chr(13)
'convert all delimiters to a rare character "Chr(135)"
LookFor = Replace(LookFor, Delimiter, Chr(135))
LookFor = Replace(LookFor, vbCrLf, Chr(135))
LookFor = Replace(LookFor, Chr(10), Chr(135))
LookFor = Replace(LookFor, Chr(13), Chr(135))

'remove repeating delimiters

Do While InStr(1, LookFor, Chr(135) & Chr(135)) > 1
    LookFor = Replace(LookFor, Chr(135) & Chr(135), Chr(135))
Loop

'Split and convert into an array
LookForA = Split(LookFor, Chr(135))

For i = 0 To UBound(LookForA)
    For Each c In CompareTable.Cells
        If LookForA(i) = c.Value Then
            If result = "" Then
                result = c.Offset(0, ColumnOffset).Value
            Else
                result = result & Delimiter & c.Offset(0, ColumnOffset).Value
            End If
        End If
        
    Next c
Next i

MultiLookup = result

End Function
