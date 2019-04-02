Function OnlyDigits(tx As String)
Dim nms, cleartx As String
nms = ":+0123456789"

    For i = 1 To Len(tx)
        If InStr(1, nms, Mid(tx, i, 1), vbTextCompare) Then cleartx = cleartx & Mid(tx, i, 1)
    Next i

OnlyDigits = cleartx

End Function

Function Merge_cells(Target As Range, Optional delim As String, Optional Empty_Cells As Boolean = False)
Dim merged As String

For Each c In Target

    If Empty_Cells = False And c.Value = "" Then GoTo nextcell
    
    merged = merged & delim & CStr(c.Value)

nextcell:
Next

Merge_cells = merged

End Function

Function ProperN(ByVal ref As Range) As String
    Dim vaArray As Variant
    Dim c As String
    Dim i As Integer
    Dim J As Integer
    Dim vaLCase As Variant
    Dim str As String

    ' Array contains terms that should be lower case
    vaLCase = Array("a", "an", "and", "in", "is", _
      "of", "or", "the", "to", "with")

    c = StrConv(ref, 3)
    'split the words into an array
    vaArray = Split(c, " ")
    For i = (LBound(vaArray) + 1) To UBound(vaArray)
        For J = LBound(vaLCase) To UBound(vaLCase)
            ' compare each word in the cell against the
            ' list of words to remain lowercase. If the
            ' Upper versions match then replace the
            ' cell word with the lowercase version.
            If UCase(vaArray(i)) = UCase(vaLCase(J)) Then
                vaArray(i) = vaLCase(J)
            End If
        Next J
    Next i

  ' rebuild the sentence
    str = ""
    For i = LBound(vaArray) To UBound(vaArray)
        str = str & " " & vaArray(i)
    Next i

    ProperN = Trim(str)
End Function

Function ColorCode(Target As Range) As Integer
Dim temp As Integer

temp = Target.Interior.ColorIndex

If temp = -4142 Then
    ColorCode = 0
Else
    ColorCode = temp
End If

End Function

Function MultiLookup(FindText As Range, CompareTable As Range, ColumnOffset As Integer, Optional Delimiter As String = vbCrLf)
result = ""
Source = FindText.Value
Source = Replace(Source, Delimiter, Chr(135))
Source = Replace(Source, vbCrLf, Chr(135))
Source = Replace(Source, Chr(10), Chr(135))
Source = Replace(Source, Chr(13), Chr(135))
Source = Replace(Source, Chr(135) & Chr(135), Chr(135))
Source = Replace(Source, Chr(135) & Chr(135), Chr(135))
Source = Replace(Source, Chr(135) & Chr(135), Chr(135))
Source = Replace(Source, Chr(135) & Chr(135), Chr(135))
Source = Replace(Source, Chr(135) & Chr(135), Chr(135))

Sourcea = Split(Source, Chr(135))

For i = 0 To UBound(Sourcea)
    For Each c In CompareTable.Cells
        If Sourcea(i) = c.Value Then
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
