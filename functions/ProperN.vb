Function ProperN(ByVal ref As Range) As String
Application.Volatile
    Dim vaArray As Variant
    Dim c As String
    Dim i As Integer
    Dim j As Integer
    Dim vaLCase As Variant
    Dim str As String

    ' Array contains terms that should be lower case
    vaLCase = Array("a", "an", "and", "in", "is", _
      "of", "or", "the", "to", "with")

    c = StrConv(ref, 3)
    'split the words into an array
    vaArray = Split(c, " ")
    For i = (LBound(vaArray) + 1) To UBound(vaArray)
        For j = LBound(vaLCase) To UBound(vaLCase)
            ' compare each word in the cell against the
            ' list of words to remain lowercase. If the
            ' Upper versions match then replace the
            ' cell word with the lowercase version.
            If UCase(vaArray(i)) = UCase(vaLCase(j)) Then
                vaArray(i) = vaLCase(j)
            End If
        Next j
    Next i

  ' rebuild the sentence
    str = ""
    For i = LBound(vaArray) To UBound(vaArray)
        str = str & " " & vaArray(i)
    Next i

    ProperN = Trim(str)
End Function
