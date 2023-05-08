Function Merge_cells(Target As Range, Optional delim As String, Optional Empty_Cells As Boolean = False)
  
Application.Volatile
Dim merged As String

For Each c In Target

    If Empty_Cells = False And c.Value = "" Then GoTo nextcell
    
    merged = merged & delim & CStr(c.Value)

nextcell:
Next

Merge_cells = merged

End Function
