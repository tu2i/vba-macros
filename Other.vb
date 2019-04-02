Sub automatic_calculation()
'
' auto_calc Macro
'

'
c = MsgBox("Do you want to enable automatic calculation?", vbQuestion + vbYesNoCancel, "Automatic calculation")
If c = vbYes Then Application.Calculation = xlCalculationAutomatic
If c = vbNo Then Application.Calculation = xlCalculationManual
End Sub
