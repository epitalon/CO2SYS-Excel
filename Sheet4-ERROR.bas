




Private Sub Worksheet_Activate()
  Cells(2, 1).Activate
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim r1 As Range


If (Target.Address <> "$K$3" And Target.Address <> "$K$5:$K$7" And Target.Address <> "$L$4:$M$4") Then Exit Sub

If Target.Address = "$K$3" Then
   answer = MsgBox("Are you sure you want to clear the errorS ?", vbYesNo + vbQuestion + vbDefaultButton2, "Caution")
   Set r1 = Range(Cells(4, "A"), Cells(65356, "J"))
   If answer = vbYes Then r1.ClearContents
ElseIf Target.Address = "$L$4:$M$4" Then
'     Set default values of errors for dissociation constants (pK) and Total Boron
'     Default values for epK are:
'        pK0   :  0.002   CO2 solubility
'        pK1   :  0.0075  Carbonate dissociation constants
'        pK2   :  0.015
'        pKb   :  0.01    Borate
'        pKw   :  0.01    Water dissociation
'        pKspa :  0.02    solubility product of Aragonite
'        pKspc :  0.02    solubility product of Calcite
'        TB    :  0.02    Total boron
'
   Cells(5, "M").Value = 0.002
   Cells(7, "M").Value = 0.0075
   Cells(9, "M").Value = 0.015
   Cells(11, "M").Value = 0.01
   Cells(13, "M").Value = 0.01
   Cells(15, "M").Value = 0.02
   Cells(17, "M").Value = 0.02
   Cells(19, "M").Value = 0.02
Else
   answer = MsgBox("Are you sure you want to clear the resulting errors ?", vbYesNo + vbQuestion + vbDefaultButton2, "Caution")
   Set r1 = Range(Cells(4, "O"), Cells(65356, "AV"))
   If answer = vbYes Then r1.ClearContents
End If


End Sub

