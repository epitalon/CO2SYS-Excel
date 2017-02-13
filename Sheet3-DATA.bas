




Private Sub Worksheet_Activate()
  Cells(2, 1).Activate
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim r1 As Range


If (Target.Address <> "$A$1:$L$1" And Target.Address <> "$M$3" And Target.Address <> "$M$5:$M$7") Then Exit Sub

If Target.Address = "$A$1:$L$1" Then
   Call main
ElseIf Target.Address = "$M$3" Then
   answer = MsgBox("Are you sure you want to clear the data ?", vbYesNo + vbQuestion + vbDefaultButton2, "Caution")
   Set r1 = Range(Cells(4, "A"), Cells(65356, "L"))
   If answer = vbYes Then r1.ClearContents
Else
   answer = MsgBox("Are you sure you want to clear the results ?", vbYesNo + vbQuestion + vbDefaultButton2, "Caution")
   nrows = Sheets("DATA").UsedRange.Rows.Count
   Set r1 = Range(Cells(4, "Q"), Cells(nrows, "BF"))
   nrows = Sheets("ERROR").UsedRange.Rows.Count
   Set r2 = Sheets("ERROR").Range(Sheets("ERROR").Cells(4, "N"), Sheets("ERROR").Cells(nrows, "AM"))
   If answer = vbYes Then
      r1.ClearContents
      ' Setting/clearing format by program makes the total number of row impssible to decrease !
      ' r1.ClearFormats
      r2.ClearContents
   End If
End If


End Sub


