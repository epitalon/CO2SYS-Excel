



Option Base 1


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

If Target.Column > 6 Or Target.Row = 1 Or (Target.Row > 1 + UBound(kopt) And Target.Column = 1) _
                                Or (Target.Row > 1 + UBound(khso4opt) And Target.Column = 2) _
                                Or (Target.Row > 1 + UBound(kfopt) And Target.Column = 3) _
                                Or (Target.Row > 1 + UBound(phopt) And Target.Column = 4) _
                                Or (Target.Row > 1 + UBound(tbopt) And Target.Column = 5) _
                                Or (Target.Row > 1 + UBound(EOSopt) And Target.Column = 6) _
                                Then Exit Sub
If Target.Cells.Count > 1 Then mmm = MsgBox("Select 1 Cell Only!", vbOKOnly, "Selection Error!"): Exit Sub

If Target.Column = 1 Then
   With Sheets("INPUT").Range(Cells(2, 1), Cells(1 + UBound(kopt), 1))
      .Interior.ColorIndex = xlNone
      .HorizontalAlignment = xlLeft
      .VerticalAlignment = xlCenter
      .WrapText = True
   End With
   WhichKs% = Target.Row - 1
   If WhichKs% = 6 Or WhichKs% = 7 Then
      Sheets("input").Range("C5").Select
   End If
ElseIf Target.Column = 2 Then
   With Sheets("INPUT").Range(Cells(2, 2), Cells(1 + UBound(khso4opt), 2))
      .Interior.ColorIndex = xlNone
      .HorizontalAlignment = xlLeft
      .VerticalAlignment = xlCenter
      .WrapText = True
   End With
   WhoseKSO4% = Target.Row - 1
ElseIf Target.Column = 3 Then
   With Sheets("INPUT").Range(Cells(2, 3), Cells(1 + UBound(kfopt), 3))
      .Interior.ColorIndex = xlNone
      .HorizontalAlignment = xlLeft
      .VerticalAlignment = xlCenter
      .WrapText = True
   End With
   WhoseKF% = Target.Row - 1
ElseIf Target.Column = 4 Then
   With Sheets("INPUT").Range(Cells(2, 4), Cells(1 + UBound(phopt), 4))
      .Interior.ColorIndex = xlNone
      .HorizontalAlignment = xlLeft
      .VerticalAlignment = xlCenter
      .WrapText = True
   End With
   pHScale% = Target.Row - 1
ElseIf Target.Column = 5 Then
   With Sheets("INPUT").Range(Cells(2, 5), Cells(1 + UBound(tbopt), 5))
      .Interior.ColorIndex = xlNone
      .HorizontalAlignment = xlLeft
      .VerticalAlignment = xlCenter
      .WrapText = True
   End With
   WhichTB% = Target.Row - 1
ElseIf Target.Column = 6 Then
   With Sheets("INPUT").Range(Cells(2, 6), Cells(1 + UBound(EOSopt), 6))
      .Interior.ColorIndex = xlNone
      .HorizontalAlignment = xlLeft
      .VerticalAlignment = xlCenter
      .WrapText = True
   End With
   WhichEOS% = Target.Row - 1
End If

With Target.Interior
    .ColorIndex = 36
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
End With
    

End Sub


