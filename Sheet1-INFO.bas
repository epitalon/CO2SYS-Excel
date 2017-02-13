



Option Base 1


Private Sub Worksheet_SelectionChange(ByVal Target As Range)


If Target.Row < 2 Or Target.Row > 1 + UBound(mess) Or Target.Column > 1 Then Exit Sub


With Sheets("INFO").Range(Cells(2, 1), Cells(1 + UBound(mess), 1))
    .Interior.ColorIndex = xlNone
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .WrapText = True
End With


With Target.Interior
    .ColorIndex = 36
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
End With

Call AboutCO2SYS(mess())
   
   Sheets("INFO").Range("C2").Value = mess(Target.Row - 1)
   
   Erase mess

End Sub
