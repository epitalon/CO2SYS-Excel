     

Option Base 1



Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If SaveAsUI = False Then
          answer = InputBox("Please, use ""Save As..."" instead.")
          If (answer <> "1967") Then
                Cancel = True
          Else
                Cancel = False
          End If
    End If
End Sub

Private Sub Workbook_Open()
        
        InitiateOK = False
        Call Initiate
        Sheets("INPUT").Activate

End Sub




