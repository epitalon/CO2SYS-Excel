' Program CO2SYS.BAS, version 01.05, 10-15-97, written by Ernie Lewis.
' This is a new version combining CO2SYSTM, FCO2TCO2, PHTCO2, and CO2BTCH.
' For more information, see the sub AboutCO2SYS.
'
'
' ***********************************************
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' PROGRAMMER'S NOTE: This program is DANGEROUSLY close to the DOS-imposed
'       64K limit due to all the print statements in the sub AboutCO2SYS.
'       Don't make any unnecessary changes or the limit will be exceeded.
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' PROGRAMMER'S NOTE: all logs are base e, any log10 is written log()/log(10)
' PROGRAMMER'S NOTE: all temps are deg C unless otherwise noted -
'       temps in deg K only occur in the subs and are expicitly noted
' PROGRAMMER'S NOTE: partials are calculated numerically and there will be
'       some roundoff error involved in this, but it should be small
' PROGRAMMER'S NOTE: pCO2 and fCO2 are both referenced to wet air. In an
'       earlier version I had xCO2 in dry air as a variable with pTot
'       assumed to be 1 atm (so essentially I had pCO2 in dry air), thus
'       there is some code that could be removed now if I chose to do so.
'       FugFac does not change with TempC very much, whereas VPFac = (1-pH2O)
'       did, so I could put it as a constant, but I left the code as it was.
' PROGRAMMER'S NOTE: the constants are converted to the chosen pH scale and
'       calculations are made on that scale. Some of the subs are designed
'       for the total scale, but for reasonable pH (>6) they will work fine.
' PROGRAMMER'S NOTE: the statement:
'       IF WhichKs% = 7 THEN TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
'       is so the TA value used is the correct one for the case used (Peng or
'       Dickson); the program is coded for TA(Dickson) in the calculation subs.
'
'
'****************************************************************************
'
'***************** variables ***********
Option Base 1
Public Const OutputNum = 10000
Public mess(1 To 14) As String
Public K(10), K0  ' these are the equilibrium constants
Public T(5):  ' these are the amounts of the various species
Public c0, FugFac
Dim datar As Range, varinpr As Range
Public param(5) As Single, VarInp(7) As Single, parami(2) As Integer
Dim VarTemp(11) As Double, VarTempF(11) As String
Dim SubFlagT As String, SubFlag(3) As String, SubFlagok As Integer

Dim datac(OutputNum, 40) ', datacC(OutputNum, 40) As Long
Public WhichKsDefault%, WhoseKSO4Default%, pHScaleDefault%, WhichTBDefault%
Public WhichKs%, WhoseKSO4%, pHScale%, WhichTB%
'Public TCO2m, fCO2m
Public phopt(4) As String, kopt(14) As String, khso4opt(2) As String, tbopt(2) As String
Public InitiateOK As Boolean

'****************************************************************************
'****************************************************************************
Sub main()
  If Not InitiateOK Then
    answer = MsgBox("Macro has not Initialized Properly!" + Chr$(LF) + "Please Close and Re-Open!" + Chr$(LF) + "Your INPUT Data Will Not Be Saved!", vbOKOnly, "Initialization Error!")
    Exit Sub
  End If
  
  Sheets("DATA").Activate
  Range("$A$3").Select: pt = 1: pt2 = 1: pt3 = 0
  
  If ActiveCell.Value <> "Salinity" Then
     answer = MsgBox("Sheet Layout was Changed...Contact CDIAC", vbOKOnly + vbCritical, "ERROR")
     Exit Sub
  End If
  answer = MsgBox("Is Your Data Properly Entered on the Worksheet?", vbDefaultButton1 + vbYesNo + vbExclamation, "Check Point")
  If answer = vbNo Then Exit Sub
  
  SkipAData = MsgBox("Calculate ""Auxiliary Data""?", vbDefaultButton2 + vbYesNo + vbExclamation, "Check Point")
  If SkipAData = vbYes Then
     SkipAData = vbNo
  ElseIf SkipAData = vbNo Then
     SkipAData = vbYes
  End If
  
 ' g3init = Range("$G$3").Value
 ' f3init = Range("$F$3").Value
  
c0 = 14
If pt = 1 Then
   For i = 2 To UBound(kopt)
      If Sheets("INPUT").Cells(i, 1).Interior.ColorIndex = 36 Then WhichKs% = i - 1
   Next i
   For i = 2 To UBound(khso4opt)
      If Sheets("INPUT").Cells(i, 2).Interior.ColorIndex = 36 Then WhoseKSO4% = i - 1
   Next i
   For i = 2 To UBound(phopt)
      If Sheets("INPUT").Cells(i, 3).Interior.ColorIndex = 36 Then pHScale% = i - 1
   Next i
   For i = 2 To UBound(tbopt)
      If Sheets("INPUT").Cells(i, 4).Interior.ColorIndex = 36 Then WhichTB% = i - 1
   Next i

 ' pHScale% = Sheets("INPUT").phcombo.ListIndex + 1
 ' WhichKs% = Sheets("INPUT").kcombo.ListIndex + 1
 ' WhoseKSO4% = Sheets("INPUT").khso4combo.ListIndex + 1
  ActiveCell.Offset(1, c0).Value = phopt(pHScale%)
  ActiveCell.Offset(3, c0).Value = kopt(WhichKs%)
  ActiveCell.Offset(5, c0).Value = khso4opt(WhoseKSO4%)
  ActiveCell.Offset(7, c0).Value = tbopt(WhichTB%)
End If
c0 = c0 + 1

startofvar = 0: endofvar = 6: startofdata = 7: endofdata = 11
Set varinpr = Range(ActiveCell.Offset(pt, startofvar), ActiveCell.Offset(pt, endofvar))
Set datar = Range(ActiveCell.Offset(pt, startofdata), ActiveCell.Offset(pt, endofdata))
stilldata = WorksheetFunction.CountA(datar) + WorksheetFunction.CountA(varinpr)

Erase datac ' Each element set to 0.
SubFlag(1) = "Pressure Set to 0. ": SubFlag(2) = "Phosp Set to 0. ": SubFlag(3) = "Si Set to 0. "

Do While (stilldata <> 0) 'loop for each point
       SubFlagok = 0: SubFlagT = ""
       If WorksheetFunction.CountA(datar) = 0 Then GoTo endmain
'          Range(ActiveCell.Offset(pt, 15), ActiveCell.Offset(pt, 35)).Value = -9
'          Range(ActiveCell.Offset(pt, 37), ActiveCell.Offset(pt, 52)).Value = -9
  '        GoTo endmain
'       End If
       kk = 1: c0 = 15: parami(1) = 0: parami(2) = 0
       For i = 1 To 7  'inputs var (S, T, P, Si, Ph) into VarInp
                       ' and Data (TA, TC, pH, pCO2) into param
          If i <= 5 Then
             param(i) = datar.Value2(1, i)
             If datar.Value2(1, i) <> Empty And kk <= 2 Then  ' And kk <= 2
                parami(kk) = i: kk = kk + 1
                If param(i) < 0 Then GoTo endmain
 '                 Range(ActiveCell.Offset(pt, 15), ActiveCell.Offset(pt, 35)).Value = -9
 '                 Range(ActiveCell.Offset(pt, 37), ActiveCell.Offset(pt, 52)).Value = -9
 '                 GoTo endmain
 '              End If
             End If
             VarInp(i) = varinpr.Value2(1, i)
          Else
             VarInp(i) = varinpr.Value2(1, i)
          End If
          If (VarInp(i) = -9 Or VarInp(i) = -999) Then
             If (i <= 2) Then 'if S or T = -999
                GoTo endmain
             Else 'if Press, P or Si = -999 then =0
                VarInp(i) = 0
                SubFlagok = 1
                SubFlagT = SubFlagT + SubFlag(i - 2)
             End If
          End If
       Next i
  '     TCO2m = datar.Value2(1, 2)
  '     fCO2m = datar.Value2(1, 4)
       Sal = VarInp(1): TempC = VarInp(2): Pdbar = VarInp(3)
       If WhichKs% = 8 Or WhichKs% = 6 Then   'GEOSECS and WATER
          VarInp(4) = 0: VarInp(5) = 0  'TP and TSi =0
          If WhichKs% = 8 Then  'Pure Water
             Sal = 0!
          End If
       Else
          VarInp(4) = VarInp(4) / 1000000!: VarInp(5) = VarInp(5) / 1000000!:
       End If
       T(4) = VarInp(4): T(5) = VarInp(5)
       Call Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
       K1 = K(1): K2 = K(2)
      ICase% = 10 * parami(1) + parami(2)
       
      If parami(1) = 4 Or parami(2) = 4 Then
           param(5) = param(4) / FugFac
      ElseIf parami(1) = 5 Or parami(2) = 5 Then
           param(4) = param(5) * FugFac
       End If
      For i = 1 To 5
          If i = 3 Then i = 4
          param(i) = param(i) / 1000000!
      Next i
       
CalculateOtherParamsAtInputConditions:
        Select Case ICase%
        Case 12: ' input TA, TC
                TA = param(1): TC = param(2)
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                param(3) = pH: param(4) = fCO2: param(5) = param(4) / FugFac
        Case 13: ' input TA, pH
                TA = param(1): pH = param(3)
                If WhichKs% = 7 Then TA = TA - T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
                Call CalculateTCfromTApH(TA, pH, K(), T(), TC)
                param(2) = TC
                If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
                Call CalculatefCO2fromTCpH(TC, pH, K0, K1, K2, fCO2): pCO2 = fCO2 / FugFac
                param(4) = fCO2: param(5) = pCO2
        Case 14, 15: ' input TA, fCO2 or pCO2
                TA = param(1)
                If ICase% = 14 Then
                    fCO2 = param(4): pCO2 = fCO2 / FugFac
                Else
                    pCO2 = param(5): fCO2 = pCO2 * FugFac
                End If
                If WhichKs% = 7 Then TA = TA - T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
                Call CalculatepHfromTAfCO2(TA, fCO2, K0, K(), T(), pH)
                Call CalculateTCfromTApH(TA, pH, K(), T(), TC)
                If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
                param(3) = pH: param(2) = TC
        Case 23: ' input TC, pH
                TC = param(2): pH = param(3)
                Call CalculateTAfromTCpH(TC, pH, K(), T(), TA)
                If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
                Call CalculatefCO2fromTCpH(TC, pH, K0, K1, K2, fCO2): pCO2 = fCO2 / FugFac
                param(1) = TA: param(4) = fCO2: param(5) = pCO2
        Case 24, 25: ' input TC, fCO2 or pCO2
                TC = param(2)
                If ICase% = 24 Then
                   fCO2 = param(4): pCO2 = fCO2 / FugFac
                Else
                   pCO2 = param(5): fCO2 = pCO2 * FugFac
                End If
                Call CalculatepHfromTCfCO2(TC, fCO2, K0, K1, K2, pH)
                If pH = -999! Then
                     param(1) = -999! / 1000000!
                Else
                   Call CalculateTAfromTCpH(TC, pH, K(), T(), TA)
                   If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
                End If
                param(1) = TA: param(3) = pH
        Case 34, 35: ' input pH, fCO2 or pCO2
                If ICase% = 35 Then
                   pCO2 = param(5): fCO2 = pCO2 * FugFac
                Else
                   fCO2 = param(4): pCO2 = fCO2 / FugFac
                End If
                pH = param(3)
                Call CalculateTCfrompHfCO2(pH, fCO2, K0, K1, K2, TC)
                Call CalculateTAfromTCpH(TC, pH, K(), T(), TA)
                If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
                param(2) = TC: param(1) = TA
        Case Else
                GoTo endmain
        End Select
 '----------------------------- 'Print Session #1 ------------------------------------
  GoTo skipP1
        For i = 1 To 5
             If (i < 4) Then
                ActiveCell.Offset(pt, c0 + i).Value = VarInp(i)
             Else
                ActiveCell.Offset(pt, c0 + i).Value = VarInp(i) * 1000000
             End If
             If i = 3 Then
                ActiveCell.Offset(pt, c0 + 5 + i).NumberFormat = "#0.000"
             Else
               param(i) = param(i) * 1000000!
               ActiveCell.Offset(pt, c0 + 5 + i).NumberFormat = "#0.0"
             End If
             ActiveCell.Offset(pt, c0 + 5 + i).Value = param(i)
        Next i
        GoTo afterP1
skipP1:
    datac(pt2, 1) = VarInp(1): datac(pt2, 2) = VarInp(2): datac(pt2, 3) = VarInp(3)
    datac(pt2, 4) = VarInp(4) * 1000000: datac(pt2, 5) = VarInp(5) * 1000000
    datac(pt2, 1 + 5) = param(1) * 1000000!: datac(pt2, 2 + 5) = param(2) * 1000000!: datac(pt2, 3 + 5) = param(3)
    datac(pt2, 4 + 5) = param(4) * 1000000!: datac(pt2, 5 + 5) = param(5) * 1000000!
afterP1:
        'pHinp = pH: fCO2inp = fCO2: pCO2inp = pCO2
        pHinp = param(3): fCO2inp = param(4): pCO2inp = param(5)
       ' c0 = 13
        c0 = c0 + 2 * (i - 1)

If SkipAData = vbYes Then
    c0 = c0 + 12    'separator
    GoTo CalculatepHfCO2AtOutputConditions
End If
'****************************************************************************
CalculateOtherStuffAtInputConditions:
        Call CalculateAlkParts(pH, TC, K(), T(), HCO3, CO3, BAlk, OH, PAlk, SiAlk, Hfree, HSO4, HF)
        If WhichKs% = 7 Then PAlk = PAlk + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
                HCO3inp = HCO3: CO3inp = CO3: CO2inp = TC - CO3inp - HCO3inp
                BAlkinp = BAlk: OHinp = OH: PAlkinp = PAlk: SiAlkinp = SiAlk
        Call RevelleFactor(WhichKs%, TA, TC, K0, K(), T(), Revelle)
                Revelleinp = Revelle
        K1 = K(1): K2 = K(2)
        Call CaSolubility(WhichKs%, Sal, TempC, Pdbar, TC, pH, K1, K2, OmegaCa, OmegaAr)
                OmegaCainp = OmegaCa: OmegaArinp = OmegaAr
                xCO2dryinp = pCO2inp / VPFac: ' this assumes pTot = 1 atm
          i = 1
                VarTemp(i) = HCO3inp * 1000000!: VarTempF(i) = "#0.0": i = i + 1
                VarTemp(i) = CO3inp * 1000000!: VarTempF(i) = "#0.0": i = i + 1
                VarTemp(i) = CO2inp * 1000000!: VarTempF(i) = "#0.0": i = i + 1
                VarTemp(i) = BAlkinp * 1000000!: VarTempF(i) = "#0.0": i = i + 1
                VarTemp(i) = OHinp * 1000000!: VarTempF(i) = "#0.0": i = i + 1
                VarTemp(i) = PAlkinp * 1000000!: VarTempF(i) = "#0.0": i = i + 1
                VarTemp(i) = SiAlkinp * 1000000!: VarTempF(i) = "#0.0": i = i + 1
                VarTemp(i) = Revelleinp: VarTempF(i) = "#0.000": i = i + 1
                VarTemp(i) = OmegaCainp: VarTempF(i) = "#0.00": i = i + 1
                VarTemp(i) = OmegaArinp: VarTempF(i) = "#0.00": i = i + 1
                VarTemp(i) = xCO2dryinp * 1000000!: VarTempF(i) = "#0.0": i = i + 1
  '----------------------------- 'Print Session #2 ------------------------------------
 GoTo skipP2
      For i = 1 To 11
         ActiveCell.Offset(pt, c0 + i).Value = VarTemp(i)
         ActiveCell.Offset(pt, c0 + i).NumberFormat = VarTempF(i)
      Next i
      c0 = c0 + i    'separator
      GoTo afterP2
skipP2:
      For i = 1 To 11
         datac(pt2, 10 + i) = VarTemp(i)
      Next i
afterP2:
'
'****************************************************************************
CalculatepHfCO2AtOutputConditions:
        'TempC = TempCout: Pdbar = Pdbarout
        TempC = VarInp(6): Pdbar = VarInp(7)
        Call Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
        K1 = K(1): K2 = K(2)
       TA = param(1): TC = param(2): pH = param(3): fCO2 = param(4): pCO2 = param(5)
        If WhichKs% = 7 Then TA = TA - T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatepHfromTATC(TA, TC, K(), T(), pH)
        If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatefCO2fromTCpH(TC, pH, K0, K1, K2, fCO2): pCO2 = fCO2 / FugFac
'
        pHout = pH: fCO2out = fCO2: pCO2out = pCO2
        VarTemp(1) = pHout: VarTemp(2) = fCO2out * 1000000!: VarTemp(3) = pCO2out * 1000000!
 '----------------------------- 'Print Session #3 ------------------------------------
  GoTo skipP3
      For i = 1 To 3
         If i <= 2 Then ActiveCell.Offset(pt, c0 + i).Value = VarInp(5 + i)
         If i = 1 Then
            ActiveCell.Offset(pt, c0 + 2 + i).NumberFormat = "#0.000"
         Else
            ActiveCell.Offset(pt, c0 + 2 + i).NumberFormat = "#0.0"
         End If
        ActiveCell.Offset(pt, c0 + 2 + i).Value = VarTemp(i)
      Next i
      c0 = c0 + i - 1 + 2
        GoTo afterP3
skipP3: 'datac( ,22)=separator on sheet
     
    datac(pt2, 23) = VarInp(6): datac(pt2, 24) = VarInp(7)
    datac(pt2, 25) = VarTemp(1): datac(pt2, 26) = VarTemp(2): datac(pt2, 27) = VarTemp(3)
afterP3:
'
If SkipAData = vbYes Then GoTo endmain
'****************************************************************************
CalculateOtherStuffAtOutputConditions:
        Call CalculateAlkParts(pH, TC, K(), T(), HCO3, CO3, BAlk, OH, PAlk, SiAlk, Hfree, HSO4, HF)
        If WhichKs% = 7 Then PAlk = PAlk + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
                HCO3out = HCO3: CO3out = CO3: CO2out = TC - CO3out - HCO3out
                BAlkout = BAlk: OHout = OH: PAlkout = PAlk: SiAlkout = SiAlk
        Call RevelleFactor(WhichKs%, TA, TC, K0, K(), T(), Revelle)
                Revelleout = Revelle
        Call CaSolubility(WhichKs%, Sal, TempC, Pdbar, TC, pH, K1, K2, OmegaCa, OmegaAr)
                OmegaCaout = OmegaCa: OmegaArout = OmegaAr
        xCO2dryout = pCO2out / VPFac: ' this assumes pTot = 1 atm
        
     ' VarTemp(1) = HCO3out: VarTemp(2) = CO3out: VarTemp(3) = CO2out
      ' VarTemp(4) = BAlkout: VarTemp(5) = OHout: VarTemp(6) = PAlkout: VarTemp(7) = SiAlkout
       'VarTemp(8) = Revelleout
       'VarTemp(9) = OmegaCaout: VarTemp(10) = OmegaArout: VarTemp(11) = xCO2dryout
        i = 1
        VarTemp(i) = HCO3out * 1000000!: VarTempF(i) = "#0.0": i = i + 1
        VarTemp(i) = CO3out * 1000000!: VarTempF(i) = "#0.0": i = i + 1
        VarTemp(i) = CO2out * 1000000!: VarTempF(i) = "#0.0": i = i + 1
        VarTemp(i) = BAlkout * 1000000!: VarTempF(i) = "#0.0": i = i + 1
        VarTemp(i) = OHout * 1000000!: VarTempF(i) = "#0.0": i = i + 1
        VarTemp(i) = PAlkout * 1000000!: VarTempF(i) = "#0.0": i = i + 1
        VarTemp(i) = SiAlkout * 1000000!: VarTempF(i) = "#0.0": i = i + 1
        VarTemp(i) = Revelleout: VarTempF(i) = "#0.000": i = i + 1
        VarTemp(i) = OmegaCaout: VarTempF(i) = "#0.00": i = i + 1
        VarTemp(i) = OmegaArout: VarTempF(i) = "#0.00": i = i + 1
        VarTemp(i) = xCO2dryout * 1000000!: VarTempF(i) = "#0.0": i = i + 1
 '----------------------------- 'Print Session #4 ------------------------------------
  GoTo skipP4
      For i = 1 To 11
         ActiveCell.Offset(pt, c0 + i).Value = VarTemp(i)
         ActiveCell.Offset(pt, c0 + i).NumberFormat = VarTempF(i)
      Next i
      c0 = c0 + i - 1
        GoTo afterP4
skipP4:
      For i = 1 To 11
         datac(pt2, 27 + i) = VarTemp(i)
      Next i

afterP4:
'
GoTo endmain
'****************************************************************************
DoPartialsHere:
        'Call PrintHeader(ICase%, pHScale%, fORp$, TA, TC, pHinp, fCO2inp, pCO2inp, TP, TSi, Sal, TempCinp, Pdbarinp, TempCout, Pdbarout)
        Select Case ICase%
        Case 12: ' input TA, TC
                Call Case1Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TA, TC, pHinp, fCO2inp, pCO2inp, pHout, fCO2out, pCO2out)
        Case 13: ' input TA, pH
                Call Case2Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TA, pHinp, TC, fCO2inp, pCO2inp, pHout, fCO2out, pCO2out)
        Case 14, 15: ' input TA, fCO2 or pCO2
                Call Case3Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TA, fCO2inp, pCO2inp, TC, pHinp, pHout, fCO2out, pCO2out)
        Case 23: ' input TC, pH
                Call Case4Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TC, pHinp, TA, fCO2inp, pCO2inp, pHout, fCO2out, pCO2out)
        Case 24, 25: ' input TC, fCO2 or pCO2
                TCfCO2Flag% = 0
                Call Case5Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TC, fCO2inp, pCO2inp, TA, pHinp, pHout, fCO2out, pCO2out, TCfCO2Flag%)
                If TCfCO2Flag% = 1 Then
                        TCfCO2Flag% = 0
                       ' Call PrintTCfCO2Warning
                       ' GoTo Start:
                End If
        Case 34, 35: ' input pH, fCO2 or pCO2
                Call Case6Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, pHinp, fCO2inp, pCO2inp, TA, TC, pHout, fCO2out, pCO2out)
        End Select

endmain:
'       Range("$F$3").Value = pt
 '      Range("$G$3").Value = pt2
      If SubFlagok = 1 Then
         datac(pt2, 40) = SubFlagT
         'For i = 1 To 40
         '   datacC(pt2, i) = 3  'red
         'Next i
      Else
         'For i = 1 To 40
         '   datacC(pt2, i) = 0  'auto
         'Next i
      End If
       
       Application.StatusBar = "Points Calculated: Total " + Str(pt) + "  This Batch " + Str(pt2)
      pt = pt + 1: pt2 = pt2 + 1
       If (pt2 > OutputNum) Then
         Application.StatusBar = "Points Calculated: Total " + Str(pt) + "  This Batch " + Str(pt2) + ".  Formating Data... "
         Range(ActiveCell.Offset(pt3 + 1, 16), ActiveCell.Offset(pt3 + pt2 - 1, 53)).Value = datac
         Range(ActiveCell.Offset(pt3 + 1, 16), ActiveCell.Offset(pt3 + pt2 - 1, 55)).Font.ColorIndex = datacC
         Erase datac
         pt3 = pt3 + pt2 - 1
         pt2 = 1
       End If
       Set varinpr = Range(ActiveCell.Offset(pt, startofvar), ActiveCell.Offset(pt, endofvar))
       Set datar = Range(ActiveCell.Offset(pt, startofdata), ActiveCell.Offset(pt, endofdata))
       stilldata = WorksheetFunction.CountA(datar) + WorksheetFunction.CountA(varinpr)
       If (stilldata = 0) Then
           Application.StatusBar = "Points Calculated: Total " + Str(pt) + "  This Batch " + Str(pt2) + ".  Formating Data... "
           Range(ActiveCell.Offset(pt3 + 1, 16), ActiveCell.Offset(pt3 + pt2 - 1, 53)).Value = datac
           For i = pt3 + 1 To pt3 + pt2 - 1
              If datac(i, 40) <> "" Then
                 Range(ActiveCell.Offset(i, 16), ActiveCell.Offset(i, 55)).Font.ColorIndex = 3
              End If
           Next i
       End If

   Loop 'end of do while stilldata
  
  'Range("$F$3").Value = f3init
  'Range("$G$3").Value = g3init
  Application.StatusBar = ""
End Sub

Sub AboutCO2SYS(mess() As String)
' SUB AboutCO2SYS, version 03.04, 10-15-97, written by Ernie Lewis.
' Inputs: none
' Outputs: none
' This prints information about the program CO2SYS.
If InStr(1, Application.OperatingSystem, "macintosh", vbTextCompare) > 0 Then
  LF = 13
Else
  LF = 10
End If

AboutMacro:
        mess(1) = ""
        mess(1) = mess(1) + "The code for this Macro was taken directly from Ernie Lewis' ""CO2SYS.BAS"" Basic Program.  "
        mess(1) = mess(1) + "(See the ""INFO"" sheet in the macro for contact information)." + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "What it Does…" + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "     From two known CO2 parameters (TA, TCO2, pH, pCO2 or fCO2), the program will calculate the other 3,"
        mess(1) = mess(1) + "as well as other quantities such as Omega, Revelle Factor, Carbonate species concentrations…(referred to as ""Auxiliary Data"" in the macro)."
        mess(1) = mess(1) + "The quantities can be calculated at 2 different sets of T and P conditions (IN and OUT)" + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "What it Doesn't do…" + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "     Unlike CO2Sys.BAS, this macro does not calculate the sensitivity of the output on the input"
        mess(1) = mess(1) + "(referred to as ""Partials"" in the original program)." + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "HOW TO RUN THE CO2Sys MACRO." + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "   In Sheet ""INFO"":" + Chr$(LF)
        mess(1) = mess(1) + "      -- You can select which section of the program you want information on by selecting the appropriate option"
        mess(1) = mess(1) + " from the ""Subject"" column on the left. The information will be listed in this text box." + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "   In Sheet ""INPUT"":" + Chr$(LF)
        mess(1) = mess(1) + "      The program gives you a number of options. When selected, the option is highlighted in yellow." + Chr$(LF)
        mess(1) = mess(1) + "      -- Select the set of CO2 constants you want to use for the calculations." + Chr$(LF)
        mess(1) = mess(1) + "      -- Select the KHSO4." + Chr$(LF)
        mess(1) = mess(1) + "      -- Select the pH scale of your data." + Chr$(LF)
         mess(1) = mess(1) + "     -- Select the Total Boron formulation you want." + Chr$(LF)
       mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "   In Sheet ""Data"":" + Chr$(LF)
        mess(1) = mess(1) + "      --  If copying from another Excel file, it is suggested to only paste the VALUES in the cells." + Chr$(LF)
        mess(1) = mess(1) + "      --  Input your data in the appropriate columns for Salinity, Temperature (oC) and Pressure "
        mess(1) = mess(1) + "(dbars). Total Si and Total P (in  umol/kg SW) are optional." + Chr$(LF)
        mess(1) = mess(1) + "      --  Input the CO2 parameters in their respective columns. If more than two are entered, "
        mess(1) = mess(1) + "the FIRST TWO from the left will be used. You may use different sets of parameters in different rows." + Chr$(LF)
        mess(1) = mess(1) + "      --  Set the output conditions at which you want your results (optional)." + Chr$(LF)
        mess(1) = mess(1) + "      --  Click the red ""Start"" Button located on the top left part of the ""Data"" sheet." + Chr$(LF)
        mess(1) = mess(1) + "      --  Calculations will stop when an entire row of data (columns ""A"" to ""L"") is empty." + Chr$(LF)
        mess(1) = mess(1) + "      --  You can either clear your data (columns ""A"" to ""L"") or clear the results (columns ""L"" to the end) "
        mess(1) = mess(1) + "by clicking on the appropriate button located on top of the column ""M"" in this sheet." + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "   After the program starts:" + Chr$(LF)
        mess(1) = mess(1) + "      --  You will be asked if you entered your data properly…this gives you a chance to cancel your action." + Chr$(LF)
        mess(1) = mess(1) + "      --  You will be asked if you want to calculate the ""Auxiliary Data"". This corresponds to Omega, "
        mess(1) = mess(1) + "Revelle Factor…etc…any column right of the pCO2 column in both the ""Input Conditions"" "
        mess(1) = mess(1) + "and the ""Output Conditions"" sections. Choosing ""No"" will save time." + Chr$(LF)
        mess(1) = mess(1) + "      --  Results at the ""Input Conditions"" are posted in columns ""Q"" to ""AK"" "
        mess(1) = mess(1) + "and are labeled ""in"". Those at the ""Output Conditions"" are posted in columns ""AM"" to ""BB"" "
        mess(1) = mess(1) + "and are labeled ""out""." + Chr$(LF)
        mess(1) = mess(1) + "      --  If Pressure, Total P or Total Si are missing, equal to -999, or -9, "
        mess(1) = mess(1) + "their value is set to zero and the calculation performed anyway. In this case, the whole corresponding "
        mess(1) = mess(1) + "row is colored in red and column ""BC"" (labeled ""SubFlag"" ) will mention which input parameter "
        mess(1) = mess(1) + "was set to zero." + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "Any problem or comment related to the code of this macro can be addressed to:" + Chr$(LF)
        mess(1) = mess(1) + "       Denis Pierrot" + Chr$(LF)
        mess(1) = mess(1) + "       CIMAS" + Chr$(LF)
        mess(1) = mess(1) + "       University of Miami" + Chr$(LF)
        mess(1) = mess(1) + "       4600 Rickenbacker Causeway" + Chr$(LF)
        mess(1) = mess(1) + "       Miami, FL 33149" + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)
        mess(1) = mess(1) + "       dpierrot@rsmas.miami.edu  or  denis.pierrot@noaa.gov" + Chr$(LF)
        mess(1) = mess(1) + Chr$(LF)




'****************************************************************************
Generalp1:
        'mess=""
        mess(2) = mess(2) + "This macro is a reproduction of the CO2SYS Program, version 01.05, "
        mess(2) = mess(2) + "written by Ernie Lewis in Visual Basic format, with a few options added." + Chr$(LF)
        mess(2) = mess(2) + "The information below is taken verbatim from that program." + Chr$(LF)
        mess(2) = mess(2) + Chr$(LF)
        mess(2) = mess(2) + "Program CO2SYS , version 01.05, written by Ernie Lewis. " + Chr$(LF)
        mess(2) = mess(2) + Chr$(LF)
        mess(2) = mess(2) + "This program takes two parameters of the CO2 system in seawater (TA, TC, "
        mess(2) = mess(2) + "pH, fCO2 or pCO2), and calculates the other two at a set of input "
        mess(2) = mess(2) + "conditions (T and P) and a set of output conditions chosen by the user. "
        mess(2) = mess(2) + "It supersedes the 1995 programs CO2SYSTM, FCO2TCO2, PHTCO2, and CO2BTCH. " + Chr$(LF)
        mess(2) = mess(2) + Chr$(LF)
        mess(2) = mess(2) + "For questions, comments, or to report any problems, please contact: " + Chr$(LF)
        mess(2) = mess(2) + "     Ernie Lewis or Doug Wallace " + Chr$(LF)
        mess(2) = mess(2) + "     Department of Applied Science " + Chr$(LF)
        mess(2) = mess(2) + "     Building 318 " + Chr$(LF)
        mess(2) = mess(2) + "     P. O. Box 5000 " + Chr$(LF)
        mess(2) = mess(2) + "     Brookhaven National Laboratory " + Chr$(LF)
        mess(2) = mess(2) + "     Upton, NY 11973-5000 " + Chr$(LF)
        mess(2) = mess(2) + Chr$(LF)
        mess(2) = mess(2) + "     elewis@bnl.gov     wallace@bnl.gov " + Chr$(LF)
        mess(2) = mess(2) + "     516-344-7406       516-344-2945 " + Chr$(LF)
        mess(2) = mess(2) + Chr$(LF)
        mess(2) = mess(2) + "This work was supported by the US Department of Energy Office of Health and "
        mess(2) = mess(2) + "Enviromental Research under contract DE-ACO2-76CH00016, through a project "
        mess(2) = mess(2) + "entitled `Inorganic Carbon for the World Ocean Circulation Experiment - "
        mess(2) = mess(2) + "World Hydrographic Program' (D.W.R. Wallace and K.M. Johnson, PIs). " + Chr$(LF)
'
'***************************************
Generalp2:
        mess(2) = mess(2) + Chr$(LF)
        mess(2) = mess(2) + "Every effort has been made to make this program as correct, complete, and "
        mess(2) = mess(2) + "user-friendly as possible. HOWEVER, the program is not failsafe and some "
        mess(2) = mess(2) + "familiarity with the CO2 system in seawater is assumed. " + Chr$(LF)
        mess(2) = mess(2) + Chr$(LF)
        mess(2) = mess(2) + "The effects of phosphate, silicate, and OH are included, as well as the "
        mess(2) = mess(2) + "non-ideality of CO2. Some programs we have evaluated do not include these, "
        mess(2) = mess(2) + "which can have a significant effect on the results. " + Chr$(LF)
        mess(2) = mess(2) + Chr$(LF)
        mess(2) = mess(2) + "In developing this program, much work was done to ensure that correct "
        mess(2) = mess(2) + "values for the various constants were used. There is a paucity of data "
        mess(2) = mess(2) + "for many of the values. Many errors were found in the literature. Whenever "
        mess(2) = mess(2) + "possible these were corrected or otherwise noted. A listing is included "
        mess(2) = mess(2) + "in the accompanying documentation. " + Chr$(LF)
        mess(2) = mess(2) + Chr$(LF)
        mess(2) = mess(2) + "This program allows for a variety of options, including: " + Chr$(LF)
        mess(2) = mess(2) + "     choice of various formulations for K1 and K2, " + Chr$(LF)
        mess(2) = mess(2) + "     two distinct formulations for KSO4 (Dickson's or Khoo's), " + Chr$(LF)
        mess(2) = mess(2) + "     choice of four pH scales (free, total, seawater, or NBS), " + Chr$(LF)
        mess(2) = mess(2) + "     use of either fugacity (fCO2) or partial pressure (pCO2) of CO2, " + Chr$(LF)
        mess(2) = mess(2) + "     and choice of any two CO2 system parameters as inputs. " + Chr$(LF) + Chr$(LF)
'
'***************************************
Generalp3:
        mess(2) = mess(2) + "INPUT: " + Chr$(LF)
        mess(2) = mess(2) + "     the salinity, " + Chr$(LF)
        mess(2) = mess(2) + "     the input temperature and pressure (or depth), " + Chr$(LF)
        mess(2) = mess(2) + "     the concentrations of silicate and phosphate, " + Chr$(LF)
        mess(2) = mess(2) + "     the two known CO2 system parameters at the input conditions. " + Chr$(LF)
        mess(2) = mess(2) + "     the output temperature and pressure (or depth), " + Chr$(LF) + Chr$(LF)
        mess(2) = mess(2) + "OUTPUT: " + Chr$(LF)
        mess(2) = mess(2) + "The program will calculate the other two CO2 system parameters at the input "
        mess(2) = mess(2) + "conditions. TA and TC, which do not vary with temperature and pressure, "
        mess(2) = mess(2) + "are used to calculate the pH and fCO2 (or pCO2) at the output conditions. " + Chr$(LF) + Chr$(LF)
        mess(2) = mess(2) + "AUXILIARY DATA: " + Chr$(LF)
        mess(2) = mess(2) + "Also calculated for both the input and the output conditions are: " + Chr$(LF)
        mess(2) = mess(2) + "     contributions to the alkalinity and carbon speciation,  " + Chr$(LF)
        mess(2) = mess(2) + "     fCO2 and pCO2, " + Chr$(LF)
        mess(2) = mess(2) + "     omega (the degree of saturation) for calcite and for aragonite, " + Chr$(LF)
        mess(2) = mess(2) + "     the Revelle, or homogeneous buffer, factor, " + Chr$(LF)
        mess(2) = mess(2) + "     pH values on all four pH scales, " + Chr$(LF)
        mess(2) = mess(2) + "     the values of pK1, pK2, pKW, and pKB. " + Chr$(LF)
'Return
'****************************************************************************
AboutpHScalesp1:
        mess(3) = ""
        mess(3) = mess(3) + "The various pH scales are inter-related by the following equations: " + Chr$(LF)
        mess(3) = mess(3) + Chr$(LF)
        mess(3) = mess(3) + "            -pHNBS     " + Chr$(LF)
        mess(3) = mess(3) + "aH = 10               = fH * Hsws" + Chr$(LF)
        mess(3) = mess(3) + Chr$(LF)
        mess(3) = mess(3) + "                     Htot                           Hsws         " + Chr$(LF)
        mess(3) = mess(3) + "   Hfree  =  ------------------  =  ------------------------------ " + Chr$(LF)
        mess(3) = mess(3) + "               1 + TS/KSO4         1 + TS/KSO4 + TF/KF " + Chr$(LF)
        mess(3) = mess(3) + Chr$(LF)
        mess(3) = mess(3) + "where aH is the activity and fH the activity coefficient of the H+ ion "
        mess(3) = mess(3) + "(this includes liquid junction effects), TS and TF are the concentrations "
        mess(3) = mess(3) + "of SO4- and fluorine, and KSO4 and KF are the dissociation constants of "
        mess(3) = mess(3) + "HSO4 and HF in seawater. " + Chr$(LF)
        '(which are inherently on the free scale)
        mess(3) = mess(3) + Chr$(LF)
        mess(3) = mess(3) + "These conversions depend on temperature, salinity, and pressure. "
        mess(3) = mess(3) + "At 20 deg C, Sal 35, and 1 atm, pH values on the total scale are (about) " + Chr$(LF)
        mess(3) = mess(3) + "     .09 units lower than those on the free scale, " + Chr$(LF)
        mess(3) = mess(3) + "     .01 units higher than those on the seawater scale, and " + Chr$(LF)
        mess(3) = mess(3) + "     .13 units lower than those on the NBS scale. " + Chr$(LF)
        mess(3) = mess(3) + Chr$(LF)
        mess(3) = mess(3) + "The concentration units for aH on the NBS scale are mol/kg-H2O. " + Chr$(LF)
        mess(3) = mess(3) + "The concentration units used here for [H] on the other scales is mol/kg-SW " + Chr$(LF)
        mess(3) = mess(3) + "(note that the free scale was originally defined in units of mol/kg-H2O). " + Chr$(LF)
        mess(3) = mess(3) + "The difference between mol/kg-SW and mol/kg-H2O is about .015 pH units "
        mess(3) = mess(3) + "at salinity 35 (the difference is nearly proportional to salinity). " + Chr$(LF) + Chr$(LF)
'
'***************************************
AboutpHScalesp2:
        mess(3) = mess(3) + "The seawater scale was formerly referred to as the total scale, and "
        mess(3) = mess(3) + "each is still sometimes referred to as the other in the literature. "
        mess(3) = mess(3) + "The fit of fH used here is valid from salinities 20 to 40. "
        mess(3) = mess(3) + "fH has been found to be electrode-dependent, and does NOT equal 1 at "
        mess(3) = mess(3) + "salinity 0 due to the liquid junction potential. " + Chr$(LF)
        mess(3) = mess(3) + "Values on the NBS pH scale are only accurate to (at best) .005. " + Chr$(LF)
        mess(3) = mess(3) + "All work on pressure effects on pH has assumed that fH is independent "
        mess(3) = mess(3) + "of pressure. Some of the pH scale conversions depend on pressure. " + Chr$(LF)
        mess(3) = mess(3) + Chr$(LF)
        mess(3) = mess(3) + "For discussions of the various pH scales, see: " + Chr$(LF)
        mess(3) = mess(3) + "   -Dickson, Deep-Sea Research 40:107-118, 1993, " + Chr$(LF)
        mess(3) = mess(3) + "   -Millero, Marine Chemistry 44:143-152, 1993, " + Chr$(LF)
        mess(3) = mess(3) + "   -Dickson, Geochemica et Cosmochemica Acta 48:2299-2308, 1984, " + Chr$(LF)
        mess(3) = mess(3) + "   -Butler, Marine Chemistry 38:251-282, 1992, " + Chr$(LF)
        mess(3) = mess(3) + "   -Culberson, C. H., Direct Potentiometry, Chapter 6 (pp. 187-261), in: " + Chr$(LF)
        mess(3) = mess(3) + "           Marine Electrochemistry, eds. M. Whitfield and D. Jagner, 1981. " + Chr$(LF)
        mess(3) = mess(3) + Chr$(LF)
        mess(3) = mess(3) + "Attention is required because in some of these papers the distinction between the "
        mess(3) = mess(3) + "total and seawater pH scales was not made. "

'Return
'****************************************************************************
AboutfCO2pCO2:
        mess(4) = ""
        mess(4) = mess(4) + "The fugacity of CO2 (fCO2) in water is defined to be the fugacity of CO2 "
        mess(4) = mess(4) + "in wet (100% water-saturated) air which is in equilibrium with the water. " + Chr$(LF)
        mess(4) = mess(4) + "pCO2, the partial pressure of CO2, is defined to be the product of the "
        mess(4) = mess(4) + "mole fraction of CO2 in WET air and the total pressure. This is the "
        mess(4) = mess(4) + "same as the product of the the mole fraction of CO2 in DRY air (xCO2(dry)) "
        mess(4) = mess(4) + "and (pTot - pH2O), where pH2O is the vapor pressure of water above seawater."
        mess(4) = mess(4) + "At pressures of order 1 atm fCO2 in air is about .3% lower than the pCO2 due"
        mess(4) = mess(4) + "to the non-ideality of CO2 (see Weiss, Marine Chemistry 2:203-215, 1974). " + Chr$(LF)
        mess(4) = mess(4) + "This program assumes a pressure near 1 atm (where most equilibrators "
        mess(4) = mess(4) + "function) for the conversion between partial pressure and fugacity. " + Chr$(LF)
        mess(4) = mess(4) + Chr$(LF)
        mess(4) = mess(4) + "fCO2 is related to TC and pH by the following equation: " + Chr$(LF)
        mess(4) = mess(4) + "                [CO2*]     TC                H*H          " + Chr$(LF)
        mess(4) = mess(4) + "     fCO2 =  ------  =  ---- *  --------------------------- " + Chr$(LF)
        mess(4) = mess(4) + "                    K0        K0      H*H + K1*H + K1*K2 " + Chr$(LF)
        mess(4) = mess(4) + "where [CO2*] is the concentration of dissolved CO2, K0 is the solubility "
        mess(4) = mess(4) + "coefficient of CO2 in seawater, and K1 and K2 are the first and second "
        mess(4) = mess(4) + "dissociation constants for carbonic acid in seawater. " + Chr$(LF)
        mess(4) = mess(4) + Chr$(LF)
        mess(4) = mess(4) + "Units for fCO2 and pCO2 in this program are uatm (micro-atmospheres). " + Chr$(LF)
        mess(4) = mess(4) + "The value of xCO2(dry) given in this program assumes pTot = 1 atmosphere. " + Chr$(LF)
        mess(4) = mess(4) + "GEOSECS and Peng et al did not distinguish between fCO2 and pCO2, nor did "
        mess(4) = mess(4) + "some other programs that we have evaluated. "
        '''''''''''''''''''''''''''''''''''''
        ' pCO2(wet) = xCO2(wet) * pTot = xCO2(dry) * (pTot - VPSW)
        ' where VPSW is the vapor pressure of water above seawater
        ' fCO2(wet) = pCO2(wet) * FugFac
        ' (pTot - VPSW) converts from wet air to dry air
        ' FugFac converts partial pressure to fugacity
        '''''''''''''''''''''''''''''''''''''
'Return
'****************************************************************************
AboutKSO4:
        mess(5) = ""
        mess(5) = mess(5) + "KSO4 is defined to be the dissociation constant for the reaction " + Chr$(LF)
        mess(5) = mess(5) + "     HSO4- = H+   +    SO4--, " + Chr$(LF)
        mess(5) = mess(5) + "thus KSO4 = [H] * [SO4] / [HSO4]. " + Chr$(LF)
        mess(5) = mess(5) + Chr$(LF)
        mess(5) = mess(5) + "Two formulations of this are still in current usage: " + Chr$(LF)
        mess(5) = mess(5) + "Khoo et al, Analytical Chemistry, 49(1):29-34, 1977, and " + Chr$(LF)
        mess(5) = mess(5) + "Dickson, Journal of Chemical Thermodynamics, 22:113-127, 1990. " + Chr$(LF)
        mess(5) = mess(5) + Chr$(LF)
        mess(5) = mess(5) + "The values of Dickson are now recommended, though many older papers used "
        mess(5) = mess(5) + "values of Khoo et al. They are between 15 to 45 % lower than those of "
        mess(5) = mess(5) + "Dickson, depending on temperature (mostly). " + Chr$(LF)
        mess(5) = mess(5) + Chr$(LF)
        mess(5) = mess(5) + "The main effect of this difference will occur when converting from one "
        mess(5) = mess(5) + "pH scale to another, or when working on a scale for which equilibrium "
        mess(5) = mess(5) + "constants must be converted (e.g., most constants were determined on "
        mess(5) = mess(5) + "either the total scale or the seawater scale). " + Chr$(LF)
        mess(5) = mess(5) + Chr$(LF)
        mess(5) = mess(5) + "Use of the Dickson values when converting from the total pH scale to the "
        mess(5) = mess(5) + "free pH scale will result in pH values which are .015 to .03 units lower "
        mess(5) = mess(5) + "than those obtained using values of Khoo et al. "
'Return
'****************************************************************************
'****************************************************************************
AboutFreshwaterOption:
        mess(6) = ""
        mess(6) = mess(6) + "For the freshwater option only [HCO3], [CO3], [OH], and [H] are included "
        mess(6) = mess(6) + "in the definition of alkalinity: TA = [HCO3] + 2[CO3] + [OH] - [H]. " + Chr$(LF)
        mess(6) = mess(6) + Chr$(LF)
        mess(6) = mess(6) + "fH, the activity coefficient of H+, does NOT equal 1 at salinity 0 due "
        mess(6) = mess(6) + "to liquid junction effects (included in its definition). It is also "
        mess(6) = mess(6) + "found to be electrode dependent. Thus, while the values of pH on the "
        mess(6) = mess(6) + "free, total, and seawater scales will coincide at salinity 0, the value "
        mess(6) = mess(6) + "on the NBS scale will differ. For these reasons, for this choice only a "
        mess(6) = mess(6) + "pH value is given without reference to a pH scale. " + Chr$(LF)
        
'         "Only one set of measurements of K1 and K2 has been made in seawater at "
 '        "salinity < 10. Though the values can be extrapolated to salinity 0 they "
  '       "change by a considerable amount over this interval (between salinities 0 "
   '      "and 5, K1 varies by a factor of 2 and K2 by between 6.5 and 9.2, depending "
    '     "on temperature). For comparison, between salinities 5 and 35 K1 varies by "
     '    "a factor of less than 1.5 and K2 less than 3). Thus a fit of K1 and K2 for "
      '   "values of salinity in this range would be prone to large uncertainty. For "
       '  "this reason, only values of K1 and K2 valid at salinity 0 are used."
        '
        mess(6) = mess(6) + "Constants used for this choice (K1, K2, and KW) are from: " + Chr$(LF)
        mess(6) = mess(6) + "     Millero, F. J., Geochemica et Cosmochemica Acta 43:1651-1661, 1979. " + Chr$(LF)
        mess(6) = mess(6) + "Pressure effects on these constants are from: " + Chr$(LF)
        mess(6) = mess(6) + "     Millero, Chap. 43, Chemical Oceanography, ed. Riley + Chester, 1983. " + Chr$(LF)
        '''''''''''''''''''''''''''''''
        ' "Further, it is inherent in the determination of K1 and K2 that the "
        ' "     seawater ratios of constituents occur, which is unlikely, "
        ' "     making it hard to define what is meant by salinity. "
        ' at 20 deg, K1(5) / K1(0) = 2.0, K1(35) / K1(5) = 1.4
        '            K2(5) / K2(0) = 8.5, K2(35) / K2(5) = 2.8
        '''''''''''''''''''''''''''''''
'Return
'****************************************************************************
AboutGEOSECSp1:
        mess(7) = ""
        mess(7) = mess(7) + "The GEOSECS option was designed to replicate the calculations performed "
        mess(7) = mess(7) + "in Chapter 3, Carbonate Chemistry, by Takahashi et al, in GEOSECS Pacific "
        mess(7) = mess(7) + "Expedition, Volume 3, by Broecker et al, 1982. " + Chr$(LF)
        mess(7) = mess(7) + Chr$(LF)
        mess(7) = mess(7) + "That work used the NBS pH scale, the values of K1 and K2 from Mehrbach "
        mess(7) = mess(7) + "et al, and the value of KB from Lyman. It did not include effects of OH, "
        mess(7) = mess(7) + "silicate, or phosphate, nor was there a correction for the non-ideality "
        mess(7) = mess(7) + "of CO2 (i.e., implying fCO2 and pCO2 are the same). Their boron "
        mess(7) = mess(7) + "concentration was about 1% lower than that used for the other choices in "
        mess(7) = mess(7) + "this program (except the choice of Peng). " + Chr$(LF)
        mess(7) = mess(7) + Chr$(LF)
        mess(7) = mess(7) + "In GEOSECS, TA and TC values from titration were used to determine pCO2, "
        mess(7) = mess(7) + "[H2CO3], [HCO3-], [CO3--], and pH, at P = 1 atm and insitu T; and "
        mess(7) = mess(7) + "[H2CO3], [HCO3-], [CO3--], aH, pH, ICP, and delta CO3-- for calcite and "
        mess(7) = mess(7) + "aragonite at insitu T and P, where aH = 10^(-pH), ICP = [Ca++][CO3--], and "
        mess(7) = mess(7) + "delta CO3-- is the difference between [CO3--] and its saturation level. " + Chr$(LF)
        mess(7) = mess(7) + Chr$(LF)
        mess(7) = mess(7) + "These last three parameters were used to describe the saturation states of "
        mess(7) = mess(7) + "calcite and aragonite. In this program only omegas, dimensionless ratios, "
        mess(7) = mess(7) + "are output for this. " + Chr$(LF) + Chr$(LF)
'
'***************************************
AboutGEOSECSp2:
        mess(7) = mess(7) + "A fit for fH was also given (for salinities 20 to 40) and is used to "
        mess(7) = mess(7) + "convert between pH scales in this program. " + Chr$(LF)
        mess(7) = mess(7) + Chr$(LF)
        mess(7) = mess(7) + "Some typographic errors in the GEOSECS report were noted and corrected: " + Chr$(LF)
        mess(7) = mess(7) + "     in the pressure dependence of K2 the given value 26.4 should be 16.4, "
        mess(7) = mess(7) + "     and the expression for ln KW should have C*ln T, not C/ln T. " + Chr$(LF)
        mess(7) = mess(7) + "That these are correct can be seen by checking the original references. " + Chr$(LF)
        mess(7) = mess(7) + Chr$(LF)
        mess(7) = mess(7) + "The ratio of Ksp(arag.) / Ksp(calc.) is given as 1..48 in the original "
        mess(7) = mess(7) + "reference (Berner, R. A., American Journal of Science 276:713-730, 1976), "
        mess(7) = mess(7) + "but the value of 1.45 given in GEOSECS was used both in that work and in "
        mess(7) = mess(7) + "this program as well for this choice. " + Chr$(LF)
        mess(7) = mess(7) + Chr$(LF)
        mess(7) = mess(7) + "The GEOSECS report also contains a discussion on the effects of OH, "
        mess(7) = mess(7) + "phosphate, and silicate (see pp. 79-82, especially Table 1 on p. 81, of "
        mess(7) = mess(7) + "Chapter 3, Carbonate Chemistry, by Takahashi et al, in GEOSECS Pacific "
        mess(7) = mess(7) + "Expedition, V. 3, by Broecker et al, 1982). From this, it can be seen how "
        mess(7) = mess(7) + "important these can be, especially for calculated values of fCO2 (or pCO2). "
        mess(7) = mess(7) + "This table has a typo: 17.8 for Aw in Pacific Surface Water should be 7.8. " + Chr$(LF)
        mess(7) = mess(7) + Chr$(LF)
        mess(7) = mess(7) + "The choice of Peng is very similar, and should be used instead if the "
        mess(7) = mess(7) + "values of OH, etc. are desired with these constants. "
'Return
'****************************************************************************
AboutPeng:
        mess(8) = mess(8) + "This choice replicates the calculation scheme of Peng et al, Tellus 39B: "
        mess(8) = mess(8) + "439-458, 1987, which is similar to GEOSECS. Peng et al worked on the NBS "
        mess(8) = mess(8) + "pH scale and included effects of phosphate, silicate, and OH, but did not "
        mess(8) = mess(8) + "distinguish between fCO2 and pCO2. The values of K1 and K2 from Mehrbach et "
        mess(8) = mess(8) + "al and the value of KB from Lyman were used. " + Chr$(LF)
        mess(8) = mess(8) + Chr$(LF)
        mess(8) = mess(8) + "They did not treat calcite and aragonite solubility or pressure effects, "
        mess(8) = mess(8) + "but these are included in this program for this choice using GEOSECS values "
        mess(8) = mess(8) + "for solubility and pressure dependence of K1, K2, and KB, and the same "
        mess(8) = mess(8) + "values for the pressure dependence of OH and phosphate and silicate "
        mess(8) = mess(8) + "dissociation as are used in constant choices 1 to 5. The concentration "
        mess(8) = mess(8) + "of boron they used was about 1% lower than that used for other choices in "
        mess(8) = mess(8) + "this program (except for GEOSECS choice). " + Chr$(LF)
        mess(8) = mess(8) + Chr$(LF)
        mess(8) = mess(8) + "The value of fH given in their paper was NOT the same as that given in the "
        mess(8) = mess(8) + "GEOSECS report as claimed, rather it had been rounded off and was therefore "
        mess(8) = mess(8) + "about 1% higher, corresponding to a change of .003 in pH. Note that the "
        mess(8) = mess(8) + "check value given in the paper does not match either fit. " + Chr$(LF)
        mess(8) = mess(8) + Chr$(LF)
        mess(8) = mess(8) + "Their definition of alkalinity (TA) differs from that of Dickson (Deep-Sea "
        mess(8) = mess(8) + "Research 28A:609-623, 1981 - used in constant choices 1-5 in this program) "
        mess(8) = mess(8) + "in that it is greater by an amount equal to the total phosphate (TP). This "
        mess(8) = mess(8) + "seems insignificant, but can affect the calculated fCO2 appreciably. "
'Return
'****************************************************************************
AboutPressureEffects:
        mess(9) = ""
        mess(9) = mess(9) + "The equilibrium constants depend on pressure as well as temperature and "
        mess(9) = mess(9) + "salinity. Data are scarce on these effects in seawater and most values are "
        mess(9) = mess(9) + "estimated from molal volume data. Few measurements have been made for K1, "
        mess(9) = mess(9) + "K2, and KB, at only a few combinations of temperature, salinity, and "
        mess(9) = mess(9) + "pressure (mostly in artificial seawater). All of the work assumed that fH, "
        mess(9) = mess(9) + "the activity coefficient of H+ (including liquid junction effects), is "
        mess(9) = mess(9) + "independent of pressure. Some of the pH scale conversions do depend on "
'         "pressure. Values of the constants should be converted to the seawater or "
'         "NBS pH scale WITHOUT pressure-corrected pH scale conversions, then "
'         "corrected for pressure, then converted back to the desired pH scale WITH "
'         "pressure-corrected pH scale conversions. Measurements have also been made "
'         "on the calcite and aragonite solubilities in seawater at pressure. "
        mess(9) = mess(9) + "pressure. Measurements have also been made on the calcite and aragonite "
        mess(9) = mess(9) + "solubilities in seawater at pressure. " + Chr$(LF)
        
        mess(9) = mess(9) + "Depth in meters and pressure in decibars are used interchangeably in this "
        mess(9) = mess(9) + "program. They differ by only 3% at 10000 dbar (less at lower pressures), "
        mess(9) = mess(9) + "well within the uncertainties of the pressure effects on the constants. " + Chr$(LF)
        mess(9) = mess(9) + "No salinity dependence of the pressure corrections is used in this program. " + Chr$(LF)
        mess(9) = mess(9) + "The values used are taken from: " + Chr$(LF)
        mess(9) = mess(9) + "     Millero, GCA 59:661-671, 1995, table 9 on p. 675, " + Chr$(LF)
        mess(9) = mess(9) + "     Millero, GCA 43:1651-1661, 1979, table 5 on p. 1657, " + Chr$(LF)
        mess(9) = mess(9) + "     Millero, Chap. 43, Chemical Oceanography, ed. Riley + Chester, 1983, " + Chr$(LF)
        mess(9) = mess(9) + "Note that some typos and inconsistencies from these papers were corrected. " + Chr$(LF)
        mess(9) = mess(9) + "     Takahashi et al, Chap. 3 in GEOSECS Pacific Expedition, v. 3, 1982. " + Chr$(LF)
'Return
'****************************************************************************
AboutCalciumSolubility:
        mess(10) = ""
        mess(10) = mess(10) + "The solubility product (Ksp) is calculated for both calcite and aragonite "
        mess(10) = mess(10) + "and the saturations states are given in terms of Omega, the solubility "
        mess(10) = mess(10) + "ratio, defined as Omega =  [CO3--]*[Ca++] / Ksp. Thus, values of Omega < 1 "
        mess(10) = mess(10) + "represent conditions of undersaturation, and values of Omega > 1 represent "
        mess(10) = mess(10) + "conditions of oversaturation. " + Chr$(LF)
        
        mess(10) = mess(10) + "The concentration of calcium, [Ca++], is assumed to be proportional to the "
        mess(10) = mess(10) + "salinity, and the carbonate, [CO3--], is calculated from TC, pH, and the "
        mess(10) = mess(10) + "values of K1 and K2 for carbonic acid. " + Chr$(LF)
        
        mess(10) = mess(10) + "The values used in this program are from: " + Chr$(LF)
        mess(10) = mess(10) + "     Mucci, American Journal of Science 283:781-799, 1983, " + Chr$(LF)
        mess(10) = mess(10) + "     Ingle, Marine Chemistry 3:301-319, 1975, " + Chr$(LF)
        mess(10) = mess(10) + "     Millero, Geochemica et Cosmochemica Acta 43:1651-1661, 1979, " + Chr$(LF)
        mess(10) = mess(10) + "     Takahashi et al, Chap. 3, GEOSECS Pacific Expedition, v. 3, 1982, " + Chr$(LF)
        mess(10) = mess(10) + "     Berner, American Journal of Science 276:713-730, 1976. " + Chr$(LF)
'Return
'****************************************************************************
AboutAlkalinity:
        mess(11) = ""
        mess(11) = mess(11) + "The definition of alkalinity (TA) used in this program for constant choices "
        mess(11) = mess(11) + "1 to 5 is the same as that of Dickson, Deep-Sea Research 28A:609-623, 1981: " + Chr$(LF)
        mess(11) = mess(11) + "     TA = [HCO3] + 2[CO3] + [B(OH)4] + [OH] + [HPO4] + 2[PO4] + [SiO(OH)3] " + Chr$(LF)
        mess(11) = mess(11) + "        + [HS] + 2[S] + [NH3] - [H] - [HSO4] - [HF] - [H3PO4], " + Chr$(LF)
        mess(11) = mess(11) + "except that the contributions of HS, S, and NH3 are not included. " + Chr$(LF)
        
        mess(11) = mess(11) + "For the choice of Peng, the definition of Peng et al, Tellus 39B:439-458, "
        mess(11) = mess(11) + "1987 is used. The main difference is that it is greater by an amount equal "
        mess(11) = mess(11) + "to the total phosphate: " + Chr$(LF)
        mess(11) = mess(11) + "     TP = [PO4---] + [HPO4--] + [H2PO4-] + [H3PO4]. " + Chr$(LF)
        mess(11) = mess(11) + "Though this seems small, it can have a large effect on the calculated fCO2. " + Chr$(LF)
        mess(11) = mess(11) + "Each umol/kg-SW of TA results in a change in about .5% in fCO2, so a value "
        mess(11) = mess(11) + "of TP = 3 umol/kg-SW (a modest amount) can result in a difference of "
        mess(11) = mess(11) + "5 to 20 uatm (or more) in fCO2 between the two definitions. " + Chr$(LF)
        
        mess(11) = mess(11) + "The definition used for the GEOSECS choice is: " + Chr$(LF)
        mess(11) = mess(11) + "     TA = [HCO3] + 2[CO3] + [H2BO3], " + Chr$(LF)
        mess(11) = mess(11) + "and for the freshwater choice is: " + Chr$(LF)
        mess(11) = mess(11) + "     TA = [HCO3] + 2[CO3] + [OH] - [H]. " + Chr$(LF)
        
        mess(11) = mess(11) + "In this program values of alkalinity are given in micro-moles per kilogram "
        mess(11) = mess(11) + "of seawater (umol/kg-SW). " + Chr$(LF)
'Return
'****************************************************************************
AboutRevelleFactor:
        mess(12) = ""
        mess(12) = mess(12) + "The Revelle, or homogeneous buffer, factor is the % change in fCO2 "
        mess(12) = mess(12) + "(or pCO2) caused by a 1% change in TC at constant alkalinity. " + Chr$(LF)
        
        mess(12) = mess(12) + "It depends on temperature, salinity, and the total alkalinity and TC "
        mess(12) = mess(12) + "(or any combination of the two CO2 system parameters) of the sample. " + Chr$(LF)
        mess(12) = mess(12) + Chr$(LF)
        mess(12) = mess(12) + "It is calculated at both the input and output conditions using: " + Chr$(LF)
        mess(12) = mess(12) + Chr$(LF)
        mess(12) = mess(12) + "     Revelle factor = (dfCO2/dTC) / (fCO2/TC) at constant TA. " + Chr$(LF)
        mess(12) = mess(12) + Chr$(LF)
        mess(12) = mess(12) + "Normal seawater values are between 8 and 20. " + Chr$(LF)
'Return
'****************************************************************************
AboutConstantsp1:
        mess(13) = ""
        mess(13) = mess(13) + "Constants are converted to the appropriate pH scale and concentration "
        mess(13) = mess(13) + "     scale, if needed, before calculations are made. " + Chr$(LF)
        mess(13) = mess(13) + Chr$(LF)
        mess(13) = mess(13) + "The value of K0 (the solubility coefficient of CO2) and the conversion "
        mess(13) = mess(13) + "between the fugacity and the partial pressure of CO2 are from " + Chr$(LF)
        mess(13) = mess(13) + "     Weiss, R. F., Marine Chemistry 2:203-215, 1974. " + Chr$(LF)
        mess(13) = mess(13) + Chr$(LF)
        mess(13) = mess(13) + "The vapor pressure of H2O above seawater is from " + Chr$(LF)
        mess(13) = mess(13) + "     Weiss, R. F., and Price, B. A., Marine Chemistry 8:347-359, 1980. " + Chr$(LF)
        mess(13) = mess(13) + Chr$(LF)
        mess(13) = mess(13) + "The concentrations of sulfate and fluorine are from (respectively) " + Chr$(LF)
        mess(13) = mess(13) + "     Morris and Riley, Deep-Sea Research 13:699-705, 1966, and " + Chr$(LF)
        mess(13) = mess(13) + "     Riley, J. P., Deep-Sea Research 12:219-220, 1965. " + Chr$(LF)
        mess(13) = mess(13) + Chr$(LF)
        mess(13) = mess(13) + "The value of KSO4, the dissociation constant for HSO4, is from either " + Chr$(LF)
        mess(13) = mess(13) + "     Khoo, et al, Analytical Chemistry, 49(1):29-34, 1977, or " + Chr$(LF)
        mess(13) = mess(13) + "     Dickson, Journal of Chemical Thermodynamics, 22:113-127, 1990. " + Chr$(LF)
        mess(13) = mess(13) + Chr$(LF)
        mess(13) = mess(13) + "KF, the dissociation constant for HF, is from " + Chr$(LF)
        mess(13) = mess(13) + "     Dickson, A. G. and Riley, J. P., Marine Chemistry 7:89-99, 1979. " + Chr$(LF)
        mess(13) = mess(13) + Chr$(LF)
        mess(13) = mess(13) + "Constants for calcium solubility and for pressure effects are given in "
        mess(13) = mess(13) + "     other information sections. " + Chr$(LF) + Chr$(LF)
'
'***************************************
AboutConstantsp2:
        mess(13) = mess(13) + "The value of KB (for boric acid), in constant choices 1 to 5, is from " + Chr$(LF)
        mess(13) = mess(13) + "     Dickson, Andrew G., Deep-Sea Research 37:755-766, 1990. " + Chr$(LF)
        mess(13) = mess(13) + "GEOSECS and Peng choices use Lyman's KB, the fit being from " + Chr$(LF)
        mess(13) = mess(13) + "     Li et al, Journal of Geophysical Research 74:5507-5525, 1969. " + Chr$(LF)
        mess(13) = mess(13)
        mess(13) = mess(13) + "The boron concentration for the GEOSECS and Peng choices,  is from " + Chr$(LF)
        mess(13) = mess(13) + "     Culkin, F., in Chemical Oceanography, ed. Riley and Skirrow, 1965. " + Chr$(LF)
        mess(13) = mess(13) + "For the other constant choices, it is an option selected from either: " + Chr$(LF)
        mess(13) = mess(13) + "     Uppstrom, Leif, Deep-Sea Research 21:161-162, 1974. " + Chr$(LF)
        mess(13) = mess(13) + "or  Lee et al., Geochimica Et Cosmochimica Acta 74 (6), 2010." + Chr$(LF)
        mess(13) = mess(13) + Chr$(LF)
        mess(13) = mess(13) + "Values of KW (for H2O), KP1, KP2, and KP3 (for phosphoric acid), and "
        mess(13) = mess(13) + "     KSi (for silicic acid) are from (in constant choices 1 to 5) " + Chr$(LF)
        mess(13) = mess(13) + "     Millero, Frank J., Geochemica et Cosmochemica Acta 59:661-677, 1995 "
        mess(13) = mess(13) + "     (some typos and inconsistencies from this paper were corrected). " + Chr$(LF)
        mess(13) = mess(13) + "The Peng choice uses KP2 and KP3 from " + Chr$(LF)
        mess(13) = mess(13) + "     Kester and Pytkowicz, Limnology and Oceanography 12:243-252, 1967, " + Chr$(LF)
        mess(13) = mess(13) + "and KSi from " + Chr$(LF)
        mess(13) = mess(13) + "     Sillen, Martell, and Bjerrum, Stability constants of metal-ion "
        mess(13) = mess(13) + "     complexes, The Chemical Society (London), Special Publ. 17:751, 1964. " + Chr$(LF)
        mess(13) = mess(13) + "For the Peng and the freshwater choice, KW is from " + Chr$(LF)
        mess(13) = mess(13) + "     Millero, F. J., Geochemica et Cosmochemica Acta 43:1651-1661, 1979. " + Chr$(LF)
        mess(13) = mess(13) + "     For the freshwater choice, the fit is a refit of data from " + Chr$(LF)
        mess(13) = mess(13) + "     Harned and Owen, Physical Chemistry of Electrolyte Solutions, 1958. " + Chr$(LF)
'
'***************************************
AboutConstantsp3:
        mess(13) = mess(13) + "Several determinations of K1 and K2 of carbonic acid have been made: " + Chr$(LF)
        mess(13) = mess(13) + "     Hansson (1973) on the total pH scale, " + Chr$(LF)
        mess(13) = mess(13) + "     Mehrbach et al (1973) on the NBS pH scale, " + Chr$(LF)
        mess(13) = mess(13) + "     Goyet and Poisson (1989) on the seawater scale, and " + Chr$(LF)
        mess(13) = mess(13) + "     Roy et al (1993) on the total scale. " + Chr$(LF)
        mess(13) = mess(13) + "The data of Hansson and Mehrbach et al, both seperately and together, " + Chr$(LF)
        mess(13) = mess(13) + "     have been refit by Dickson and Millero (1987) on the seawater scale. " + Chr$(LF)
        mess(13) = mess(13) + "GEOSECS and Peng et al used the fit given in Mehrbach et al. " + Chr$(LF)
        mess(13) = mess(13) + "For freshwater, Millero (1979) refit data from Harned and Davis " + Chr$(LF)
        mess(13) = mess(13) + "     for K1 and Harned and Scholes for K2. " + Chr$(LF)
        mess(13) = mess(13) + Chr$(LF)
        mess(13) = mess(13) + "The following are approximate 2s PRECISIONs of the fits of the data: "
        mess(13) = mess(13) + "     (REMEMBER THAT PRECISION AND ACCURACY ARE NOT THE SAME!): " + Chr$(LF)
        mess(13) = mess(13) + "                                              K1         K2 " + Chr$(LF)
        mess(13) = mess(13) + "                                             ----        ---- " + Chr$(LF)
        mess(13) = mess(13) + "     Roy                                   2%        1.5% " + Chr$(LF)
        mess(13) = mess(13) + "     Goyet and Poisson             2.5%     4.5% " + Chr$(LF)
        mess(13) = mess(13) + "     Hansson, refit by DM         3%       4% " + Chr$(LF)
        mess(13) = mess(13) + "     Mehrbach, refit by DM     2.5%     4.5% " + Chr$(LF)
        mess(13) = mess(13) + "     DM combined fit                4%        6% " + Chr$(LF)
        mess(13) = mess(13) + "     Mehrbach's fit                  1.2%      2% " + Chr$(LF)
        mess(13) = mess(13) + "     freshwater choice             0.5%    0.7% " + Chr$(LF) + Chr$(LF)
'
'***************************************
AboutConstantsp4:
        mess(13) = mess(13) + "References are: " + Chr$(LF)
        mess(13) = mess(13) + "     Roy, et al, Marine Chemistry 44:249-267,1993 " + Chr$(LF)
        mess(13) = mess(13) + "        see also: Erratum, Marine Chemistry 45:337, 1994 " + Chr$(LF)
        mess(13) = mess(13) + "        and Erratum, Marine Chemistry 52:183, 1996 " + Chr$(LF)
        mess(13) = mess(13) + "     Goyet and Poisson, Deep-Sea Research 36:1635-1654, 1989 " + Chr$(LF)
        mess(13) = mess(13) + "     Hansson, Deep-Sea Research 20:461-478, 1973 " + Chr$(LF)
        mess(13) = mess(13) + "     Hansson, Acta Chemica Scandanavia, 27:931-944, 1973, " + Chr$(LF)
        mess(13) = mess(13) + "     Mehrbach et al, Limnology and Oceaneanography, 18:897-907, 1973 " + Chr$(LF)
        mess(13) = mess(13) + "     Dickson and Millero, Deep-Sea Research, 34:1733-1743,1987 " + Chr$(LF)
        mess(13) = mess(13) + "        see also Corrigenda, Deep-Sea Research, 36:983, 1989 " + Chr$(LF)
        mess(13) = mess(13) + "     Millero, F. J., Geochemica et Cosmochemica Acta 43:1651-1661, 1979 " + Chr$(LF)
        mess(13) = mess(13) + "     Harned and Davis, J American Chemical Society, 65:2030-2037, 1943 " + Chr$(LF)
        mess(13) = mess(13) + "     Harned and Scholes, J American Chemical Society, 43:1706-1709, 1941 " + Chr$(LF) + Chr$(LF)
        mess(13) = mess(13) + "     Cai and Wang, Limnol. Oceanogr. 43:657-668, 1998" + Chr$(LF) + Chr$(LF)
        mess(13) = mess(13) + "     Lueker et al., Mar. Chem. 70:105-119, 2000 " + Chr$(LF) + Chr$(LF)
        mess(13) = mess(13) + "     Mojica Prieto and Millero, Geochim. et Cosmochim. Acta. 66:2529-2540, 2002 " + Chr$(LF) + Chr$(LF)
        mess(13) = mess(13) + "     Millero et al., Deep-Sea Res. I (49) 1705-1723, 2002 " + Chr$(LF) + Chr$(LF)
        mess(13) = mess(13) + "     Millero et al., Mar.Chem. 100:80-94, 2006" + Chr$(LF) + Chr$(LF)
        mess(13) = mess(13) + "     Millero, Marine and Freshwater Research, v. 61, p. 139-142, 2010" + Chr$(LF) + Chr$(LF)
        
        mess(13) = mess(13) + "A very useful reference for all aspects of the CO2 system in seawater is " + Chr$(LF)
        mess(13) = mess(13) + "     Guide to best practices for ocean CO2 measurements. " + Chr$(LF)
        mess(13) = mess(13) + "     PICES Special Publication 3, 191 pp. " + Chr$(LF)
        mess(13) = mess(13) + "     Dickson, A.G., Sabine, C.L. and Christian, J.R. (Eds.) 2007. "
'****************************************************************************
AboutMacroHistory1:
        mess(14) = ""
        mess(14) = mess(14) + "Previous versions (2007): CO2sys_macro_PC.xls   and CO2sys_macro_MAC.xls" + Chr$(LF) + Chr$(LF)
        mess(14) = mess(14) + "      . Two separate files for PC and MAC versions." + Chr$(LF)
        mess(14) = mess(14) + Chr$(LF)
        mess(14) = mess(14) + "Version 1.0 (10 Octobre 2011): CO2sys_2011.xls" + Chr$(LF) + Chr$(LF)
        mess(14) = mess(14) + "      . Combined PC and MAC versions of previous macro into one file working on both platforms." + Chr$(LF)
        mess(14) = mess(14) + Chr$(LF)
        mess(14) = mess(14) + "Version 2.0 (19 July 2012): CO2sys_2011.xls" + Chr$(LF) + Chr$(LF)
        mess(14) = mess(14) + "      . New R formulation from ""NIST Physical Reference Data (http://physics.nist.gov/cgi-bin/cuu/Value?r)""" + Chr$(LF)
        mess(14) = mess(14) + "               Difference with old formulation  is not numerically significant." + Chr$(LF)
        mess(14) = mess(14) + "      . Matched formulation of Uppstrom's Total Boron with Matlab program (same numerical results)." + Chr$(LF)
        mess(14) = mess(14) + "      . Added option of Total Boron from Lee et al., 2010" + Chr$(LF)
        mess(14) = mess(14) + "      . Added a few formulations for K1, K2:" + Chr$(LF)
        mess(14) = mess(14) + "                - Cai and Wang, 1998" + Chr$(LF)
        mess(14) = mess(14) + "                - Lueker et al., 2000" + Chr$(LF)
        mess(14) = mess(14) + "                - Mojica Prieto et al., 2002" + Chr$(LF)
        mess(14) = mess(14) + "                - Millero et al., 2002" + Chr$(LF)
        mess(14) = mess(14) + "                - Millero et al., 2006" + Chr$(LF)
        mess(14) = mess(14) + "                - Millero, 2010" + Chr$(LF)
        mess(14) = mess(14) + "      . Updated the ""INFO"" section" + Chr$(LF)
        mess(14) = mess(14) + "      . Added the ""Macro Version History"" option in ""INFO"" Sheet." + Chr$(LF)
        mess(14) = mess(14) + "      . Version number is displayed in cell B2 when the ""About this Macro"" option in ""INFO"" Sheet is selected." + Chr$(LF)
        mess(14) = mess(14) + Chr$(LF)
        mess(14) = mess(14) + "Version 2.1 (18 September 2012): CO2sys_v2.1.xls" + Chr$(LF) + Chr$(LF)
        mess(14) = mess(14) + "      . Corrected an error in the code which affected the results when the constants of 'Millero et. al., 2002'" + Chr$(LF)
        mess(14) = mess(14) + "                and 'Millero, 2010' were selected." + Chr$(LF)
        mess(14) = mess(14) + "      . References to  'Cai and Wang, 2008' have been corrected to 'Cai and Wang, 1998'" + Chr$(LF)
        mess(14) = mess(14) + "      . Incorporated version number in the name of the file and removed it from the 'INFO' sheet (see v.2.0)" + Chr$(LF)

        mess(14) = mess(14) + Chr$(LF)
'Return
'****************************************************************************

End Sub

Sub Initiate()
   

'********************** DEFAULT**************************
WhichKsDefault% = 4             ' D&M refit
WhoseKSO4Default% = 1 ' Dickson's KSO4
pHScaleDefault% = 2   ' Total pH scale
WhichTBDefault% = 1             ' Uppstrom
'BatchDefault% = 1        ' single-input mode
'TA = 0.0023:            ' mol/kg-SW
'TC = 0.0021:            ' mol/kg-SW
'pHinp = 7.9:            '
'fCO2inp = 0.0006:       ' atm
'pCO2inp = 0.0006:       ' atm
'Sal = 35!:              ' mille
'TempCinp = 20!:         ' deg C
'Pdbarinp = 0!:          ' decibars
'TempCout = 5!:          ' deg C
'Pdbarout = 0!:          ' decibars
' for batch-input mode:
'NHeaderLines% = 1:      ' number of header lines in input file
'NIDFields% = 1:         ' number of ID fields per sample
'MVD = -9:               ' missing value designator
'MVFlag$ = "Y":          ' missing value flag
'********************** END OF DEFAULT**************************
'********************** CO2**************************
'co2opt(1) = "fCO2": co2opt(2) = "pCO2"
'********************** END OF CO2 **************************
'********************** pH SCALES **************************
phopt(1) = "Total scale (mol/kg-SW) "
phopt(2) = "Seawater scale (mol/kg-SW) "
phopt(3) = "Free scale (mol/kg-SW) "
phopt(4) = "NBS scale (mol/kg-H2O) "
'********************** END OF pH SCALES **************************
'********************** KHSO4 **************************
khso4opt(1) = "Dickson "
khso4opt(2) = "Khoo et al "
'********************** SET OF CONSTANTS **************************
        kopt(1) = " K1, K2 from Roy, et al., 1993 "   '2s PRECISION about 2% in K1 and 1.5% in K2. "
        kopt(2) = " K1, K2 from Goyet and Poisson, 1989 "  '2s PRECISION about 2.5% in K1 and 4.5% in K2. "
        kopt(3) = " K1, K2 from Hansson, 1973 refit by Dickson and Millero, 1987 "  '2s PRECISION about 3% in K1 and 4% in K2. "
        kopt(4) = " K1, K2 from Mehrbach et al., 1973 refit by Dickson and Millero, 1987 "  '2s PRECISION about 2.5% in K1 and 4.5% in K2. "
        kopt(5) = " K1, K2 from Hansson and Mehrbach refit by Dickson and Millero, 1987 " '2s PRECISION about 4% in K1 and 6% in K2. "
        kopt(6) = " GEOSECS constants (NBS scale); K1, K2 from Mehrbach et al., 1973 " '2s PRECISION about 1.2% in K1 and 2.0% in K2. "
        kopt(7) = " Constants from Peng et al. (NBS scale); K1, K2 from Mehrbach et al. "  '2s PRECISION about 1.2% in K1 and 2.0% in K2. "
        kopt(8) = " Salinity = 0 (freshwater); K1, K2 from Millero, 1979 "   '2s PRECISION about 0.5% in K1 and 0.7% in K2. "
        kopt(9) = " K1, K2 from Cai and Wang, 1998"
        kopt(10) = " K1, K2 from Lueker et al., 2000"
        kopt(11) = " K1, K2 from Mojica Prieto et al., 2002"
        kopt(12) = " K1, K2 from Millero et al., 2002"
        kopt(13) = " K1, K2 from Millero et al., 2006"
        kopt(14) = " K1, K2 from Millero, 2010"
       ' For 6) or 7), the pH scale is set to NBS. This can be changed later.
       'In each case the constants are converted to the chosen pH scale.
'********************** END OF SET CONSTANTS **************************
'********************** Total Boron **************************
tbopt(1) = "Uppstrom, 1974 "
tbopt(2) = "Lee et al., 2010 "
'********************** END TB  **************************
Sheets("INPUT").Select
'Range("A1").Select
For i = 1 To UBound(kopt)
   If i <= UBound(khso4opt) Then Sheets("INPUT").Range("A1").Offset(i, 1).Value = khso4opt(i)
   If i <= UBound(phopt) Then Sheets("INPUT").Range("A1").Offset(i, 2).Value = phopt(i)
   Sheets("INPUT").Range("A1").Offset(i, 0).Value = kopt(i)
Next i
With Sheets("INPUT").Range(Cells(2, 1).Address, Cells(1 + UBound(kopt), 1).Address)
   .Interior.ColorIndex = xlNone
   .HorizontalAlignment = xlLeft
   .VerticalAlignment = xlCenter
   .WrapText = True
End With
With Sheets("INPUT").Range(Cells(2, 2).Address, Cells(1 + UBound(khso4opt), 2).Address)
   .Interior.ColorIndex = xlNone
   .HorizontalAlignment = xlLeft
   .VerticalAlignment = xlCenter
   .WrapText = True
End With
With Sheets("INPUT").Range(Cells(2, 3).Address, Cells(1 + UBound(phopt), 3).Address)
   .Interior.ColorIndex = xlNone
   .HorizontalAlignment = xlLeft
   .VerticalAlignment = xlCenter
   .WrapText = True
End With
With Sheets("INPUT").Range(Cells(2, 4).Address, Cells(1 + UBound(tbopt), 3).Address)
   .Interior.ColorIndex = xlNone
   .HorizontalAlignment = xlLeft
   .VerticalAlignment = xlCenter
   .WrapText = True
End With
With Sheets("INPUT").Cells(1 + WhichKsDefault%, 1).Interior
        .ColorIndex = 36
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
End With
With Sheets("INPUT").Cells(1 + WhoseKSO4Default%, 2).Interior
        .ColorIndex = 36
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
End With
With Sheets("INPUT").Cells(1 + pHScaleDefault%, 3).Interior
        .ColorIndex = 36
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
End With
With Sheets("INPUT").Cells(1 + WhichTBDefault%, 4).Interior
        .ColorIndex = 36
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
End With
 
'********************** END OF KHSO4**************************
'Sheets("INPUT").phcombo.List = phopt
'Sheets("INPUT").kcombo.List = kopt
'Sheets("INPUT").khso4combo.List = khso4opt
  
'Sheets("INPUT").phcombo.ListIndex = (pHScaleDefault% - 1)
'Sheets("INPUT").kcombo.ListIndex = (WhichKsDefault% - 1)
'Sheets("INPUT").khso4combo.ListIndex = (WhoseKSO4Default% - 1)

'**************** INITIATES INFO SHEET ***********************
 Dim ms(14) As String
 ms(1) = "About this Macro"
 ms(2) = "General Information ": ms(3) = "pH Scales": ms(4) = "fCO2, pCO2"
 ms(5) = "KSO4": ms(6) = "Freshwater Option"
 ms(7) = "GEOSECS Option ": ms(8) = "Peng Option ":  ms(9) = "Pressure Effects"
 ms(10) = "Calcium Carbonate Solubility (Omega Values) "
 ms(11) = "Alkalinity": ms(12) = "Revelle Factor": ms(13) = "Constants": ms(14) = "Macro Version History"

Sheets("INFO").Select
'Sheets("INFO").Range("A1").Select
Sheets("INFO").Range("A1").Value = "SUBJECT"
For i = 1 To UBound(ms)
  If Sheets("INFO").Range("A1").Offset(i, 0).Value <> ms(i) Then
     Sheets("INFO").Range("A1").Offset(i, 0).Value = ms(i)
  End If
Next i
'Sheets("INFO").Range("A2").Select

Erase ms
InitiateOK = True

End Sub

Sub Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
' SUB Constants, version 04.01, 10-13-97, written by Ernie Lewis.
' Inputs: pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar
' Outputs: K0, K(), T(), fH, FugFac, VPFac
' This finds the constants of the CO2 system in seawater or freshwater,
' corrects them for pressure, and reports them on the chosen pH scale.
' The process is as follows: the constants (except KS, KF which stay on the
' free scale - these are only corrected for pressure) are
'       1) evaluated as they are given in the literature
'       2) converted to the SWS scale in mol/kg-SW or to the NBS scale
'       3) corrected for pressure
'       4) converted to the SWS pH scale in mol/kg-SW
'       5) converted to the chosen pH scale
'
'
'*******************************
'       PROGRAMMER'S NOTE: all logs are log base e
'       PROGRAMMER'S NOTE: all constants are converted to the pH scale
'               pHScale% (the chosen one) in units of mol/kg-SW
'               except KS and KF are on the free scale
'               and KW is in units of (mol/kg-SW)^2
'
'
'*******************************
        'RGasConstant = 83.1451: 'bar-cm3/(mol-K): ' = 8.31451 N-m/(mol-K)
       RGasConstant = 83.144621: 'bar-cm3/(mol-K)   from NIST Physical Reference Data (http://physics.nist.gov/cgi-bin/cuu/Value?r)
       TempK = TempC + 273.15
        RT = RGasConstant * TempK
        sqrSal = Sqr(Sal)
        logTempK = Log(TempK)
        Pbar = Pdbar / 10!
'       deltaVs are in cm3/mole
'       Kappas are in cm3/mole/bar
'
'
'****************************************************************************
CalculateTB:
        Select Case WhichKs%
        Case 1, 2, 3, 4, 5, 9, 10, 11, 12, 13, 14
                Select Case WhichTB%
                Case 1        ' Uppstrom, L., Deep-Sea Research 21:161-162, 1974:
                        'TB = (0.000232 / 10.811) * (Sal / 1.80655): ' in mol/kg-SW
                        ' this is .000416 * Sal / 35. = .0000119 * Sal
                        TB = 0.0004157 * Sal / 35# ' in mol/kg-SW
                 Case 2 ' Lee, Kim, Byrne, Millero, Feely, Yong-Ming Liu. 2010.  Geochimica Et Cosmochimica Acta 74 (6)
                        TB = 0.0004326 * Sal / 35# ' in mol/kg-SW
                 End Select
       Case 6, 7
                ' Culkin, F., in Chemical Oceanography,
                ' ed. Riley and Skirrow, 1965:
                ' GEOSECS references this, but this value is not explicitly
                ' given here
                TB = 0.0004106 * Sal / 35!: ' in mol/kg-SW
                ' this is .00001173 * Sal
                ' this is about 1% lower than Uppstrom's value
        Case 8
                TB = 0!
        End Select
'
'
'*******************************
CalculateTF:
        ' Riley, J. P., Deep-Sea Research 12:219-220, 1965:
        TF = (0.000067 / 18.998) * (Sal / 1.80655): ' in mol/kg-SW
        ' this is .000068 * Sal / 35. = .00000195 * Sal
'
'
'*******************************
CalculateTS:
        ' Morris, A. W., and Riley, J. P., Deep-Sea Research 13:699-705, 1966:
        TS = (0.14 / 96.062) * (Sal / 1.80655): ' in mol/kg-SW
        ' this is .02824 * Sal / 35. = .0008067 * Sal
'
'
'*******************************
MakeTMatrix:
        T(1) = TB
        T(2) = TF
        T(3) = TS
'       T(4) = TP
'       T(5) = TSi
'       These last two were set earlier.
'
'
'****************************************************************************
CalculateK0:
        ' Weiss, R. F., Marine Chemistry 2:203-215, 1974.
        TTT = TempK / 100!
        lnK0 = -60.2409 + 93.4517 / TTT + 23.3585 * Log(TTT)
        lnK0 = lnK0 + Sal * (0.023517 - 0.023656 * TTT + 0.0047036 * TTT * TTT)
        K0 = Exp(lnK0): ' this is in mol/kg-SW/atm
'
'
'****************************************************************************
CalculateIonS:
'       This is from the DOE handbook, Chapter 5, p. 13/22, eq. 7.2.4:
        IonS = 19.924 * Sal / (1000! - 1.005 * Sal)
'
'
'****************************************************************************
CalculateKS:
        Select Case WhoseKSO4%
        Case 1: ' Dickson's value
'               Dickson, A. G., J. Chemical Thermodynamics, 22:113-127, 1990
'               The goodness of fit is .021.
'               It was given in mol/kg-H2O. I convert it to mol/kg-SW.
'               TYPO!!!!!! on p. 121: the constant e9 should be e8.
'
'       This is from eqs 22 and 23 on p. 123, and Table 4 on p 121:
        lnKS = -4276.1 / TempK + 141.328 - 23.093 * logTempK
        lnKS = lnKS + (-13856! / TempK + 324.57 - 47.986 * logTempK) * Sqr(IonS)
        lnKS = lnKS + (35474! / TempK - 771.54 + 114.723 * logTempK) * IonS
        lnKS = lnKS + (-2698! / TempK) * Sqr(IonS) * IonS
        lnKS = lnKS + (1776! / TempK) * IonS * IonS
        KS = Exp(lnKS): ' this is on the free pH scale in mol/kg-H2O
        KS = KS * (1! - 0.001005 * Sal): ' convert to mol/kg-SW
'
'
        Case 2
'               Khoo, et al, Analytical Chemistry, 49(1):29-34, 1977
'               KS was found by titrations with a hydrogen electrode
'               of artificial seawater containing sulfate (but without F)
'               at 3 salinities from 20 to 45 and artificial seawater NOT
'               containing sulfate (nor F) at 16 salinities from 15 to 45,
'               both at temperatures from 5 to 40 deg C.
'               KS is on the Free pH scale (inherently so).
'               It was given in mol/kg-H2O. I convert it to mol/kg-SW.
'               He finds log(beta) which = my pKS;
'               his beta is an association constant.
'               The rms error is .0021 in pKS, or about .5% in KS.
'
'               This is equation 20 on p. 33:
        pKS = 647.59 / TempK - 6.3451 + 0.019085 * TempK - 0.5208 * Sqr(IonS)
        KS = 10! ^ (-pKS): ' this is on the free pH scale in mol/kg-H2O
        KS = KS * (1! - 0.001005 * Sal): ' convert to mol/kg-SW
        End Select
'
'
'****************************************************************************
CalculateKF:
        ' Dickson, A. G. and Riley, J. P., Marine Chemistry 7:89-99, 1979:
        lnKF = 1590.2 / TempK - 12.641 + 1.525 * Sqr(IonS)
        KF = Exp(lnKF): ' this is on the free pH scale in mol/kg-H2O
        KF = KF * (1! - 0.001005 * Sal): ' convert to mol/kg-SW
'
'
'****************************************************************************
CalculatepHScaleConversionFactors:
'       These are NOT pressure-corrected.
        SWStoTOT = (1! + TS / KS) / (1! + TS / KS + TF / KF)
        FREEtoTOT = 1! + TS / KS
'
'
'
CalculatefH:
'       Use GEOSECS's value for cases 1,2,3,4,5 (and 6) to convert pH scales.
        Select Case WhichKs%
        Case 1, 2, 3, 4, 5, 6, 9, 10, 11, 12, 13, 14
                ' Takahashi et al, Chapter 3 in GEOSECS Pacific Expedition,
                ' v. 3, 1982 (p. 80):
                fH = 1.2948 - 0.002036 * TempK
                fH = fH + (0.0004607 - 0.000001475 * TempK) * Sal * Sal
        Case 7
                ' Peng et al, Tellus 39B:439-458, 1987:
                ' They reference the GEOSECS report, but round the value
                ' given there off so that it is about .008 (1%) lower. It
                ' doesn't agree with the check value they give on p. 456.
                fH = 1.29 - 0.00204 * TempK
                fH = fH + (0.00046 - 0.00000148 * TempK) * Sal * Sal
        Case 8
                fH = 1!: ' this shouldn't occur in the program for this case
        End Select
'
'
'****************************************************************************
CalculateKB:
        Select Case WhichKs%
        Case 1, 2, 3, 4, 5, 9, 10, 11, 12, 13, 14
                ' Dickson, A. G., Deep-Sea Research 37:755-766, 1990:
                lnKBtop = -8966.9 - 2890.53 * sqrSal - 77.942 * Sal
                lnKBtop = lnKBtop + 1.728 * sqrSal * Sal - 0.0996 * Sal * Sal
                lnKB = lnKBtop / TempK
                lnKB = lnKB + 148.0248 + 137.1942 * sqrSal + 1.62142 * Sal
                lnKB = lnKB + (-24.4344 - 25.085 * sqrSal - 0.2474 * Sal) * logTempK
                lnKB = lnKB + 0.053105 * sqrSal * TempK
                KB = Exp(lnKB): ' this is on the total pH scale in mol/kg-SW
                KB = KB / SWStoTOT: ' convert to SWS pH scale
'
'
        Case 6, 7
                ' This is for GEOSECS and Peng et al.
                ' Lyman, John, UCLA Thesis, 1957
                ' fit by Li et al, JGR 74:5507-5525, 1969:
                logKB = -9.26 + 0.00886 * Sal + 0.01 * TempC
                KB = 10! ^ (logKB): ' this is on the NBS scale
                KB = KB / fH: ' convert to the SWS scale
'
'
        Case 8
                KB = 0!
        End Select
'
'
'***************************************************************************
CalculateKW:
        Select Case WhichKs%
        Case 1, 2, 3, 4, 5, 9, 10, 11, 12, 13, 14
                ' Millero, Geochemica et Cosmochemica Acta 59:661-677, 1995.
                ' his check value of 1.6 umol/kg-SW should be 6.2
                lnKW = 148.9802 - 13847.26 / TempK - 23.6521 * logTempK
                lnKW = lnKW + (-5.977 + 118.67 / TempK + 1.0495 * logTempK) * sqrSal
                lnKW = lnKW - 0.01615 * Sal
                KW = Exp(lnKW): ' this is on the SWS pH scale in (mol/kg-SW)^2
'
'
        Case 6
                KW = 0!: ' GEOSECS doesn't include OH effects
'
'
        Case 7
                ' Millero, Geochemica et Cosmochemica Acta 43:1651-1661, 1979
                lnKW = 148.9802 - 13847.26 / TempK - 23.6521 * logTempK
                lnKW = lnKW + (-79.2447 + 3298.72 / TempK + 12.0408 * logTempK) * sqrSal
                lnKW = lnKW - 0.019813 * Sal
                KW = Exp(lnKW): ' this is on the SWS pH scale
'
'
        Case 8
                ' Millero, Geochemica et Cosmochemica Acta 43:1651-1661, 1979
                ' refit data of Harned and Owen, The Physical Chemistry of
                ' Electrolyte Solutions, 1958
                lnKW = 148.9802 - 13847.26 / TempK - 23.6521 * logTempK
                KW = Exp(lnKW)
        End Select
'
'
'***************************************************************************
CalculateKP1KP2KP3KSi:
        Select Case WhichKs%
        Case 1, 2, 3, 4, 5, 9, 10, 11, 12, 13, 14
                ' Yao and Millero, Aquatic Geochemistry 1:53-88, 1995
                ' KP1, KP2, KP3 are on the SWS pH scale in mol/kg-SW.
                ' KSi was given on the SWS pH scale in molal units.
'
                lnKP1 = -4576.752 / TempK + 115.54 - 18.453 * logTempK
                lnKP1 = lnKP1 + (-106.736 / TempK + 0.69171) * sqrSal
                lnKP1 = lnKP1 + (-0.65643 / TempK - 0.01844) * Sal
                KP1 = Exp(lnKP1)
'
'
                lnKP2 = -8814.715 / TempK + 172.1033 - 27.927 * logTempK
                lnKP2 = lnKP2 + (-160.34 / TempK + 1.3566) * sqrSal
                lnKP2 = lnKP2 + (0.37335 / TempK - 0.05778) * Sal
                KP2 = Exp(lnKP2)
'
'
                lnKP3 = -3070.75 / TempK - 18.126
                lnKP3 = lnKP3 + (17.27039 / TempK + 2.81197) * sqrSal
                lnKP3 = lnKP3 + (-44.99486 / TempK - 0.09984) * Sal
                KP3 = Exp(lnKP3)
'
'
                lnKSi = -8904.2 / TempK + 117.4 - 19.334 * logTempK
                lnKSi = lnKSi + (-458.79 / TempK + 3.5913) * Sqr(IonS)
                lnKSi = lnKSi + (188.74 / TempK - 1.5998) * IonS
                lnKSi = lnKSi + (-12.1652 / TempK + 0.07871) * IonS * IonS
                KSi = Exp(lnKSi): ' this is on the SWS pH scale in mol/kg-H2O
                KSi = KSi * (1! - 0.001005 * Sal): ' convert to mol/kg-SW
'
'
        Case 7
                KP1 = 0.02:
                ' Peng et al don't include the contribution from this term,
                ' but it is so small it doesn't contribute. It needs to be
                ' kept so that the routines work ok.
                ' KP2, KP3 from Kester, D. R., and Pytkowicz, R. M.,
                ' Limnology and Oceanography 12:243-252, 1967:
                ' these are only for sals 33 to 36 and are on the NBS scale
                KP2 = Exp(-9.039 - 1450! / TempK): ' this is on the NBS scale
                KP2 = KP2 / fH: ' convert to SWS scale
                KP3 = Exp(4.466 - 7276 / TempK): ' this is on the NBS scale
                KP3 = KP3 / fH: ' convert to SWS scale
                ' Sillen, Martell, and Bjerrum,  Stability constants of metal-ion complexes,
                ' The Chemical Society (London), Special Publ. 17:751, 1964:
                KSi = 0.0000000004: ' this is on the NBS scale
                KSi = KSi / fH: ' convert to SWS scale
        Case 6, 8
                KP1 = 0!
                KP2 = 0!
                KP3 = 0!
                KSi = 0!
                ' Neither the GEOSECS choice nor the freshwater choice
                ' include contributions from phosphate or silicate.
        End Select
'
'
'***************************************************************************
CalculateK1K2:
        Select Case WhichKs%
'************************************************
        Case 1: ' ROY et al, Marine Chemistry, 44:249-267, 1993
'               (see also: Erratum, Marine Chemistry 45:337, 1994
'               and Erratum, Marine Chemistry 52:183, 1996)
'               !!! Typo: in the abstract on p. 249: in the eq. for lnK1* the
'                       last term should have S raised to the power 1.5.
'               They claim standard deviations (p. 254) of the fits as
'               .0048 for lnK1 (.5% in K1) and .007 in lnK2 (.7% in K2).
'               They also claim (p. 258) 2s precisions of .004 in pK1 and
'               .006 in pK2. These are consistent, but Andrew Dickson
'               (personal communication) obtained an rms deviation of about
'               .004 in pK1 and .003 in pK2. This would be a 2s precision
'               of about 2% in K1 and 1.5% in K2.
'
'
'       This is eq. 29 on p. 254 and what they use in their abstract:
        lnK1 = 2.83655 - 2307.1266 / TempK - 1.5529413 * logTempK
        lnK1 = lnK1 + (-0.20760841 - 4.0484 / TempK) * sqrSal
        lnK1 = lnK1 + 0.08468345 * Sal - 0.00654208 * sqrSal * Sal
        K1 = Exp(lnK1): ' this is on the total pH scale in mol/kg-H2O
        K1 = K1 * (1! - 0.001005 * Sal): ' convert to mol/kg-SW
        K1 = K1 / SWStoTOT: ' convert to SWS pH scale
'
'       This is eq. 30 on p. 254 and what they use in their abstract:
        lnK2 = -9.226508 - 3351.6106 / TempK - 0.2005743 * logTempK
        lnK2 = lnK2 + (-0.106901773 - 23.9722 / TempK) * sqrSal
        lnK2 = lnK2 + 0.1130822 * Sal - 0.00846934 * sqrSal * Sal
        K2 = Exp(lnK2): ' this is on the total pH scale in mol/kg-H2O
        K2 = K2 * (1! - 0.001005 * Sal): ' convert to mol/kg-SW
        K2 = K2 / SWStoTOT: ' convert to SWS pH scale
'
'
'************************************************
        Case 2: ' GOYET AND POISSON, Deep-Sea Research, 36(11):1635-1654, 1989
'               The 2s precision in pK1 is .011, or 2.5% in K1.
'               The 2s precision in pK2 is .02, or 4.5% in K2.
'
'       This is in Table 5 on p. 1652 and what they use in the abstract:
        pK1 = 812.27 / TempK + 3.356 - 0.00171 * Sal * logTempK
        pK1 = pK1 + 0.000091 * Sal * Sal
        K1 = 10! ^ (-pK1): ' this is on the SWS pH scale in mol/kg-SW
'
'       This is in Table 5 on p. 1652 and what they use in the abstract:
        pK2 = 1450.87 / TempK + 4.604 - 0.00385 * Sal * logTempK
        pK2 = pK2 + 0.000182 * Sal * Sal
        K2 = 10! ^ (-pK2): ' this is on the SWS pH scale in mol/kg-SW
'
'
'************************************************
        Case 3: ' HANSSON refit BY DICKSON AND MILLERO
'               Dickson and Millero, Deep-Sea Research, 34(10):1733-1743, 1987
'               (see also Corrigenda, Deep-Sea Research, 36:983, 1989)
'               refit data of Hansson, Deep-Sea Research, 20:461-478, 1973
'               and Hansson, Acta Chemica Scandanavia, 27:931-944, 1973.
'               on the SWS pH scale in mol/kg-SW.
'               Hansson gave his results on the Total scale (he called it
'                       the seawater scale) and in mol/kg-SW.
'               !!! Typo in DM on p. 1739 in Table 4: the equation for pK2*
'                       for Hansson should have a .000132 *S^2
'                       instead of a .000116 *S^2.
'               The 2s precision in pK1 is .013, or 3% in K1.
'               The 2s precision in pK2 is .017, or 4.1% in K2.
'
'       This is from Table 4 on p. 1739.
        pK1 = 851.4 / TempK + 3.237 - 0.0106 * Sal + 0.000105 * Sal * Sal
        K1 = 10! ^ (-pK1): ' this is on the SWS pH scale in mol/kg-SW
'
'       This is from Table 4 on p. 1739.
        pK2 = -3885.4 / TempK + 125.844 - 18.141 * logTempK
        pK2 = pK2 - 0.0192 * Sal + 0.000132 * Sal * Sal
        K2 = 10! ^ (-pK2): ' this is on the SWS pH scale in mol/kg-SW
'
'
'************************************************
        Case 4: ' MEHRBACH refit BY DICKSON AND MILLERO
'               Dickson and Millero, Deep-Sea Research, 34(10):1733-1743, 1987
'               (see also Corrigenda, Deep-Sea Research, 36:983, 1989)
'               refit data of Mehrbach et al, Limn Oc, 18(6):897-907, 1973
'               on the SWS pH scale in mol/kg-SW.
'               Mehrbach et al gave results on the NBS scale.
'               The 2s precision in pK1 is .011, or 2.6% in K1.
'               The 2s precision in pK2 is .020, or 4.6% in K2.
'
'       This is in Table 4 on p. 1739.
        pK1 = 3670.7 / TempK - 62.008 + 9.7944 * logTempK
        pK1 = pK1 - 0.0118 * Sal + 0.000116 * Sal * Sal
        K1 = 10! ^ (-pK1): ' this is on the SWS pH scale in mol/kg-SW
'
'       This is in Table 4 on p. 1739.
        pK2 = 1394.7 / TempK + 4.777 - 0.0184 * Sal + 0.000118 * Sal * Sal
'        If TCO2m >= 2050 Then
 '         pK2 = pK2 - 0.00015 * (TCO2m - 2050)
 '       End If
        K2 = 10! ^ (-pK2): ' this is on the SWS pH scale in mol/kg-SW
'
'
'************************************************
        Case 5: ' HANSSON and MEHRBACH refit BY DICKSON AND MILLERO
'               Dickson and Millero, Deep-Sea Research,34(10):1733-1743, 1987
'               (see also Corrigenda, Deep-Sea Research, 36:983, 1989)
'               refit data of Hansson, Deep-Sea Research, 20:461-478, 1973,
'               Hansson, Acta Chemica Scandanavia, 27:931-944, 1973,
'               and Mehrbach et al, Limnol. Oceanogr.,18(6):897-907, 1973
'               on the SWS pH scale in mol/kg-SW.
'               !!! Typo in DM on p. 1740 in Table 5: the second equation
'                       should be pK2* =, not pK1* =.
'               The 2s precision in pK1 is .017, or 4% in K1.
'               The 2s precision in pK2 is .026, or 6% in K2.
'
'       This is in Table 5 on p. 1740.
        pK1 = 845! / TempK + 3.248 - 0.0098 * Sal + 0.000087 * Sal * Sal
        K1 = 10! ^ (-pK1): ' this is on the SWS pH scale in mol/kg-SW
'
'       This is in Table 5 on p. 1740.
        pK2 = 1377.3 / TempK + 4.824 - 0.0185 * Sal + 0.000122 * Sal * Sal
        K2 = 10! ^ (-pK2): ' this is on the SWS pH scale in mol/kg-SW
'
'
'************************************************
        Case 6, 7:
'               GEOSECS and Peng et al use K1, K2 from Mehrbach et al,
'               Limnology and Oceanography, 18(6):897-907, 1973.
'               The 2s precision in pK1 is .005, or 1.2% in K1.
'               The 2s precision in pK2 is .008, or 2% in K2.

        logK1 = 13.7201 - 0.031334 * TempK - 3235.76 / TempK
        logK1 = logK1 - 0.000013 * Sal * TempK + 0.1032 * sqrSal
        K1 = 10! ^ (logK1): ' this is on the NBS scale
        K1 = K1 / fH: ' convert to SWS scale
'
        logK2 = -5371.9645 - 1.671221 * TempK + 128375.28 / TempK
        logK2 = logK2 + 2194.3055 * logTempK / Log(10!) - 0.22913 * Sal
        logK2 = logK2 - 18.3802 * Log(Sal) / Log(10!) + 0.00080944 * Sal * TempK
        logK2 = logK2 + 5617.11 * Log(Sal) / Log(10!) / TempK - 2.136 * Sal / TempK
        K2 = 10! ^ (logK2): ' this is on the NBS scale
        K2 = K2 / fH: ' convert to SWS scale
'
'
'************************************************
        Case 8
'               Millero, F. J., Geochemica et Cosmochemica Acta 43:1651-1661, 1979:
'               K1 from refit data from Harned and Davis,
'                       J American Chemical Society, 65:2030-2037, 1943.
'               K2 from refit data from Harned and Scholes,
'                       J American Chemical Society, 43:1706-1709, 1941.
'       These are the thermodynamic constants:
        lnK1 = 290.9097 - 14554.21 / TempK - 45.0575 * logTempK
        K1 = Exp(lnK1)
        lnK2 = 207.6548 - 11843.79 / TempK - 33.6485 * logTempK
        K2 = Exp(lnK2)
'
'
'************************************************
 '************************************************
        Case 9   ' From Cai and Wang 1998, for estuarine use.
            ' Data used in this work is from:
            ' K1: Merhback (1973) for S>15, for S<15: Mook and Keone (1975)
            ' K2: Merhback (1973) for S>20, for S<20: Edmond and Gieskes (1970)
            ' Sigma of residuals between fits and above data: Â±0.015, +0.040 for K1 and K2, respectively.
            ' Sal 0-40, Temp 0.2-30
            ' Limnol. Oceanogr. 43(4) (1998) 657-668
            ' On the NBS scale
            ' Their check values for F1 don't work out, not sure if this was correctly published...
            F1 = 200.1 / TempK + 0.322
            pK1 = 3404.71 / TempK + 0.032786 * TempK - 14.8435 - 0.071692 * F1 * Sal ^ 0.5 + 0.0021487 * Sal
            K1 = 10 ^ -pK1        ' this is on the NBS scale
                K1 = K1 / fH                ' convert to SWS scale (uncertain at low Sal due to junction potential)
            F2 = -129.24 / TempK + 1.4381
            pK2 = 2902.39 / TempK + 0.02379 * TempK - 6.498 - 0.3191 * F2 * Sal ^ 0.5 + 0.0198 * Sal
            K2 = 10 ^ -pK2       ' this is on the NBS scale
            K2 = K2 / fH                ' convert to SWS scale (uncertain at low Sal due to junction potential)
        
        Case 10    ' From Lueker, Dickson, Keeling, 2000
           ' This is Mehrbach's data refit after conversion to the total scale, for comparison with their equilibrator work.
         ' Mar. Chem. 70 (2000) 105-119
         ' Total scale and kg-sw
            pK1 = 3633.86 / TempK - 61.2172 + 9.6777 * Log(TempK) - 0.011555 * Sal + 0.0001152 * Sal ^ 2
            K1 = 10 ^ -pK1          ' this is on the total pH scale in mol/kg-SW
             K1 = K1 / SWStoTOT            ' convert to SWS pH scale
            pK2 = 471.78 / TempK + 25.929 - 3.16967 * Log(TempK) - 0.01781 * Sal + 0.0001122 * Sal ^ 2
            K2 = 10 ^ -pK2          ' this is on the total pH scale in mol/kg-SW
            K2 = K2 / SWStoTOT                ' convert to SWS pH scale

        Case 11    ' Mojica Prieto and Millero 2002. Geochim. et Cosmochim. Acta. 66(14) 2529-2540.
            ' sigma for pK1 is reported to be 0.0056
            ' sigma for pK2 is reported to be 0.010
            ' This is from the abstract and pages 2536-2537
            pK1 = -43.6977 - 0.0129037 * Sal + 0.0001364 * Sal ^ 2 + 2885.378 / TempK + 7.045159 * Log(TempK)
            pK2 = -452.094 + 13.142162 * Sal - 0.0008101 * Sal ^ 2 + 21263.61 / TempK + 68.483143 * Log(TempK) + _
                        (-581.4428 * Sal + 0.259601 * Sal ^ 2) / TempK - 1.967035 * Sal * Log(TempK)
            K1 = 10# ^ -pK1 ' this is on the SWS pH scale in mol/kg-SW
            K2 = 10# ^ -pK2 ' this is on the SWS pH scale in mol/kg-SW
        
        Case 12    ' Millero et al., 2002. Deep-Sea Res. I (49) 1705-1723.
            ' Calculated from overdetermined WOCE-era field measurements
            ' sigma for pK1 is reported to be 0.005
            ' sigma for pK2 is reported to be 0.008
            ' This is from page 1715
            pK1 = 6.359 - 0.00664 * Sal - 0.01322 * TempC + 0.00004989 * TempC ^ 2
            pK2 = 9.867 - 0.01314 * Sal - 0.01904 * TempC + 0.00002448 * TempC ^ 2
            K1 = 10# ^ -pK1 ' this is on the SWS pH scale in mol/kg-SW
            K2 = 10# ^ -pK2 ' this is on the SWS pH scale in mol/kg-SW
        
        Case 13    ' From Millero et al. 2006 work on pK1 and pK2 from titrations
            ' Millero, Graham, Huang, Bustos-Serrano, Pierrot. Mar.Chem. 100 (2006) 80-94.
            ' S=1 to 50, T=0 to 50. On seawater scale (SWS). From titrations in Gulf Stream seawater.
            pK10 = -126.34048 + 6320.813 / TempK + 19.568224 * Log(TempK)
            A1 = 13.4191 * Sal ^ 0.5 + 0.0331 * Sal - 0.0000533 * Sal ^ 2
            B1 = -530.123 * Sal ^ 0.5 - 6.103 * Sal
            C1 = -2.0695 * Sal ^ 0.5
            pK1 = A1 + B1 / TempK + C1 * Log(TempK) + pK10     ' pK1 sigma = 0.0054
            K1 = 10# ^ -(pK1)
            pK20 = -90.18333 + 5143.692 / TempK + 14.613358 * Log(TempK)
            A2 = 21.0894 * Sal ^ 0.5 + 0.1248 * Sal - 0.0003687 * Sal ^ 2
            b2 = -772.483 * Sal ^ 0.5 - 20.051 * Sal
            C2 = -3.3336 * Sal ^ 0.5
            pK2 = A2 + b2 / TempK + C2 * Log(TempK) + pK20        'pK2 sigma = 0.011
            K2 = 10# ^ -(pK2)
        
        Case 14    ' From Millero, 2010, also for estuarine use.
            ' Marine and Freshwater Research, v. 61, p. 139-142.
            ' Fits through compilation of real seawater titration results:
            ' Mehrbach et al. (1973), Mojica-Prieto & Millero (2002), Millero et al. (2006)
            ' Constants for K's on the SWS
            ' This is from page 141
            pK10 = -126.34048 + 6320.813 / TempK + 19.568224 * Log(TempK)
            ' This is from their table 2, page 140.
            A1 = 13.4038 * Sal ^ 0.5 + 0.03206 * Sal - 0.00005242 * Sal ^ 2
            B1 = -530.659 * Sal ^ 0.5 - 5.821 * Sal
            C1 = -2.0664 * Sal ^ 0.5
            pK1 = pK10 + A1 + B1 / TempK + C1 * Log(TempK)
            K1 = 10# ^ -pK1
            ' This is from page 141
            pK20 = -90.18333 + 5143.692 / TempK + 14.613358 * Log(TempK)
            ' This is from their table 3, page 140.
            A2 = 21.3728 * Sal ^ 0.5 + 0.1218 * Sal - 0.0003688 * Sal ^ 2
            b2 = -788.289 * Sal ^ 0.5 - 19.189 * Sal
            C2 = -3.374 * Sal ^ 0.5
            pK2 = pK20 + A2 + b2 / TempK + C2 * Log(TempK)
            K2 = 10# ^ -pK2
 
    End Select
'
'***************************************************************************
'***************************************************************************
CorrectKsForPressureNow:
' Currently: For WhichKs% = 1 to 7, all Ks (except KF and KS, which are on
'       the free scale) are on the SWS scale.
'       For WhichKs% = 6, KW set to 0, KP1, KP2, KP3, KSi don't matter.
'       For WhichKs% = 8, K1, K2, and KW are on the "pH" pH scale
'       (the pH scales are the same in this case); the other Ks don't matter.
'
'
' No salinity dependence is given for the pressure coefficients here.
' It is assumed that the salinity is at or very near Sal = 35.
' These are valid for the SWS pH scale, but the difference between this and
' the total only yields a difference of .004 pH units at 1000 bars, much
' less than the uncertainties in the values.
'
'
'****************************************************************************
' The sources used are:
' Millero, 1995:
'       Millero, F. J., Thermodynamics of the carbon dioxide system in the
'       oceans, Geochemica et Cosmochemica Acta 59:661-677, 1995.
'       See table 9 and eqs. 90-92, p. 675.
'       TYPO!!!: a factor of 10^3 was left out of the definition of Kappa
'       TYPO!!!: the value of R given is incorrect with the wrong units
'       TYPO!!!: the values of the a's for H2S and H2O are from the 1983
'                values for fresh water
'       TYPO!!!: the value of a1 for B(OH)3 should be +.1622
'       !!! Table 9 on p. 675 has no values for Si.
'       There are a variety of other typos in Table 9 on p. 675.
'       There are other typos in the paper, and most of the check values
'       given don't check.
' Millero, 1992:
'       Millero, Frank J., and Sohn, Mary L., Chemical Oceanography,
'       CRC Press, 1992. See chapter 6.
'       TYPO!!!: this chapter has numerous typos (eqs. 36, 52, 56, 65, 72,
'               79, and 96 have typos).
' Millero, 1983:
'       Millero, Frank J., Influence of pressure on chemical processes in
'       the sea. Chapter 43 in Chemical Oceanography, eds. Riley, J. P. and
'       Chester, R., Academic Press, 1983.
'       TYPO!!!: p. 51, eq. 94: the value -26.69 should be -25.59
'       TYPO!!!: p. 51, eq. 95: the term .1700t should be .0800t
'       these two are necessary to match the values given in Table 43.24
' Millero, 1979:
'       Millero, F. J., The thermodynamics of the carbon dioxide system
'       in seawater, Geochemica et Cosmochemica Acta 43:1651-1661, 1979.
'       See table 5 and eqs. 7, 7a, 7b on pp. 1656-1657.
' Takahashi et al, in GEOSECS Pacific Expedition, v. 3, 1982.
'       TYPO!!!: the pressure dependence of K2 should have a 16.4, not 26.4
'       This matches the GEOSECS results and is in Edmond and Gieskes.
' Culberson, C. H. and Pytkowicz, R. M., Effect of pressure on carbonic acid,
'       boric acid, and the pH of seawater, Limnology and Oceanography
'       13:403-417, 1968.
' Edmond, John M. and Gieskes, J. M. T. M., The calculation of the degree of
'       seawater with respect to calcium carbonate under in situ conditions,
'       Geochemica et Cosmochemica Acta, 34:1261-1291, 1970.
'
'
'****************************************************************************
' These references often disagree and give different fits for the same thing.
' They are not always just an update either; that is, Millero, 1995 may agree
'       with Millero, 1979, but differ from Millero, 1983.
' For 22 = 7 (Peng choice) I used the same factors for KW, KP1, KP2,
'       KP3, and KSi as for the other cases. Peng et al didn't consider the
'       case of P different from 0. GEOSECS did consider pressure, but didn't
'       include Phos, Si, or OH, so including the factors here won't matter.
' For WhichKs% = 8 (freshwater) the values are from Millero, 1983 (for K1, K2,
'       and KW). The other aren't used (TB = TS = TF = TP = TSi = 0.), so
'       including the factors won't matter.
'
'
'************************************************
CorrectK1K2KBForPressure:
        Select Case WhichKs%
        Case 1, 2, 3, 4, 5, 9, 10, 11, 12, 13, 14
PressureEffectsOnK1:
'               These are from Millero, 1995.
'               They are the same as Millero, 1979 and Millero, 1992.
'               They are from data of Culberson and Pytkowicz, 1968.
                deltaV = -25.5 + 0.1271 * TempC
                'deltaV = deltaV - .151 * (Sal - 34.8): ' Millero, 1979
                Kappa = (-3.08 + 0.0877 * TempC) / 1000!
                'Kappa = Kappa  - .578 * (Sal - 34.8)/1000.: ' Millero, 1979
                lnK1fac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'               The fits given in Millero, 1983 are somewhat different.
'
'
PressureEffectsOnK2:
'               These are from Millero, 1995.
'               They are the same as Millero, 1979 and Millero, 1992.
'               They are from data of Culberson and Pytkowicz, 1968.
                deltaV = -15.82 - 0.0219 * TempC
                'deltaV = deltaV + .321 * (Sal - 34.8): ' Millero, 1979
                Kappa = (1.13 - 0.1475 * TempC) / 1000!
                'Kappa = Kappa - .314 * (Sal - 34.8) / 1000!: ' Millero, 1979
                lnK2fac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'               The fit given in Millero, 1983 is different.
'               Not by a lot for deltaV, but by much for Kappa!!!. '
'
'
PressureEffectsOnKB:
'               This is from Millero, 1979.
'               It is from data of Culberson and Pytkowicz, 1968.
                deltaV = -29.48 + 0.1622 * TempC - 0.002608 * TempC * TempC
'               Millero, 1983 has:
                'deltaV = -28.56 + .1211 * TempC - .000321 * TempC * TempC
'               Millero, 1992 has:
                'deltaV = -29.48 + .1622 * TempC + .295 * (Sal - 34.8)
'               Millero, 1995 has:
                'deltaV = -29.48 - .1622 * TempC + .002608 * TempC * TempC
                'deltaV = deltaV + .295 * (Sal - 34.8): ' Millero, 1979
                Kappa = -2.84 / 1000!: ' Millero, 1979
'               Millero, 1992 and Millero, 1995 also have this.
                'Kappa = Kappa + .354 * (Sal - 34.8) / 1000!: ' Millero,1979
'               Millero, 1983 has:
                'Kappa = (-3! + .0427 * TempC) / 1000!
                lnKBfac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'
'
'**************************
        Case 6, 7
'               GEOSECS Pressure Effects On K1, K2, KB (on the NBS scale)
'               Takahashi et al, GEOSECS Pacific Expedition v. 3, 1982 quotes
'               Culberson and Pytkowicz, L and O 13:403-417, 1968:
'               but the fits are the same as those in
'               Edmond and Gieskes, GCA, 34:1261-1291, 1970
'               who in turn quote Li, personal communication
'
                lnK1fac = (24.2 - 0.085 * TempC) * Pbar / RT
                lnK2fac = (16.4 - 0.04 * TempC) * Pbar / RT
'               Takahashi et al had 26.4, but 16.4 is from Edmond and Gieskes
'               and matches the GEOSECS results
                lnKBfac = (27.5 - 0.095 * TempC) * Pbar / RT
'
'
'**************************
        Case 8
PressureEffectsOnK1inFreshWater:
'               This is from Millero, 1983.
                deltaV = -30.54 + 0.1849 * TempC - 0.0023366 * TempC * TempC
                Kappa = (-6.22 + 0.1368 * TempC - 0.001233 * TempC * TempC) / 1000!
                lnK1fac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'
'
PressureEffectsOnK2inFreshWater:
'               This is from Millero, 1983.
                deltaV = -29.81 + 0.115 * TempC - 0.001816 * TempC * TempC
                Kappa = (-5.74 + 0.093 * TempC - 0.001896 * TempC * TempC) / 1000!
                lnK2fac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'
'
                lnKBfac = 0! ': this doesn't matter since TB = 0 for this case
'
'
        End Select
'************************************************
CorrectKWForPressure:
        Select Case WhichKs%
        Case 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14
' GEOSECS doesn't include OH term, so this won't matter.
' Peng et al didn't include pressure, but here I assume that the KW correction
'       is the same as for the other seawater cases.
PressureEffectsOnKW:
'               This is from Millero, 1983 and his programs CO2ROY(T).BAS.
                deltaV = -20.02 + 0.1119 * TempC - 0.001409 * TempC * TempC
'               Millero, 1992 and Millero, 1995 have:
                'deltaV = -25.6 + .2324*TempC - .0036246*TempC*TempC
'               This is the freshwater value listed in Millero, 1983.
'               The difference is about 4 to 5 over the range 0 < TempC < 20,
'               which corresponds to a change in KW(P) of 3% at 200 bar,
'               8% at 500 bar, and 18% at 1000 bar.
'               This is probably correct since in Millero, 1983 values of
'               -deltaVs are less in seawater than pure water in all cases.
                Kappa = (-5.13 + 0.0794 * TempC) / 1000!: ' Millero, 1983
'               Millero, 1995 has this too, but Millero, 1992 is different.
                lnKWfac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'               Millero, 1979 does not list values for these.
'
'
        Case 8
PressureEffectsOnKWinFreshWater:
'               This is from Millero, 1983.
                deltaV = -25.6 + 0.2324 * TempC - 0.0036246 * TempC * TempC
                Kappa = (-7.33 + 0.1368 * TempC - 0.001233 * TempC * TempC) / 1000!
                lnKWfac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'               !!! NOTE the temperature dependence of KappaK1 and KappaKW
'               for fresh water in Millero, 1983 are the same.
'
'
        End Select
'************************************************
PressureEffectsOnKF:
'       This is from Millero, 1995, which is the same as Millero, 1983.
'       It is assumed that KF is on the free pH scale.
        deltaV = -9.78 - 0.009 * TempC - 0.000942 * TempC * TempC
        Kappa = (-3.91 + 0.054 * TempC) / 1000!
        lnKFfac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'
'
PressureEffectsOnKS:
'       This is from Millero, 1995, which is the same as Millero, 1983.
'       It is assumed that KS is on the free pH scale.
        deltaV = -18.03 + 0.0466 * TempC + 0.000316 * TempC * TempC
        Kappa = (-4.53 + 0.09 * TempC) / 1000!
        lnKSfac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'
'
'************************************************
CorrectKP1KP2KP3KSiForPressure:
' These corrections don't matter for the GEOSECS choice (WhichKs% = 6) and
'       the freshwater choice (WhichKs% = 8). For the Peng choice I assume
'       that they are the same as for the other choices (WhichKs% = 1 to 5).
' The corrections for KP1, KP2, and KP3 are from Millero, 1995, which are the
'       same as Millero, 1983.
'
'
PressureEffectsOnKP1:
        deltaV = -14.51 + 0.1211 * TempC - 0.000321 * TempC * TempC
        Kappa = (-2.67 + 0.0427 * TempC) / 1000!
        lnKP1fac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'
'
PressureEffectsOnKP2:
        deltaV = -23.12 + 0.1758 * TempC - 0.002647 * TempC * TempC
        Kappa = (-5.15 + 0.09 * TempC) / 1000!
        lnKP2fac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'
'
PressureEffectsOnKP3:
        deltaV = -26.57 + 0.202 * TempC - 0.003042 * TempC * TempC
        Kappa = (-4.08 + 0.0714 * TempC) / 1000!
        lnKP3fac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'
'
PressureEffectsOnKSi:
' !!!! The only mention of this is Millero, 1995 where it is stated that the
'       values have been estimated from the values of boric acid. HOWEVER,
'       there is no listing of the values in the table.
'       I used the values for boric acid from above.
        deltaV = -29.48 + 0.1622 * TempC - 0.002608 * TempC * TempC
        Kappa = -2.84 / 1000!
        lnKSifac = (-deltaV + 0.5 * Kappa * Pbar) * Pbar / RT
'
'
'************************************************
CorrectKsForPressureHere:
        K1fac = Exp(lnK1fac): K1 = K1 * K1fac
        K2fac = Exp(lnK2fac): K2 = K2 * K2fac
        KWfac = Exp(lnKWfac): KW = KW * KWfac
        KBfac = Exp(lnKBfac): KB = KB * KBfac
        KFfac = Exp(lnKFfac): KF = KF * KFfac
        KSfac = Exp(lnKSfac): KS = KS * KSfac
        KP1fac = Exp(lnKP1fac): KP1 = KP1 * KP1fac
        KP2fac = Exp(lnKP2fac): KP2 = KP2 * KP2fac
        KP3fac = Exp(lnKP3fac): KP3 = KP3 * KP3fac
        KSifac = Exp(lnKSifac): KSi = KSi * KSifac
'
'
'***************************************************************************
MakeKMatrix:
        K(1) = K1: K(2) = K2: K(3) = KW: K(4) = KB: K(5) = KF
        K(6) = KS: K(7) = KP1: K(8) = KP2: K(9) = KP3: K(10) = KSi
'
'
'************************************************
CorrectpHScaleConversionsForPressure:
'       fH has been assumed to be independent of pressure.
        SWStoTOT = (1! + TS / KS) / (1! + TS / KS + TF / KF)
        FREEtoTOT = 1! + TS / KS
'       The values KS and KF are already pressure-corrected, so the pH scale
'       conversions are now valid at pressure.
'
'
'************************************************
FindpHScaleConversionFactor:
        Select Case pHScale%: ' this is the scale they will be put on
                'Case "pH"
                        'there are only K1, K2, and KW and they should be ok
                 '       pHfactor = 1!
                Case 2 'SWS
                        ' they are all on this now
                        pHfactor = 1!
                Case 1 'total
                        pHfactor = SWStoTOT
                Case 3 '"pHfree"
                        pHfactor = SWStoTOT / FREEtoTOT
                Case 4 '"pHNBS"
                        pHfactor = fH
        End Select
'
'
'************************************************
ConvertFromSWSpHScaleToChosenScale:
        For ii% = 1 To 4
                K(ii%) = K(ii%) * pHfactor
        Next ii%
        ' KS and KF remain on the free pH scale
        For ii% = 7 To 10
                K(ii%) = K(ii%) * pHfactor
        Next ii%
'
'
'************************************************
' The constants should all be on the chosen pH scale at pressure.
'
'
'***************************************************************************
CalculateFugacityConstants:
' !!! This assumes that the pressure is at one atmosphere, or close to it.
' Otherwise, the Pres term in the exponent affects the results.
'       Weiss, R. F., Marine Chemistry 2:203-215, 1974.
'       Delta and B in cm3/mol
        Delta = (57.7 - 0.118 * TempK)
        B = -1636.75 + 12.0408 * TempK - 0.0327957 * TempK * TempK
        B = B + 3.16528 * 0.00001 * TempK * TempK * TempK
'
'
'       For a mixture of CO2 and air at 1 atm (at low CO2 concentrations):
        P1atm = 1.01325: ' in bar
        FugFac = Exp((B + 2! * Delta) * P1atm / RT)
'
'
'************************************************
        If WhichKs% = 6 Or WhichKs% = 7 Then FugFac = 1!
'       GEOSECS and Peng assume pCO2 = fCO2, or FugFac = 1
'
'
'****************************************************************************
CalculateVPFac:
' Weiss, R. F., and Price, B. A., Nitrous oxide solubility in water and
'       seawater, Marine Chemistry 8:347-359, 1980.
' They fit the data of Goff and Gratch (1946) with the vapor pressure
'       lowering by sea salt as given by Robinson (1954).
' This fits the more complicated Goff and Gratch, and Robinson equations
'       from 273 to 313 deg K and 0 to 40 Sal with a standard error
'       of .015%, about 5 uatm over this range.
' This may be on IPTS-29 since they didn't mention the temperature scale,
'       and the data of Goff and Gratch came before IPTS-48.
' The references are:
' Goff, J. A. and Gratch, S., Low pressure properties of water from -160 deg
'       to 212 deg F, Transactions of the American Society of Heating and
'       Ventilating Engineers 52:95-122, 1946.
' Robinson, Journal of the Marine Biological Association of the U. K.
'       33:449-455, 1954.
'
'
'       This is eq. 10 on p. 350.
'       This is in atmospheres.
        VPWP = Exp(24.4543 - 67.4509 * (100! / TempK) - 4.8489 * Log(TempK / 100!))
        VPCorrWP = Exp(-0.000544 * Sal)
        VPSWWP = VPWP * VPCorrWP
        VPFac = 1! - VPSWWP: ' this assumes 1 atmosphere
End Sub

Sub CalculateAlkParts(pH, TC, K(), T(), HCO3, CO3, BAlk, OH, PAlk, SiAlk, Hfree, HSO4, HF)
' SUB CalculateAlkParts, version 01.03, 10-10-97, written by Ernie Lewis.
' Inputs: pH, TC, K(), T()
' Outputs: HCO3, CO3, BAlk, OH, PAlk, SiAlk, Hfree, HSO4, HF
' This calculates the various contributions to the alkalinity.
' Though it is coded for H on the total pH scale, for the pH values occuring
' in seawater (pH > 6) it will be equally valid on any pH scale (H terms
' negligible) as long as the K constants are on that scale.
'
'
        K1 = K(1): K2 = K(2): KW = K(3): KB = K(4): KF = K(5)
        KS = K(6): KP1 = K(7): KP2 = K(8): KP3 = K(9): KSi = K(10)
        TB = T(1): TF = T(2): TS = T(3): TP = T(4): TSi = T(5)
'
'
        H = 10! ^ (-pH)
        HCO3 = TC * K1 * H / (K1 * H + H * H + K1 * K2)
        CO3 = TC * K1 * K2 / (K1 * H + H * H + K1 * K2)
        BAlk = TB * KB / (KB + H)
        OH = KW / H
                PhosTop = KP1 * KP2 * H + 2! * KP1 * KP2 * KP3 - H * H * H
                PhosBot = H * H * H + KP1 * H * H + KP1 * KP2 * H + KP1 * KP2 * KP3
        PAlk = TP * PhosTop / PhosBot
        ' this is good to better than .0006*TP:
                'PAlk = TP*(-H/(KP1+H) + KP2/(KP2+H) + KP3/(KP3+H))
        SiAlk = TSi * KSi / (KSi + H)
        FREEtoTOT = (1! + TS / KS): ' pH scale conversion factor
        Hfree = H / FREEtoTOT: ' for H on the total scale
        HSO4 = TS / (1! + KS / Hfree): ' since KS is on the free scale
        HF = TF / (1! + KF / Hfree): ' since KF is on the free scale
End Sub

Sub CalculatefCO2fromTCpH(TC, pH, K0, K1, K2, fCO2)
' SUB CalculatefCO2fromTCpH, version 02.02, 12-13-96, written by Ernie Lewis.
' Inputs: TC, pH, K0, K1, K2
' Output: fCO2
' This calculates fCO2 from TC and pH, using K0, K1, and K2.
'
'
        H = 10! ^ (-pH)
        fCO2 = TC * H * H / (H * H + K1 * H + K1 * K2) / K0
End Sub

Sub CalculatepHfromTAfCO2(TA, fCO2, K0, K(), T(), pH)
' SUB CalculatepHfromTAfCO2, version 04.01, 10-13-97, written by Ernie Lewis.
' Inputs: TA, fCO2, K0, K(), T()
' Output: pH
' This calculates pH from TA and fCO2 using K1 and K2 by Newton's method.
' It tries to solve for the pH at which Residual = 0.
' The starting guess is pH = 8.
' Though it is coded for H on the total pH scale, for the pH values occuring
' in seawater (pH > 6) it will be equally valid on any pH scale (H terms
' negligible) as long as the K constants are on that scale.
'
'
        K1 = K(1): K2 = K(2): KW = K(3): KB = K(4): KF = K(5)
        KS = K(6): KP1 = K(7): KP2 = K(8): KP3 = K(9): KSi = K(10)
        TB = T(1): TF = T(2): TS = T(3): TP = T(4): TSi = T(5)
'
'
        pHGuess = 8!: ' this is the first guess
        pHTol = 0.0001: ' this is .0001 pH units
        ln10 = Log(10!)
        pH = pHGuess
        Do
                H = 10! ^ (-pH)
                HCO3 = K0 * K1 * fCO2 / H
                CO3 = K0 * K1 * K2 * fCO2 / (H * H)
                CAlk = HCO3 + 2! * CO3
                BAlk = TB * KB / (KB + H)
                OH = KW / H
                        PhosTop = KP1 * KP2 * H + 2! * KP1 * KP2 * KP3 - H * H * H
                        PhosBot = H * H * H + KP1 * H * H + KP1 * KP2 * H + KP1 * KP2 * KP3
                PAlk = TP * PhosTop / PhosBot
                SiAlk = TSi * KSi / (KSi + H)
                FREEtoTOT = (1! + TS / KS): ' pH scale conversion factor
                Hfree = H / FREEtoTOT: ' for H on the total scale
                HSO4 = TS / (1! + KS / Hfree): ' since KS is on the free scale
                HF = TF / (1! + KF / Hfree): ' since KF is on the free scale
                Residual = TA - CAlk - BAlk - OH - PAlk - SiAlk + Hfree + HSO4 + HF
'
'               find Slope dTA/dpH
'               (this is not exact, but keeps all important terms):
                Slope = ln10 * (HCO3 + 4! * CO3 + BAlk * H / (KB + H) + OH + H)
                deltapH = Residual / Slope: ' this is Newton's method
                ' to keep the jump from being too big:
                Do While Abs(deltapH) > 1!: deltapH = deltapH / 2!: Loop
                pH = pH + deltapH
        Loop While Abs(deltapH) > pHTol
End Sub

Sub CalculatepHfromTATC(TA, TC, K(), T(), pH)
' SUB CalculatepHfromTATC, version 04.01, 10-13-96, written by Ernie Lewis.
' Inputs: TA, TC, K(), T()
' Output: pH
' This calculates pH from TA and TC using K1 and K2 by Newton's method.
' It tries to solve for the pH at which Residual = 0.
' The starting guess is pH = 8.
' Though it is coded for H on the total pH scale, for the pH values occuring
' in seawater (pH > 6) it will be equally valid on any pH scale (H terms
' negligible) as long as the K constants are on that scale.
'
'
        K1 = K(1): K2 = K(2): KW = K(3): KB = K(4): KF = K(5)
        KS = K(6): KP1 = K(7): KP2 = K(8): KP3 = K(9): KSi = K(10)
        TB = T(1): TF = T(2): TS = T(3): TP = T(4): TSi = T(5)
'
'
        pHGuess = 8!: ' this is the first guess
        pHTol = 0.0001: ' this is .0001 pH units
        ln10 = Log(10!)
        pH = pHGuess
        Do
                H = 10! ^ (-pH)
                Denom = (H * H + K1 * H + K1 * K2)
                CAlk = TC * K1 * (H + 2! * K2) / Denom
                BAlk = TB * KB / (KB + H)
                OH = KW / H
                        PhosTop = KP1 * KP2 * H + 2! * KP1 * KP2 * KP3 - H * H * H
                        PhosBot = H * H * H + KP1 * H * H + KP1 * KP2 * H + KP1 * KP2 * KP3
                PAlk = TP * PhosTop / PhosBot
                SiAlk = TSi * KSi / (KSi + H)
                FREEtoTOT = (1! + TS / KS): ' pH scale conversion factor
                Hfree = H / FREEtoTOT: ' for H on the total scale
                HSO4 = TS / (1! + KS / Hfree): ' since KS is on the free scale
                HF = TF / (1! + KF / Hfree): ' since KF is on the free scale
                Residual = TA - CAlk - BAlk - OH - PAlk - SiAlk + Hfree + HSO4 + HF
'
'               find Slope dTA/dpH:
'               (this is not exact, but keeps all important terms):
                Slope = ln10 * (TC * K1 * H * (H * H + K1 * K2 + 4! * H * K2) / Denom / Denom + BAlk * H / (KB + H) + OH + H)
                deltapH = Residual / Slope: ' this is Newton's method
                ' to keep the jump from being too big:
                Do While Abs(deltapH) > 1!: deltapH = deltapH / 2!: Loop
                pH = pH + deltapH
        Loop While Abs(deltapH) > pHTol
End Sub

Sub CalculatepHfromTCfCO2(TC, fCO2, K0, K1, K2, pH)
' SUB CalculatepHfromTCfCO2, version 02.02, 11-12-96, written by Ernie Lewis.
' Inputs: TC, fCO2, K0, K1, K2
' Output: pH
' This calculates pH from TC and fCO2 using K0, K1, and K2 by solving the
'       quadratic in H: fCO2 * K0 = TC * H * H / (K1 * H + H * H + K1 * K2).
' If there is not a real root, then pH is returned as -999.
'
'
        RR = K0 * fCO2 / TC
        If RR >= 1 Then pH = -999!: Exit Sub
        ' check after sub to see if pH = -999.
        Discr = (K1 * RR) * (K1 * RR) + 4! * (1! - RR) * (K1 * K2 * RR)
        H = 0.5 * (K1 * RR + Sqr(Discr)) / (1! - RR)
        If (H <= 0) Then
           pH = -999
        Else
           pH = Log(H) / Log(0.1)
        End If
End Sub

Sub CalculateTAfromTCpH(TC, pH, K(), T(), TA)
' SUB CalculateTAfromTCpH, version 02.02, 10-10-97, written by Ernie Lewis.
' Inputs: TC, pH, K(), T()
' Output: TA
' This calculates TA from TC and pH.
' Though it is coded for H on the total pH scale, for the pH values occuring
' in seawater (pH > 6) it will be equally valid on any pH scale (H terms
' negligible) as long as the K constants are on that scale.
'
'
        K1 = K(1): K2 = K(2): KW = K(3): KB = K(4): KF = K(5)
        KS = K(6): KP1 = K(7): KP2 = K(8): KP3 = K(9): KSi = K(10)
        TB = T(1): TF = T(2): TS = T(3): TP = T(4): TSi = T(5)
'
'
        H = 10! ^ (-pH)
        CAlk = TC * K1 * (H + 2! * K2) / (H * H + K1 * H + K1 * K2)
        BAlk = TB * KB / (KB + H)
        OH = KW / H
                PhosTop = KP1 * KP2 * H + 2! * KP1 * KP2 * KP3 - H * H * H
                PhosBot = H * H * H + KP1 * H * H + KP1 * KP2 * H + KP1 * KP2 * KP3
        PAlk = TP * PhosTop / PhosBot
        SiAlk = TSi * KSi / (KSi + H)
        FREEtoTOT = (1! + TS / KS): ' pH scale conversion factor
        Hfree = H / FREEtoTOT: ' for H on the total scale
        HSO4 = TS / (1! + KS / Hfree): ' since KS is on the free scale
        HF = TF / (1! + KF / Hfree): ' since KF is on the free scale
        TA = CAlk + BAlk + OH + PAlk + SiAlk - Hfree - HSO4 - HF
End Sub

Sub CalculateTCfrompHfCO2(pH, fCO2, K0, K1, K2, TC)
' SUB CalculateTCfrompHfCO2, version 01.02, 12-13-96, written by Ernie Lewis.
' Inputs: pH, fCO2, K0, K1, K2
' Output: TC
' This calculates TC from pH and fCO2, using K0, K1, and K2.
'
'
        H = 10! ^ (-pH)
        TC = K0 * fCO2 * (H * H + K1 * H + K1 * K2) / (H * H)
End Sub

Sub CalculateTCfromTApH(TA, pH, K(), T(), TC)
' SUB CalculateTCfromTApH, version 02.03, 10-10-97, written by Ernie Lewis.
' Inputs: TA, pH, K(), T()
' Output: TC
' This calculates TC from TA and pH.
' Though it is coded for H on the total pH scale, for the pH values occuring
' in seawater (pH > 6) it will be equally valid on any pH scale (H terms
' negligible) as long as the K constants are on that scale.
'
'
        K1 = K(1): K2 = K(2): KW = K(3): KB = K(4): KF = K(5)
        KS = K(6): KP1 = K(7): KP2 = K(8): KP3 = K(9): KSi = K(10)
        TB = T(1): TF = T(2): TS = T(3): TP = T(4): TSi = T(5)
'
'
        H = 10! ^ (-pH)
        BAlk = TB * KB / (KB + H)
        OH = KW / H
                PhosTop = KP1 * KP2 * H + 2! * KP1 * KP2 * KP3 - H * H * H
                PhosBot = H * H * H + KP1 * H * H + KP1 * KP2 * H + KP1 * KP2 * KP3
        PAlk = TP * PhosTop / PhosBot
        SiAlk = TSi * KSi / (KSi + H)
        FREEtoTOT = (1! + TS / KS): ' pH scale conversion factor
        Hfree = H / FREEtoTOT: ' for H on the total scale
        HSO4 = TS / (1! + KS / Hfree): ' since KS is on the free scale
        HF = TF / (1! + KF / Hfree): ' since KF is on the free scale
        CAlk = TA - BAlk - OH - PAlk - SiAlk + Hfree + HSO4 + HF
        TC = CAlk * (H * H + K1 * H + K1 * K2) / (K1 * (H + 2! * K2))
End Sub

Sub Case1Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TA, TC, pHinp, fCO2inp, pCO2inp, pHout, fCO2out, pCO2out)
' SUB Case1Partials, version 01.04, 03-12-97, written by Ernie Lewis.
' Inputs: pHScale%, WhichKs%, WhoseKSO4%, fORp$
' Inputs: Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout
' Inputs: TA, TC, pHinp, fCO2inp, pCO2inp, pHout, fCO2out, pCO2out
' Outputs: none
' This calculates and prints the partials for case 1: input TA, TC.
'
'
        TA0 = TA: TC0 = TC: Sal0 = Sal
        pHinp0 = pHinp: pHout0 = pHout
        fCO2inp0 = fCO2inp: pCO2inp0 = pCO2inp
        fCO2out0 = fCO2out: pCO2out0 = pCO2out
       ' Call SetParametersForPartials(dTA, dTC, dpH, dfCO2, dSal, dTempC, dPdbar, pcdK0, pcdK1, pcdK2)
' **************
'Increase TA by dTA
        TA = TA0 + dTA
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase1Partials:
                GoSub CalculateStuffForCase1Partials:
                dpHinpdTA = (pH - pHinp0) / dTA
                dfCO2inpdTA = (fCO2 - fCO2inp0) / dTA
                dpCO2inpdTA = (pCO2 - pCO2inp0) / dTA
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdTA = (pH - pHout0) / dTA
                dfCO2outdTA = (fCO2 - fCO2out0) / dTA
                dpCO2outdTA = (pCO2 - pCO2out0) / dTA
        TA = TA0
' **************
'Increase TC by dTC
        TC = TC0 + dTC
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase1Partials:
                GoSub CalculateStuffForCase1Partials:
                dpHinpdTC = (pH - pHinp0) / dTC
                dfCO2inpdTC = (fCO2 - fCO2inp0) / dTC
                dpCO2inpdTC = (pCO2 - pCO2inp0) / dTC
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdTC = (pH - pHout0) / dTC
                dfCO2outdTC = (fCO2 - fCO2out0) / dTC
                dpCO2outdTC = (pCO2 - pCO2out0) / dTC
        TC = TC0
' **************
'Increase Sal by dSal
        Sal = Sal0 + dSal
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase1Partials:
                GoSub CalculateStuffForCase1Partials:
                dpHinpdSal = (pH - pHinp0) / dSal
                dfCO2inpdSal = (fCO2 - fCO2inp0) / dSal
                dpCO2inpdSal = (pCO2 - pCO2inp0) / dSal
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdSal = (pH - pHout0) / dSal
                dfCO2outdSal = (fCO2 - fCO2out0) / dSal
                dpCO2outdSal = (pCO2 - pCO2out0) / dSal
        Sal = Sal0
' **************
'Increase TempCinp by dTempC
'       Do at Tinp, Pinp
                TempC = TempCinp + dTempC: Pdbar = Pdbarinp
                GoSub GetConstantsforCase1Partials:
                GoSub CalculateStuffForCase1Partials:
                dpHinpdTempCinp = (pH - pHinp0) / dTempC
                dfCO2inpdTempCinp = (fCO2 - fCO2inp0) / dTempC
                dpCO2inpdTempCinp = (pCO2 - pCO2inp0) / dTempC
'       Output results not affected by changes in Tinp in this case
' **************
'Increase Pdbar by dPdbar
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp + dPdbar
                GoSub GetConstantsforCase1Partials:
                GoSub CalculateStuffForCase1Partials:
                dpHinpdPdbarinp = (pH - pHinp0) / dPdbar
                dfCO2inpdPdbarinp = (fCO2 - fCO2inp0) / dPdbar
                dpCO2inpdPdbarinp = (pCO2 - pCO2inp0) / dPdbar
'       Output results not affected by changes in Pinp in this case
' **************
'Increase K0 by pcdK0 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase1Partials:
                K0 = K0 * (1! + pcdK0 / 100!)
                GoSub CalculateStuffForCase1Partials:
                ' pH doesn't depend on K0
                dfCO2inppcdK0 = (fCO2 - fCO2inp0) / pcdK0
                dpCO2inppcdK0 = (pCO2 - pCO2inp0) / pcdK0
'       Output results not affected by changes in K0 at input conditions in this case
'***************
'Increase K1 by pcdK1 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase1Partials:
                K(1) = K(1) * (1! + pcdK1 / 100!): K1 = K(1)
                GoSub CalculateStuffForCase1Partials:
                dpHinppcdK1 = (pH - pHinp0) / pcdK1
                dfCO2inppcdK1 = (fCO2 - fCO2inp0) / pcdK1
                dpCO2inppcdK1 = (pCO2 - pCO2inp0) / pcdK1
'       Output results not affected by changes in K1 at input conditions in this case
' **************
'Increase K2 by pcdK2 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase1Partials:
                K(2) = K(2) * (1! + pcdK2 / 100!): K2 = K(2)
                GoSub CalculateStuffForCase1Partials:
                dpHinppcdK2 = (pH - pHinp0) / pcdK2
                dfCO2inppcdK2 = (fCO2 - fCO2inp0) / pcdK2
                dpCO2inppcdK2 = (pCO2 - pCO2inp0) / pcdK2
'       Output results not affected by changes in K2 at input conditions in this case
'
'
' ******************************************
PrintPartialsForCase1:
        AAA$ = "#.####  ####.#           #.####  ####.# "
        If fORp$ = "f" Then
                ''Print USING; "                           \    \   fCO2            \    \   fCO2 "; pHScale%; pHScale%
                ''Print "    change per             ------  ------           ------  ------ "
                ''Print USING; "     \ /                   " + AAA$; pHinp0; fCO2inp0 * 1000000!; pHout0; fCO2out0 * 1000000!
        ElseIf fORp$ = "p" Then
                ''Print USING; "                           \    \   pCO2            \    \   pCO2 "; pHScale%; pHScale%
                ''Print "    change per             ------  ------           ------  ------ "
                ''Print USING; "     \ /                   " + AAA$; pHinp0; pCO2inp0 * 1000000!; pHout0; pCO2out0 * 1000000!
        End If
        ''Print USING; "    1 umol/kg in TA        " + AAA$; dpHinpdTA / 1000000!; dfCO2inpdTA; dpHoutdTA / 1000000!; dfCO2outdTA
        ''Print USING; "    1 umol/kg in TC        " + AAA$; dpHinpdTC / 1000000!; dfCO2inpdTC; dpHoutdTC / 1000000!; dfCO2outdTC
        If WhichKs% <> 8 Then
                ''Print USING; "    1 in salinity          " + AAA$; dpHinpdSal; dfCO2inpdSal * 1000000!; dpHoutdSal; dfCO2outdSal * 1000000!
        End If
        'Print USING; "    1 deg C in input T     #.####  ####.# "; dpHinpdTempCinp; dfCO2inpdTempCinp * 1000000!
        'Print USING; "    100 dbar in input P    #.####  ####.# "; dpHinpdPdbarinp * 100!; dfCO2inpdPdbarinp * 1000000! * 100!
        'Print USING; "    1% K0 at input T               ####.# "; dfCO2inppcdK0 * 1000000!
        'Print USING; "    1% K1 at input T, P    #.####  ####.# "; dpHinppcdK1; dfCO2inppcdK1 * 1000000!
        'Print USING; "    1% K2 at input T, P    #.####  ####.# "; dpHinppcdK2; dfCO2inppcdK2 * 1000000!
Exit Sub
'****************************************************************************
GetConstantsforCase1Partials:
        Call Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
        K1 = K(1): K2 = K(2)
Return
CalculateStuffForCase1Partials:
        If WhichKs% = 7 Then TA = TA - T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatepHfromTATC(TA, TC, K(), T(), pH)
        If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatefCO2fromTCpH(TC, pH, K0, K1, K2, fCO2): pCO2 = fCO2 / FugFac
Return
End Sub

Sub Case2Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TA, pHinp, TC, fCO2inp, pCO2inp, pHout, fCO2out, pCO2out)
' SUB Case2Partials, version 01.04, 03-12-97, written by Ernie Lewis.
' Inputs: pHScale%, WhichKs%, WhoseKSO4%, fORp$
' Inputs: Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout
' Inputs: TA, pHinp, TC, fCO2inp, pCO2inp, pHout, fCO2out, pCO2out
' Outputs: none
' This calculates and prints the partials for case 2: input TA, pH.
'
'
        TA0 = TA: TC0 = TC: Sal0 = Sal
        pHinp0 = pHinp: pHout0 = pHout
        fCO2inp0 = fCO2inp: pCO2inp0 = pCO2inp
        fCO2out0 = fCO2out: pCO2out0 = pCO2out
        'Call SetParametersForPartials(dTA, dTC, dpH, dfCO2, dSal, dTempC, dPdbar, pcdK0, pcdK1, pcdK2)
' **************
'Increase TA by dTA
        TA = TA0 + dTA
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase2Partials:
                pH = pHinp0
                GoSub CalculateStuffForCase2Partials:
                dTCdTA = (TC - TC0) / dTA
                dfCO2inpdTA = (fCO2 - fCO2inp0) / dTA
                dpCO2inpdTA = (pCO2 - pCO2inp0) / dTA
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdTA = (pH - pHout0) / dTA
                dfCO2outdTA = (fCO2 - fCO2out0) / dTA
                dpCO2outdTA = (pCO2 - pCO2out0) / dTA
        TA = TA0
' **************
'Increase pH by dpH (this is pH at input conditions)
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase2Partials:
                pH = pHinp0 + dpH
                GoSub CalculateStuffForCase2Partials:
                dTCdpH = (TC - TC0) / dpH
                dfCO2inpdpH = (fCO2 - fCO2inp0) / dpH
                dpCO2inpdpH = (pCO2 - pCO2inp0) / dpH
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdpH = (pH - pHout0) / dpH
                dfCO2outdpH = (fCO2 - fCO2out0) / dpH
                dpCO2outdpH = (pCO2 - pCO2out0) / dpH
' **************
'Increase Sal by dSal
        Sal = Sal0 + dSal
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase2Partials:
                pH = pHinp0
                GoSub CalculateStuffForCase2Partials:
                dTCdSal = (TC - TC0) / dSal
                dfCO2inpdSal = (fCO2 - fCO2inp0) / dSal
                dpCO2inpdSal = (pCO2 - pCO2inp0) / dSal
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdSal = (pH - pHout0) / dSal
                dfCO2outdSal = (fCO2 - fCO2out0) / dSal
                dpCO2outdSal = (pCO2 - pCO2out0) / dSal
        Sal = Sal0
' **************
'Increase TempCinp by dTempC
'       Do at Tinp, Pinp
                TempC = TempCinp + dTempC: Pdbar = Pdbarinp
                GoSub GetConstantsforCase2Partials:
                pH = pHinp0
                GoSub CalculateStuffForCase2Partials:
                dTCdTempCinp = (TC - TC0) / dTempC
                dfCO2inpdTempCinp = (fCO2 - fCO2inp0) / dTempC
                dpCO2inpdTempCinp = (pCO2 - pCO2inp0) / dTempC
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdTempCinp = (pH - pHout0) / dTempC
                dfCO2outdTempCinp = (fCO2 - fCO2out0) / dTempC
                dpCO2outdTempCinp = (pCO2 - pCO2out0) / dTempC
' **************
'Increase Pdbarinp by dPdbar
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp + dPdbar
                GoSub GetConstantsforCase2Partials:
                pH = pHinp0
                GoSub CalculateStuffForCase2Partials:
                dTCdPdbarinp = (TC - TC0) / dPdbar
                dfCO2inpdPdbarinp = (fCO2 - fCO2inp0) / dPdbar
                dpCO2inpdPdbarinp = (pCO2 - pCO2inp0) / dPdbar
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdPdbarinp = (pH - pHout0) / dPdbar
                dfCO2outdPdbarinp = (fCO2 - fCO2out0) / dPdbar
                dpCO2outdPdbarinp = (pCO2 - pCO2out0) / dPdbar
' **************
'Increase K0 by pcdK0 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase2Partials:
                K0 = K0 * (1! + pcdK0 / 100!)
                pH = pHinp0
                GoSub CalculateStuffForCase2Partials:
                ' TC does not depend on K0
                dfCO2inppcdK0 = (fCO2 - fCO2inp0) / pcdK0
                dpCO2inppcdK0 = (pCO2 - pCO2inp0) / pcdK0
'       Output results not affected by changes in K0 at input conditions in this case
' **************
'Increase K1 by pcdK1 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase2Partials:
                K(1) = K(1) * (1! + pcdK1 / 100!): K1 = K(1)
                pH = pHinp0
                GoSub CalculateStuffForCase2Partials:
                dTCpcdK1 = (TC - TC0) / pcdK1
                dfCO2inppcdK1 = (fCO2 - fCO2inp0) / pcdK1
                dpCO2inppcdK1 = (pCO2 - pCO2inp0) / pcdK1
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK1 = (pH - pHout0) / pcdK1
                dfCO2outpcdK1 = (fCO2 - fCO2out0) / pcdK1
                dpCO2outpcdK1 = (pCO2 - pCO2out0) / pcdK1
' **************
'Increase K2 by pcdK2 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase2Partials:
                K(2) = K(2) * (1! + pcdK2 / 100!): K2 = K(2)
                pH = pHinp0
                GoSub CalculateStuffForCase2Partials:
                dTCpcdK2 = (TC - TC0) / pcdK2
                dfCO2inppcdK2 = (fCO2 - fCO2inp0) / pcdK2
                dpCO2inppcdK2 = (pCO2 - pCO2inp0) / pcdK2
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK2 = (pH - pHout0) / pcdK2
                dfCO2outpcdK2 = (fCO2 - fCO2out0) / pcdK2
                dpCO2outpcdK2 = (pCO2 - pCO2out0) / pcdK2
'
'
' ******************************************
        AAA$ = "####.#  ####.#           #.####  ####.# "
PrintPartialsForCase2:
        If fORp$ = "f" Then
                'Print USING; "                             TC     fCO2            \    \   fCO2 "; pHScale%
                'Print "    change per             ------  ------           ------  ------ "
                'Print USING; "     \ /                   " + AAA$; TC0 * 1000000!; fCO2inp0 * 1000000!; pHout0; fCO2out0 * 1000000!
        ElseIf fORp$ = "p" Then
                'Print USING; "                             TC     pCO2            \    \   pCO2 "; pHScale%
                'Print "    change per             ------  ------           ------  ------ "
                'Print USING; "     \ /                   " + AAA$; TC0 * 1000000!; pCO2inp0 * 1000000!; pHout0; pCO2out0 * 1000000!
        End If
        'Print USING; "    1 umol/kg in TA        " + AAA$; dTCdTA; dfCO2inpdTA; dpHoutdTA / 1000000!; dfCO2outdTA
        'Print USING; "    .001 in input pH       " + AAA$; dTCdpH * 1000000! / 1000!; dfCO2inpdpH * 1000000! / 1000!; dpHoutdpH / 1000!; dfCO2outdpH * 1000000! / 1000!
        If WhichKs% <> 8 Then
                'Print USING; "    1 in salinity          " + AAA$; dTCdSal * 1000000!; dfCO2inpdSal * 1000000!; dpHoutdSal; dfCO2outdSal * 1000000!
        End If
        'Print USING; "    1 deg C in input T     " + AAA$; dTCdTempCinp * 1000000!; dfCO2inpdTempCinp * 1000000!; dpHoutdTempCinp; dfCO2outdTempCinp * 1000000!
        'Print USING; "    100 dbar in input P    " + AAA$; dTCdPdbarinp * 1000000! * 100!; dfCO2inpdPdbarinp * 1000000! * 100!; dpHoutdPdbarinp * 100!; dfCO2outdPdbarinp * 1000000! * 100!
        'Print USING; "    1% K0 at input T               ####.# "; dfCO2inppcdK0 * 1000000!
        'Print USING; "    1% K1 at input T, P    " + AAA$; dTCpcdK1 * 1000000!; dfCO2inppcdK1 * 1000000!; dpHoutpcdK1; dfCO2outpcdK1 * 1000000!
        'Print USING; "    1% K2 at input T, P    " + AAA$; dTCpcdK2 * 1000000!; dfCO2inppcdK2 * 1000000!; dpHoutpcdK2; dfCO2outpcdK2 * 1000000!
'
'
        TC = TC0: ' to pass back the value that came in
Exit Sub
'****************************************************************************
GetConstantsforCase2Partials:
        Call Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
        K1 = K(1): K2 = K(2)
Return
CalculateStuffForCase2Partials:
        If WhichKs% = 7 Then TA = TA - T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculateTCfromTApH(TA, pH, K(), T(), TC)
        If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatefCO2fromTCpH(TC, pH, K0, K1, K2, fCO2): pCO2 = fCO2 / FugFac
Return
End Sub

Sub Case3Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TA, fCO2inp, pCO2inp, TC, pHinp, pHout, fCO2out, pCO2out)
' SUB Case3Partials, version 01.04, 03-12-97, written by Ernie Lewis.
' Inputs: pHScale%, WhichKs%, WhoseKSO4%, fORp$
' Inputs: Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout
' Inputs: TA, fCO2inp, pCO2inp, TC, pHinp, pHout, fCO2out, pCO2out
' Outputs: none
' This calculates and prints the partials for case 3: input TA, fCO2 or pCO2.
'
'
End Sub

Sub Case4Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TC, pHinp, TA, fCO2inp, pCO2inp, pHout, fCO2out, pCO2out)
' SUB Case4Partials, version 01.04, 03-12-97, written by Ernie Lewis.
' Inputs: pHScale%, WhichKs%, WhoseKSO4%, fORp$
' Inputs: Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout
' Inputs: TC, pHinp, TA, fCO2inp, pCO2inp, pHout, fCO2out, pCO2out
' Outputs: none
' This calculates and prints the partials for case 4: input TC, pH.
'
'
        TA0 = TA: TC0 = TC: Sal0 = Sal
        pHinp0 = pHinp: pHout0 = pHout
        fCO2inp0 = fCO2inp: pCO2inp0 = pCO2inp
        fCO2out0 = fCO2out: pCO2out0 = pCO2out
        'Call SetParametersForPartials(dTA, dTC, dpH, dfCO2, dSal, dTempC, dPdbar, pcdK0, pcdK1, pcdK2)
' **************
'Increase TC by dTC
        TC = TC0 + dTC
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase4Partials:
                pH = pHinp0
                GoSub CalculateStuffForCase4Partials:
                dTAdTC = (TA - TA0) / dTC
                pCO2 = fCO2 / FugFac
                dfCO2inpdTC = (fCO2 - fCO2inp0) / dTC
                dpCO2inpdTC = (pCO2 - pCO2inp0) / dTC
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdTC = (pH - pHout0) / dTC
                dfCO2outdTC = (fCO2 - fCO2out0) / dTC
                dpCO2outdTC = (pCO2 - pCO2out0) / dTC
        TC = TC0
' **************
'Increase pH by dpH (this is pH at input conditions)
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase4Partials:
                pH = pHinp0 + dpH
                GoSub CalculateStuffForCase4Partials:
                dTAdpH = (TA - TA0) / dpH
                dfCO2inpdpH = (fCO2 - fCO2inp0) / dpH
                dpCO2inpdpH = (pCO2 - pCO2inp0) / dpH
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdpH = (pH - pHout0) / dpH
                dfCO2outdpH = (fCO2 - fCO2out0) / dpH
                dpCO2outdpH = (pCO2 - pCO2out0) / dpH
' **************
'Increase Sal by dSal
        Sal = Sal0 + dSal
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase4Partials:
                pH = pHinp0
                GoSub CalculateStuffForCase4Partials:
                dTAdSal = (TA - TA0) / dSal
                dfCO2inpdSal = (fCO2 - fCO2inp0) / dSal
                dpCO2inpdSal = (pCO2 - pCO2inp0) / dSal
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdSal = (pH - pHout0) / dSal
                dfCO2outdSal = (fCO2 - fCO2out0) / dSal
                dpCO2outdSal = (pCO2 - pCO2out0) / dSal
        Sal = Sal0
' **************
'Increase TempCinp by dTempC
'       Do at Tinp, Pinp
                TempC = TempCinp + dTempC: Pdbar = Pdbarinp
                GoSub GetConstantsforCase4Partials:
                pH = pHinp0
                GoSub CalculateStuffForCase4Partials:
                dTAdTempCinp = (TA - TA0) / dTempC
                dfCO2inpdTempCinp = (fCO2 - fCO2inp0) / dTempC
                dpCO2inpdTempCinp = (pCO2 - pCO2inp0) / dTempC
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdTempCinp = (pH - pHout0) / dTempC
                dfCO2outdTempCinp = (fCO2 - fCO2out0) / dTempC
                dpCO2outdTempCinp = (pCO2 - pCO2out0) / dTempC
' **************
'Increase Pdbarinp by dPdbar
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp + dPdbar
                GoSub GetConstantsforCase4Partials:
                pH = pHinp0
                GoSub CalculateStuffForCase4Partials:
                dTAdPdbarinp = (TA - TA0) / dPdbar
                dfCO2inpdPdbarinp = (fCO2 - fCO2inp0) / dPdbar
                dpCO2inpdPdbarinp = (pCO2 - pCO2inp0) / dPdbar
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdPdbarinp = (pH - pHout0) / dPdbar
                dfCO2outdPdbarinp = (fCO2 - fCO2out0) / dPdbar
                dpCO2outdPdbarinp = (pCO2 - pCO2out0) / dPdbar
' **************
'Increase K0 by pcdK0 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase4Partials:
                K0 = K0 * (1! + pcdK0 / 100!)
                pH = pHinp0
                GoSub CalculateStuffForCase4Partials:
                ' TA does not depend on K0 but I need TA ( = TA0) for output
                dfCO2inppcdK0 = (fCO2 - fCO2inp0) / pcdK0
                dpCO2inppcdK0 = (pCO2 - pCO2inp0) / pcdK0
'       Output results not affected by changes in K0 at input conditions in this case
' **************
'Increase K1 by pcdK1 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase4Partials:
                K(1) = K(1) * (1! + pcdK1 / 100!): K1 = K(1)
                pH = pHinp0
                GoSub CalculateStuffForCase4Partials:
                dTApcdK1 = (TA - TA0) / pcdK1
                dfCO2inppcdK1 = (fCO2 - fCO2inp0) / pcdK1
                dpCO2inppcdK1 = (pCO2 - pCO2inp0) / pcdK1
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK1 = (pH - pHout0) / pcdK1
                dfCO2outpcdK1 = (fCO2 - fCO2out0) / pcdK1
                dpCO2outpcdK1 = (pCO2 - pCO2out0) / pcdK1
' **************
'Increase K2 by pcdK2 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase4Partials:
                K(2) = K(2) * (1! + pcdK2 / 100!): K2 = K(2)
                pH = pHinp0
                GoSub CalculateStuffForCase4Partials:
                dTApcdK2 = (TA - TA0) / pcdK2
                dfCO2inppcdK2 = (fCO2 - fCO2inp0) / pcdK2
                dpCO2inppcdK2 = (pCO2 - pCO2inp0) / pcdK2
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK2 = (pH - pHout0) / pcdK2
                dfCO2outpcdK2 = (fCO2 - fCO2out0) / pcdK2
                dpCO2outpcdK2 = (pCO2 - pCO2out0) / pcdK2
'
'
' ******************************************
        AAA$ = "####.#  ####.#           #.####  ####.# "
PrintPartialsForCase4:
        If fORp$ = "f" Then
                'Print USING; "                             TA     fCO2            \    \   fCO2 "; pHScale%
                'Print "    change per             ------  ------           ------  ------ "
                'Print USING; "     \ /                   " + AAA$; TA0 * 1000000!; fCO2inp0 * 1000000!; pHout0; fCO2out0 * 1000000!
        ElseIf fORp$ = "p" Then
                'Print USING; "                             TA     pCO2            \    \   pCO2 "; pHScale%
                'Print "    change per             ------  ------           ------  ------ "
                'Print USING; "     \ /                   " + AAA$; TA0 * 1000000!; pCO2inp0 * 1000000!; pHout0; pCO2out0 * 1000000!
        End If
        'Print USING; "    1 umol/kg in TC        " + AAA$; dTAdTC; dfCO2inpdTC; dpHoutdTC / 1000000!; dfCO2outdTC
        'Print USING; "    .001 in input pH       " + AAA$; dTAdpH * 1000000! / 1000!; dfCO2inpdpH * 1000000! / 1000!; dpHoutdpH / 1000!; dfCO2outdpH * 1000000! / 1000!
        If WhichKs% <> 8 Then
                'Print USING; "    1 in salinity          " + AAA$; dTAdSal * 1000000!; dfCO2inpdSal * 1000000!; dpHoutdSal; dfCO2outdSal * 1000000!
        End If
        'Print USING; "    1 deg C in input T     " + AAA$; dTAdTempCinp * 1000000!; dfCO2inpdTempCinp * 1000000!; dpHoutdTempCinp; dfCO2outdTempCinp * 1000000!
        'Print USING; "    100 dbar in input P    " + AAA$; dTAdPdbarinp * 1000000! * 100!; dfCO2inpdPdbarinp * 1000000! * 100!; dpHoutdPdbarinp * 100!; dfCO2outdPdbarinp * 1000000! * 100!
        'Print USING; "    1% K0 at input T               ####.# "; dfCO2inppcdK0 * 1000000!
        'Print USING; "    1% K1 at input T, P    " + AAA$; dTApcdK1 * 1000000!; dfCO2inppcdK1 * 1000000!; dpHoutpcdK1; dfCO2outpcdK1 * 1000000!
        'Print USING; "    1% K2 at input T, P    " + AAA$; dTApcdK2 * 1000000!; dfCO2inppcdK2 * 1000000!; dpHoutpcdK2; dfCO2outpcdK2 * 1000000!
'
'
        TA = TA0: ' to pass back the value that came in
Exit Sub
'****************************************************************************
GetConstantsforCase4Partials:
        Call Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
        K1 = K(1): K2 = K(2)
Return
CalculateStuffForCase4Partials:
        Call CalculateTAfromTCpH(TC, pH, K(), T(), TA)
        If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatefCO2fromTCpH(TC, pH, K0, K1, K2, fCO2): pCO2 = fCO2 / FugFac
Return
End Sub

Sub Case5Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, TC, fCO2inp, pCO2inp, TA, pHinp, pHout, fCO2out, pCO2out, TCfCO2Flag%)
' SUB Case5Partials, version 01.04, 03-12-97, written by Ernie Lewis.
' Inputs: pHScale%, WhichKs%, WhoseKSO4%, fORp$
' Inputs: Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout
' Inputs: TC, fCO2inp, pCO2inp, TA, pHinp, pHout, fCO2out, pCO2out
' Inputs: TCfCO2Flag%
' Outputs: TCfCO2Flag%
' This calculates and prints the partials for case 5: input TC, fCO2 or pCO2.
'
'
        TA0 = TA: TC0 = TC: Sal0 = Sal
        pHinp0 = pHinp: pHout0 = pHout
        fCO2inp0 = fCO2inp
        fCO2out0 = fCO2out: pCO2out0 = pCO2out
        'Call SetParametersForPartials(dTA, dTC, dpH, dfCO2, dSal, dTempC, dPdbar, pcdK0, pcdK1, pcdK2)
' **************
'Increase TC by dTC
        TC = TC0 + dTC
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase5Partials:
                fCO2 = fCO2inp0
                GoSub CalculateStuffForCase5Partials:
                dpHinpdTC = (pH - pHinp0) / dTC
                dTAdTC = (TA - TA0) / dTC
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdTC = (pH - pHout0) / dTC
                dfCO2outdTC = (fCO2 - fCO2out0) / dTC
                dpCO2outdTC = (pCO2 - pCO2out0) / dTC
        TC = TC0
' **************
'Increase fCO2 by dfCO2 (this is fCO2 at input conditions)
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase5Partials:
'               this must be after call to constants for correct FugFac
                dpCO2 = dfCO2 / FugFac
                fCO2 = fCO2inp0 + dfCO2
                GoSub CalculateStuffForCase5Partials:
                dpHinpdfCO2 = (pH - pHinp0) / dfCO2
                dpHinpdpCO2 = (pH - pHinp0) / dpCO2
                dTAdfCO2 = (TA - TA0) / dfCO2
                dTAdpCO2 = (TA - TA0) / dpCO2
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdfCO2 = (pH - pHout0) / dfCO2
                dpHoutdpCO2 = (pH - pHout0) / dpCO2
                dfCO2outdfCO2 = (fCO2 - fCO2out0) / dfCO2
                dpCO2outdpCO2 = (pCO2 - pCO2out0) / dpCO2
' **************
'Increase Sal by dSal
        Sal = Sal0 + dSal
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase5Partials:
                fCO2 = fCO2inp0
'               since conversion of pCO2 to fCO2 depends on Sal, Temp:
                If fORp$ = "p" Then fCO2 = pCO2inp * FugFac
                GoSub CalculateStuffForCase5Partials:
                dTAdSal = (TA - TA0) / dSal
                dpHinpdSal = (pH - pHinp0) / dSal
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdSal = (pH - pHout0) / dSal
                dfCO2outdSal = (fCO2 - fCO2out0) / dSal
                dpCO2outdSal = (pCO2 - pCO2out0) / dSal
        Sal = Sal0
' **************
'Increase TempCinp by dTempC
'       Do at Tinp, Pinp
                TempC = TempCinp + dTempC: Pdbar = Pdbarinp
                GoSub GetConstantsforCase5Partials:
                fCO2 = fCO2inp0
'               since conversion of pCO2 to fCO2 depends on Sal, Temp:
                If fORp$ = "p" Then fCO2 = pCO2inp * FugFac
                GoSub CalculateStuffForCase5Partials:
                dTAdTempCinp = (TA - TA0) / dTempC
                dpHinpdTempCinp = (pH - pHinp0) / dTempC
'        Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdTempCinp = (pH - pHout0) / dTempC
                dfCO2outdTempCinp = (fCO2 - fCO2out0) / dTempC
                dpCO2outdTempCinp = (pCO2 - pCO2out0) / dTempC
' **************
'Increase Pdbarinp by dPdbar
'        Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp + dPdbar
                GoSub GetConstantsforCase5Partials:
                fCO2 = fCO2inp0
                GoSub CalculateStuffForCase5Partials:
                dTAdPdbarinp = (TA - TA0) / dPdbar
                dpHinpdPdbarinp = (pH - pHinp0) / dPdbar
'        Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdPdbarinp = (pH - pHout0) / dPdbar
                dfCO2outdPdbarinp = (fCO2 - fCO2out0) / dPdbar
                dpCO2outdPdbarinp = (pCO2 - pCO2out0) / dPdbar
' **************
'Increase K0 by pcdK0 % at input conditions only
'        Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase5Partials:
                K0 = K0 * (1! + pcdK0 / 100!)
                fCO2 = fCO2inp0
                GoSub CalculateStuffForCase5Partials:
                dpHinppcdK0 = (pH - pHinp0) / pcdK0
                dTApcdK0 = (TA - TA0) / pcdK0
'        Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK0 = (pH - pHout0) / pcdK0
                dfCO2outpcdK0 = (fCO2 - fCO2out0) / pcdK0
                dpCO2outpcdK0 = (pCO2 - pCO2out0) / pcdK0
' **************
'Increase K1 by pcdK1 % at input conditions only
'        Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase5Partials:
                K(1) = K(1) * (1! + pcdK1 / 100!): K1 = K(1)
                fCO2 = fCO2inp0
                GoSub CalculateStuffForCase5Partials:
                dTApcdK1 = (TA - TA0) / pcdK1
                dpHinppcdK1 = (pH - pHinp0) / pcdK1
'        Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK1 = (pH - pHout0) / pcdK1
                dfCO2outpcdK1 = (fCO2 - fCO2out0) / pcdK1
                dpCO2outpcdK1 = (pCO2 - pCO2out0) / pcdK1
' **************
'Increase K2 by pcdK2 % at input conditions only
'        Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase5Partials:
                K(2) = K(2) * (1! + pcdK2 / 100!): K2 = K(2)
                fCO2 = fCO2inp0
                GoSub CalculateStuffForCase5Partials:
                dpHinppcdK2 = (pH - pHinp0) / pcdK2
                dTApcdK2 = (TA - TA0) / pcdK2
'        Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK2 = (pH - pHout0) / pcdK2
                dfCO2outpcdK2 = (fCO2 - fCO2out0) / pcdK2
                dpCO2outpcdK2 = (pCO2 - pCO2out0) / pcdK2
'
'
' ******************************************
        AAA$ = "####.#  #.####           #.####  ####.# "
PrintPartialsForCase5:
        If fORp$ = "f" Then
                'Print USING; "                             TA    \    \           \    \   fCO2 "; pHScale%; pHScale%
                'Print "    change per             ------  ------           ------  ------ "
                'Print USING; "     \ /                   " + AAA$; TA0 * 1000000!; pHinp0; pHout0; fCO2out0 * 1000000!
                'Print USING; "    1 umol/kg in TC        " + AAA$; dTAdTC; dpHinpdTC / 1000000!; dpHoutdTC / 1000000!; dfCO2outdTC
                'Print USING; "    1 uatm in input fCO2   " + AAA$; dTAdfCO2; dpHinpdfCO2 / 1000000!; dpHoutdfCO2 / 1000000!; dfCO2outdfCO2
        ElseIf fORp$ = "p" Then
                'Print USING; "                             TA    \    \           \    \   pCO2 "; pHScale%; pHScale%
                'Print "    change per             ------  ------           ------  ------ "
                'Print USING; "     \ /                   " + AAA$; TA0 * 1000000!; pHinp0; pHout0; pCO2out0 * 1000000!
                'Print USING; "    1 umol/kg in TC        " + AAA$; dTAdTC; dpHinpdTC / 1000000!; dpHoutdTC / 1000000!; dpCO2outdTC
                'Print USING; "    1 uatm in input pCO2   " + AAA$; dTAdpCO2; dpHinpdpCO2 / 1000000!; dpHoutdpCO2 / 1000000!; dpCO2outdpCO2
        End If
        If WhichKs% <> 8 Then
                'Print USING; "    1 in salinity          " + AAA$; dTAdSal * 1000000!; dpHinpdSal; dpHoutdSal; dfCO2outdSal * 1000000!
        End If
        'Print USING; "    1 deg C in input T     " + AAA$; dTAdTempCinp * 1000000!; dpHinpdTempCinp; dpHoutdTempCinp; dfCO2outdTempCinp * 1000000!
        'Print USING; "    100 dbar in input P    " + AAA$; dTAdPdbarinp * 1000000! * 100!; dpHinpdPdbarinp * 100!; dpHoutdPdbarinp * 100!; dfCO2outdPdbarinp * 1000000! * 100!
        'Print USING; "    1% K0 at input T, P    " + AAA$; dTApcdK0 * 1000000!; dpHinppcdK0; dpHoutpcdK0; dfCO2outpcdK0 * 1000000!
        'Print USING; "    1% K1 at input T, P    " + AAA$; dTApcdK1 * 1000000!; dpHinppcdK1; dpHoutpcdK1; dfCO2outpcdK1 * 1000000!
        'Print USING; "    1% K2 at input T, P    " + AAA$; dTApcdK2 * 1000000!; dpHinppcdK2; dpHoutpcdK2; dfCO2outpcdK2 * 1000000!
'
'
        TA = TA0: ' to pass back the value that came in
Exit Sub
'****************************************************************************
GetConstantsforCase5Partials:
        Call Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
        K1 = K(1): K2 = K(2)
Return
CalculateStuffForCase5Partials:
        Call CalculatepHfromTCfCO2(TC, fCO2, K0, K1, K2, pH)
        If pH = -999! Then
                TCfCO2Flag% = 1
                TA = TA0: ' to pass back the value that came in
                Exit Sub
        '       this means that the TC, fCO2 combination becomes
        '       physically unrealizable during calculations in this sub
        End If
        Call CalculateTAfromTCpH(TC, pH, K(), T(), TA)
        If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
Return
End Sub

Sub Case6Partials(pHScale%, WhichKs%, WhoseKSO4%, fORp$, Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout, pHinp, fCO2inp, pCO2inp, TA, TC, pHout, fCO2out, pCO2out)
' SUB Case6Partials, version 01.04, 03-12-97, written by Ernie Lewis.
' Inputs: pHScale%, WhichKs%, WhoseKSO4%, fORp$
' Inputs: Sal, K(), T(), TempCinp, TempCout, Pdbarinp, Pdbarout
' Inputs: pHinp, fCO2inp, pCO2inp, TA, TC, pHout, fCO2out, pCO2out
' Outputs: none
' This calculates and prints the partials for case 6: input pH, fCO2 or pCO2.
'
'
        TA0 = TA: TC0 = TC: Sal0 = Sal
        pHinp0 = pHinp: pHout0 = pHout
        fCO2inp0 = fCO2inp
        fCO2out0 = fCO2out: pCO2out0 = pCO2out
        'Call SetParametersForPartials(dTA, dTC, dpH, dfCO2, dSal, dTempC, dPdbar, pcdK0, pcdK1, pcdK2)
' **************
'Increase pH by dpH (this is pH at input conditions)
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase6Partials:
                pH = pHinp0 + dpH: fCO2 = fCO2inp0
                GoSub CalculateStuffForCase6Partials:
                dTAdpH = (TA - TA0) / dpH
                dTCdpH = (TC - TC0) / dpH
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdpH = (pH - pHout0) / dpH
                dfCO2outdpH = (fCO2 - fCO2out0) / dpH
                dpCO2outdpH = (pCO2 - pCO2out0) / dpH
' **************
'Increase fCO2 by dfCO2 (this is fCO2 at input conditions)
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase6Partials:
'               this must be after call to constants for correct FugFac
                dpCO2 = dfCO2 / FugFac
                pH = pHinp0: fCO2 = fCO2inp0 + dfCO2
                GoSub CalculateStuffForCase6Partials:
                dTAdfCO2 = (TA - TA0) / dfCO2
                dTAdpCO2 = (TA - TA0) / dpCO2
                dTCdfCO2 = (TC - TC0) / dfCO2
                dTCdpCO2 = (TC - TC0) / dpCO2
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdfCO2 = (pH - pHout0) / dfCO2
                dpHoutdpCO2 = (pH - pHout0) / dpCO2
                dfCO2outdfCO2 = (fCO2 - fCO2out0) / dfCO2
                dpCO2outdpCO2 = (pCO2 - pCO2out0) / dpCO2
' **************
'Increase Sal by dSal
        Sal = Sal0 + dSal
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase6Partials:
                pH = pHinp0: fCO2 = fCO2inp0
'               since conversion of pCO2 to fCO2 depends on Sal, Temp:
                If fORp$ = "p" Then fCO2 = pCO2inp * FugFac
                GoSub CalculateStuffForCase6Partials:
                dTCdSal = (TC - TC0) / dSal
                dTAdSal = (TA - TA0) / dSal
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdSal = (pH - pHout0) / dSal
                dfCO2outdSal = (fCO2 - fCO2out0) / dSal
                dpCO2outdSal = (pCO2 - pCO2out0) / dSal
        Sal = Sal0
' **************
'Increase TempCinp by dTempC
'       Do at Tinp, Pinp
                TempC = TempCinp + dTempC: Pdbar = Pdbarinp
                GoSub GetConstantsforCase6Partials:
                pH = pHinp0: fCO2 = fCO2inp0
'               since conversion of pCO2 to fCO2 depends on Sal, Temp:
                If fORp$ = "p" Then fCO2 = pCO2inp * FugFac
                GoSub CalculateStuffForCase6Partials:
                dTAdTempCinp = (TA - TA0) / dTempC
                dTCdTempCinp = (TC - TC0) / dTempC
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdTempCinp = (pH - pHout0) / dTempC
                dfCO2outdTempCinp = (fCO2 - fCO2out0) / dTempC
                dpCO2outdTempCinp = (pCO2 - pCO2out0) / dTempC
' **************
'Increase Pdbarinp by dPdbar
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp + dPdbar
                GoSub GetConstantsforCase6Partials:
                pH = pHinp0: fCO2 = fCO2inp0
                GoSub CalculateStuffForCase6Partials:
                dTAdPdbarinp = (TA - TA0) / dPdbar
                dTCdPdbarinp = (TC - TC0) / dPdbar
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutdPdbarinp = (pH - pHout0) / dPdbar
                dfCO2outdPdbarinp = (fCO2 - fCO2out0) / dPdbar
                dpCO2outdPdbarinp = (pCO2 - pCO2out0) / dPdbar
' **************
'Increase K0 by pcdK0 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase6Partials:
                K0 = K0 * (1! + pcdK0 / 100!)
                pH = pHinp0: fCO2 = fCO2inp0
                GoSub CalculateStuffForCase6Partials:
                dTApcdK0 = (TA - TA0) / pcdK0
                dTCpcdK0 = (TC - TC0) / pcdK0
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK0 = (pH - pHout0) / pcdK0
                dfCO2outpcdK0 = (fCO2 - fCO2out0) / pcdK0
                dpCO2outpcdK0 = (pCO2 - pCO2out0) / pcdK0
' **************
'Increase K1 by pcdK1 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase6Partials:
                K(1) = K(1) * (1! + pcdK1 / 100!): K1 = K(1)
                pH = pHinp0: fCO2 = fCO2inp0
                GoSub CalculateStuffForCase6Partials:
                dTApcdK1 = (TA - TA0) / pcdK1
                dTCpcdK1 = (TC - TC0) / pcdK1
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK1 = (pH - pHout0) / pcdK1
                dfCO2outpcdK1 = (fCO2 - fCO2out0) / pcdK1
                dpCO2outpcdK1 = (pCO2 - pCO2out0) / pcdK1
' **************
'Increase K2 by pcdK2 % at input conditions only
'       Do at Tinp, Pinp
                TempC = TempCinp: Pdbar = Pdbarinp
                GoSub GetConstantsforCase6Partials:
                K(2) = K(2) * (1! + pcdK2 / 100!): K2 = K(2)
                pH = pHinp0: fCO2 = fCO2inp0
                GoSub CalculateStuffForCase6Partials:
                dTApcdK2 = (TA - TA0) / pcdK2
                dTCpcdK2 = (TC - TC0) / pcdK2
'       Do at Tout, Pout
                TempC = TempCout: Pdbar = Pdbarout
                Call FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
                dpHoutpcdK2 = (pH - pHout0) / pcdK2
                dfCO2outpcdK2 = (fCO2 - fCO2out0) / pcdK2
                dpCO2outpcdK2 = (pCO2 - pCO2out0) / pcdK2
'
'
' ******************************************
        AAA$ = "####.#  ####.#           #.####  ####.# "
PrintPartialsForCase6:
        If fORp$ = "f" Then
                'Print USING; "                             TA      TC             \    \   fCO2 "; pHScale%
                'Print "    change per             ------  ------           ------  ------ "
                'Print USING; "     \ /                   " + AAA$; TA0 * 1000000!; TC0 * 1000000!; pHout0; fCO2out0 * 1000000!
                'Print USING; "    .001 in input pH       " + AAA$; dTAdpH * 1000000! / 1000!; dTCdpH * 1000000! / 1000!; dpHoutdpH / 1000!; dfCO2outdpH * 1000000! / 1000!
                'Print USING; "    1 uatm in input fCO2   " + AAA$; dTAdfCO2; dTCdfCO2; dpHoutdfCO2 / 1000000!; dfCO2outdfCO2
        ElseIf fORp$ = "p" Then
                'Print USING; "                             TA      TC             \    \   pCO2 "; pHScale%
                'Print "    change per             ------  ------           ------  ------ "
                'Print USING; "     \ /                   " + AAA$; TA0 * 1000000!; TC0 * 1000000!; pHout0; pCO2out0 * 1000000!
                'Print USING; "    .001 in input pH       " + AAA$; dTAdpH * 1000000! / 1000!; dTCdpH * 1000000! / 1000!; dpHoutdpH / 1000!; dpCO2outdpH * 1000000! / 1000!
                'Print USING; "    1 uatm in input pCO2   " + AAA$; dTAdpCO2; dTCdpCO2; dpHoutdpCO2 / 1000000!; dpCO2outdpCO2
        End If
        If WhichKs% <> 8 Then
                'Print USING; "    1 in salinity          " + AAA$; dTAdSal * 1000000!; dTCdSal * 1000000!; dpHoutdSal; dfCO2outdSal * 1000000!
        End If
        'Print USING; "    1 deg C in input T     " + AAA$; dTAdTempCinp * 1000000!; dTCdTempCinp * 1000000!; dpHoutdTempCinp; dfCO2outdTempCinp * 1000000!
        'Print USING; "    100 dbar in input P    " + AAA$; dTAdPdbarinp * 1000000! * 100!; dTCdPdbarinp * 1000000! * 100!; dpHoutdPdbarinp * 100!; dfCO2outdPdbarinp * 1000000! * 100!
        'Print USING; "    1% K0 at input T       " + AAA$; dTApcdK0 * 1000000!; dTCpcdK0 * 1000000!; dpHoutpcdK0; dfCO2outpcdK0 * 1000000!
        'Print USING; "    1% K1 at input T, P    " + AAA$; dTApcdK1 * 1000000!; dTCpcdK1 * 1000000!; dpHoutpcdK1; dfCO2outpcdK1 * 1000000!
        'Print USING; "    1% K2 at input T, P    " + AAA$; dTApcdK2 * 1000000!; dTCpcdK2 * 1000000!; dpHoutpcdK2; dfCO2outpcdK2 * 1000000!
'
'
        TA = TA0: TC = TC0: ' to pass back the values that came in
Exit Sub
'****************************************************************************
GetConstantsforCase6Partials:
        Call Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
        K1 = K(1): K2 = K(2)
Return
CalculateStuffForCase6Partials:
        Call CalculateTCfrompHfCO2(pH, fCO2, K0, K1, K2, TC)
        Call CalculateTAfromTCpH(TC, pH, K(), T(), TA)
        If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
Return
End Sub

Sub CaSolubility(WhichKs%, Sal, TempC, Pdbar, TC, pH, K1, K2, OmegaCa, OmegaAr)
' SUB CaSolubility, version 01.05, 05-23-97, written by Ernie Lewis.
' Inputs: WhichKs%, Sal, TempC, Pdbar, TC, pH, K1, K2
' Outputs: OmegaCa, OmegaAr
' This calculates omega, the solubility ratio, for calcite and aragonite.
' This is defined by: Omega = [CO3--]*[Ca++] / Ksp,
'       where Ksp is the solubility product (either KCa or KAr).
'
'
'****************************************************************************
' These are from:
' Mucci, Alphonso, The solubility of calcite and aragonite in seawater
'       at various salinities, temperatures, and one atmosphere total
'       pressure, American Journal of Science 283:781-799, 1983.
' Ingle, S. E., Solubility of calcite in the ocean,
'       Marine Chemistry 3:301-319, 1975,
' Millero, Frank, The thermodynamics of the carbonate system in seawater,
'       Geochemica et Cosmochemica Acta 43:1651-1661, 1979.
' Ingle et al, The solubility of calcite in seawater at atmospheric pressure
'       and 35%o salinity, Marine Chemistry 1:295-307, 1973.
' Berner, R. A., The solubility of calcite and aragonite in seawater in
'       atmospheric pressure and 34.5%o salinity, American Journal of
'       Science 276:713-730, 1976.
' Takahashi et al, in GEOSECS Pacific Expedition, v. 3, 1982.
' Culberson, C. H. and Pytkowicz, R. M., Effect of pressure on carbonic acid,
'       boric acid, and the pH of seawater, Limnology and Oceanography
'       13:403-417, 1968.
'
'
'****************************************************************************
        RGasConstant = 83.1451: 'bar-cm3/(mol-K)
        TempK = TempC + 273.15
        RT = RGasConstant * TempK
        Pbar = Pdbar / 10!
        logTempK = Log(TempK)
        sqrSal = Sqr(Sal)
'       deltaVs are in cm3/mole
'       Kappas are in cm3/mole/bar
'       PROGRAMMER'S NOTE: all logs are log base e
'
'
'****************************************************************************
CalculateCa:
'       Riley, J. P. and Tongudai, M., Chemical Geology 2:263-269, 1967:
        Ca = 0.02128 / 40.087 * (Sal / 1.80655): ' in mol/kg-SW
'       this is .010285 * Sal / 35
'
'
CalciteSolubility:
'       Mucci, Alphonso, Amer. J. of Science 283:781-799, 1983.
        logKCa = -171.9065 - 0.077993 * TempK + 2839.319 / TempK
        logKCa = logKCa + 71.595 * logTempK / Log(10!)
        logKCa = logKCa + (-0.77712 + 0.0028426 * TempK + 178.34 / TempK) * sqrSal
        logKCa = logKCa - 0.07711 * Sal + 0.0041249 * sqrSal * Sal
'       sd fit = .01 (for Sal part, not part independent of Sal)
        KCa = 10! ^ (logKCa): ' this is in (mol/kg-SW)^2
'
'
AragoniteSolubility:
'       Mucci, Alphonso, Amer. J. of Science 283:781-799, 1983.
        logKAr = -171.945 - 0.077993 * TempK + 2903.293 / TempK
        logKAr = logKAr + 71.595 * logTempK / Log(10!)
        logKAr = logKAr + (-0.068393 + 0.0017276 * TempK + 88.135 / TempK) * sqrSal
        logKAr = logKAr - 0.10018 * Sal + 0.0059415 * sqrSal * Sal
'       sd fit = .009 (for Sal part, not part independent of Sal)
        KAr = 10! ^ (logKAr): ' this is in (mol/kg-SW)^2
'
'
PressureCorrectionForCalcite:
'       Ingle, Marine Chemistry 3:301-319, 1975
'       same as in Millero, GCA 43:1651-1661, 1979, but Millero, GCA 1995
'       has typos (-.5304, -.3692, and 10^3 for Kappa factor)
        deltaVKCa = -48.76 + 0.5304 * TempC
        KappaKCa = (-11.76 + 0.3692 * TempC) / 1000!
        lnKCafac = (-deltaVKCa + 0.5 * KappaKCa * Pbar) * Pbar / RT
        KCa = KCa * Exp(lnKCafac)
'
'
PressureCorrectionForAragonite:
'       Millero, Geochemica et Cosmochemica Acta 43:1651-1661, 1979,
'       same as Millero, GCA 1995 except for typos (-.5304, -.3692,
'       and 10^3 for Kappa factor)
        deltaVKAr = deltaVKCa + 2.8
        KappaKAr = KappaKCa
        lnKArfac = (-deltaVKAr + 0.5 * KappaKAr * Pbar) * Pbar / RT
        KAr = KAr * Exp(lnKArfac)
'
'
'****************************************************************************
If WhichKs% = 6 Or WhichKs% = 7 Then
CalculateCaforGEOSECS:
'       Culkin, F, in Chemical Oceanography, ed. Riley and Skirrow, 1965:
'       (quoted in Takahashi et al, GEOSECS Pacific Expedition v. 3, 1982)
        Ca = 0.01026 * Sal / 35!
'       Culkin gives Ca = (.0213 / 40.078) * (Sal / 1.80655) in mol/kg-SW
'       which corresponds to Ca = .01030 * Sal / 35.
'
'
CalculateKCaforGEOSECS:
'       Ingle et al, Marine Chemistry 1:295-307, 1973 is referenced in
'       (quoted in Takahashi et al, GEOSECS Pacific Expedition v. 3, 1982
'       but the fit is actually from Ingle, Marine Chemistry 3:301-319, 1975)
        KCa = 0.0000001 * (-34.452 - 39.866 * Sal ^ (1 / 3) + 110.21 * Log(Sal) / Log(10!) - 0.0000075752 * TempK * TempK)
'       this is in (mol/kg-SW)^2
'
'
CalculateKArforGEOSECS:
'       Berner, R. A., American Journal of Science 276:713-730, 1976:
'       (quoted in Takahashi et al, GEOSECS Pacific Expedition v. 3, 1982)
        KAr = 1.45 * KCa: ' this is in (mol/kg-SW)^2
'       Berner (p. 722) states that he uses 1.48.
'       It appears that 1.45 was used in the GEOSECS calculations
'
'
CalculatePressureEffectsOnKCaKArGEOSECS:
'       Culberson and Pytkowicz, Limnology and Oceanography 13:403-417, 1968
'       (quoted in Takahashi et al, GEOSECS Pacific Expedition v. 3, 1982
'       but their paper is not even on this topic).
'       The fits appears to be new in the GEOSECS report.
'       I can't find them anywhere else.
        KCa = KCa * Exp((36! - 0.2 * TempC) * Pbar / RT)
        KAr = KAr * Exp((33.3 - 0.22 * TempC) * Pbar / RT)
End If
'
'
'****************************************************************************
CalculateOmegasHere:
        H = 10! ^ (-pH)
        CO3 = TC * K1 * K2 / (K1 * H + H * H + K1 * K2)
        OmegaCa = CO3 * Ca / KCa: 'dimensionless
        OmegaAr = CO3 * Ca / KAr: 'dimensionless
End Sub

Sub FindpHfCO2fromTATC(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar, pH, fCO2, pCO2)
' SUB FindpHfCO2fromTATC, version 01.02, 10-10-97, written by Ernie Lewis.
' Inputs: pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempC, Pdbar
' Outputs: pH, fCO2, pCO2
' This calculates pH, fCO2, and pCO2 from TA and TC at output conditions.
'
'
        Call Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
        K1 = K(1): K2 = K(2)
'
        If WhichKs% = 7 Then TA = TA - T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatepHfromTATC(TA, TC, K(), T(), pH)
        If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
ICI:
        Call CalculatefCO2fromTCpH(TC, pH, K0, K1, K2, fCO2): pCO2 = fCO2 / FugFac
End Sub

Sub PrintpHspKs(pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T(), TempCinp, Pdbarinp, TempCout, Pdbarout)
' SUB PrintpHspKs, version 02.01, 10-10-97, written by Ernie Lewis.
' Inputs: pHScale%, WhichKs%, WhoseKSO4%, TA, TC, Sal, K(), T()
' Inputs:  TempCinp, Pdbarinp, TempCout, Pdbarout
' Outputs: none
' This calculates and prints the pH on all scales, and pK1, pK2, pKW, and pKB
'       on the given scale pHScale%.
'
'
FindpHsAndpKsAtInputConditions:
        TempC = TempCinp: Pdbar = Pdbarinp
        GoSub FindpHspKs:
        pHinp = pH: fHinp = fH: pHNBSinp = pHNBS
        pHfreeinp = pHfree: pHtotinp = pHtot: pHswsinp = pHsws
        pK1inp = pK1: pK2inp = pK2: pKWinp = pKW: pKBinp = pKB
'
'
FindpHsAndpKsAtOutputConditions:
        TempC = TempCout: Pdbar = Pdbarout
        GoSub FindpHspKs:
        pHout = pH: fHout = fH: pHNBSout = pHNBS
        pHfreeout = pHfree: pHtotout = pHtot: pHswsout = pHsws
        pK1out = pK1: pK2out = pK2: pKWout = pKW: pKBout = pKB
'
'
        S10$ = "          "
        AA2$ = "      ##.###                  ##.### "
        AA1$ = S10$ + AA2$
        If WhichKs% = 8 Then
                'Print USING; S10$ + "  pH  " + AA1$; pHinp; pHout
                'Print
                'Print USING; S10$ + "  pK1 " + AA1$; pK1inp; pK1out
                'Print USING; S10$ + "  pK2 " + AA1$; pK2inp; pK2out
                'Print USING; S10$ + "  pKW " + AA1$; pKWinp; pKWout
                Exit Sub
        End If
'
'
        'Print USING; "       pHtot (mol/kg-SW)  " + AA2$; pHtotinp; pHtotout
        'Print USING; "       pHsws (mol/kg-SW)  " + AA2$; pHswsinp; pHswsout
        'Print USING; "       pHfree (mol/kg-SW) " + AA2$; pHfreeinp; pHfreeout
        'Print USING; "       pHNBS (mol/kg-H2O) " + AA2$; pHNBSinp; pHNBSout
        'Print USING; "       fH                 " + AA2$; fHinp; fHout
        'Print
        'Print
        'Print "    These are on the "; pHScale%;: 'Print " scale ";
        Select Case pHScale%
                Case 1, 2, 3 '"pHtot", "pHsws", "pHfree"
                        'Print "(mol/kg-SW): "
                Case 4   '"pHNBS"
                        'Print "(mol/kg-H2O): "
        End Select
        'Print USING; S10$ + "  pK1 " + AA1$; pK1inp; pK1out
        'Print USING; S10$ + "  pK2 " + AA1$; pK2inp; pK2out
        If WhichKs% <> 6 Then
                'Print USING; S10$ + "  pKW " + AA1$; pKWinp; pKWout
'               GEOSECS doesn't include OH so KW is carried as 0 in this case
        End If
        'Print USING; S10$ + "  pKB " + AA1$; pKBinp; pKBout
Exit Sub
'****************************************************************************
FindpHspKs:
        Call Constants(pHScale%, WhichKs%, WhoseKSO4%, Sal, TempC, Pdbar, K0, K(), T(), fH, FugFac, VPFac)
        K1 = K(1): K2 = K(2)
        If WhichKs% = 7 Then TA = TA - T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatepHfromTATC(TA, TC, K(), T(), pH)
        If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call FindpHOnAllScales(pHScale%, pH, K(), T(), fH, pHNBS, pHfree, pHtot, pHsws)
        pK1 = Log(K1) / Log(0.1)
        pK2 = Log(K2) / Log(0.1)
        If WhichKs% <> 8 Then pKB = Log(K(4)) / Log(0.1)
        If WhichKs% <> 6 Then pKW = Log(K(3)) / Log(0.1)
'               GEOSECS doesn't include OH so KW is carried as 0 in this case
Return
End Sub

Sub FindpHOnAllScales(pHScale%, pH, K(), T(), fH, pHNBS, pHfree, pHtot, pHsws)
' SUB FindpHOnAllScales, version 01.02, 01-08-97, written by Ernie Lewis.
' Inputs: pHScale%, pH, K(), T(), fH
' Outputs: pHNBS, pHfree, pHTot, pHSWS
' This takes the pH on the given scale and finds the pH on all scales.
'
'
        TS = T(3): TF = T(2)
        KS = K(6): KF = K(5): 'these are at the given T, S, P
        FREEtoTOT = (1! + TS / KS): ' pH scale conversion factor
        SWStoTOT = (1! + TS / KS) / (1! + TS / KS + TF / KF): ' pH scale conversion factor
        Select Case pHScale%: ' this is the pH scale pH is on now
                Case 4  '"pHNBS"
                        factor = -Log(SWStoTOT) / Log(0.1) + Log(fH) / Log(0.1)
                Case 3  '"pHfree"
                        factor = -Log(FREEtoTOT) / Log(0.1)
                Case 1  '"pHtot"
                        factor = 0!
                Case 2  '"pHsws"
                        factor = -Log(SWStoTOT) / Log(0.1)
        End Select
        pHtot = pH - factor: ' pH comes into this sub on the given scale
        pHNBS = pHtot - Log(SWStoTOT) / Log(0.1) + Log(fH) / Log(0.1)
        pHfree = pHtot - Log(FREEtoTOT) / Log(0.1)
        pHsws = pHtot - Log(SWStoTOT) / Log(0.1)
End Sub

Sub RevelleFactor(WhichKs%, TA, TC, K0, K(), T(), Revelle)
' SUB RevelleFactor, version 01.03, 01-07-97, written by Ernie Lewis.
' Inputs: WhichKs%, TA, TC, K0, K(), T()
' Outputs: Revelle
' This calculates the Revelle factor (dfCO2/dTC)|TA/(fCO2/TC).
' It only makes sense to talk about it at pTot = 1 atm, but it is computed
'       here at the given K(), which may be at pressure <> 1 atm. Care must
'       thus be used to see if there is any validity to the number computed.
'
'
        If TC = 0! Then Revelle = 0!: Exit Sub
        K1 = K(1): K2 = K(2)
        TC0 = TC
        dTC = 0.000001: ' 1 umol/kg-SW
'
'
' Find fCO2 at TA, TC + dTC
        TC = TC0 + dTC
        GoSub GetfCO2:
        fCO2plus = fCO2
'
'
' Find fCO2 at TA, TC - dTC
        TC = TC0 - dTC
        GoSub GetfCO2:
        fCO2minus = fCO2
'
CalculateRevelleFactor:
        Revelle = (fCO2plus - fCO2minus) / dTC / ((fCO2plus + fCO2minus) / TC)
        ' at constant TA
'
'
ResetTC:
        TC = TC0
Exit Sub
'****************************************************************************
GetfCO2:
        If WhichKs% = 7 Then TA = TA - T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatepHfromTATC(TA, TC, K(), T(), pH)
        If WhichKs% = 7 Then TA = TA + T(4): ' PAlk(Peng) = PAlk(Dickson) + TP
        Call CalculatefCO2fromTCpH(TC, pH, K0, K1, K2, fCO2)
Return
End Sub
