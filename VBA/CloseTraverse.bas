Attribute VB_Name = "CloseTraverse"
' Topic; Close Traverse Program
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 06/06/2025
'
'Option Base 1
Const Pi As Single = 3.141592654

'-------------------------------- Private Function --------------------------------'

'Convert Degrees to Radian.

Private Function DegtoRad(d)

    DegtoRad = d * (Pi / 180)

 End Function

'Convert Radian to Degrees.

 Private Function RadtoDeg(r)

    RadtoDeg = r * (180 / Pi)

 End Function
Private Function DMStoDeg(d, m, s)
 
    dd = d
    mm = m / 60
    ss = s / 3600
    
    DMStoDeg = dd + mm + ss

End Function

Private Function DegtoDMS(deg, DMS)

    dd = Int(deg)
    mm = Int((deg - Int(deg)) * 60)
    ss = (((deg - Int(deg)) * 60) - Int((deg - Int(deg)) * 60)) * 60
    
    Select Case UCase$(DMS)
        Case "D"
            DegtoDMS = dd
        Case "M"
            DegtoDMS = mm
        Case "S"
            DegtoDMS = ss
    End Select

End Function

Function DegtoDMSStr(degrees)
        deg = Abs(degrees)
        
        dd = Int(deg)
        mm = Int((deg - Int(deg)) * 60)
        ss = (((deg - Int(deg)) * 60) - Int((deg - Int(deg)) * 60)) * 60
        
        If degrees >= 0 Then
            DegtoDMSStr = " " & Application.Text(dd, "000") & ChrW(&HB0) & " " & Application.Text(mm, "00") & "' " & Application.Text(ss, "00.00") & """"
        Else
            DegtoDMSStr = "- " & Application.Text(dd, "000") & ChrW(&HB0) & " " & Application.Text(mm, "00") & "' " & Application.Text(ss, "00.00") & """"
        End If

End Function


Private Function DirecDistAz(EStart, NStart, EEnd, NEnd, DA)
    Dim De As Single
    Dim dN As Single
    Dim Q As Single
    
    De = EEnd - EStart: dN = NEnd - NStart
    
    Distance = Sqr(De ^ 2 + dN ^ 2)
    
    If dN <> 0 Then Q = RadtoDeg(Atn(De / dN))
      If dN = 0 Then
        If De > 0 Then
          AZi = 90
        ElseIf De < 0 Then
          AZi = 270
        Else
          AZi = False
      End If
    ElseIf dN > 0 Then
      If De > 0 Then
          AZi = Q
      ElseIf De < 0 Then
          AZi = 360 + Q
      End If
    ElseIf dN < 0 Then
          AZi = 180 + Q
    End If
    
    Select Case UCase$(DA)
      Case "D"
          DirecDistAz = Distance
      Case "A"
          DirecDistAz = AZi
    End Select

End Function

Private Function AziClosure(Az1, SumA, n)

    AziC = Az1 + SumA - ((n + 2) * 180)
    
    If AziC < 0 Then
        AziClosure = AziC + 360
    Else
        AziClosure = AziC
    End If
    
End Function

Private Function AdjustedAzi(Az1, A, c)

    Az2 = Az1 + (A + c / 3600)
    
    If Az2 - 180 < 0 Then
        Az3 = Az2 - 180 + 360
    ElseIf Az2 - 180 > 360 Then
        Az3 = Az2 - 360 - 180
    Else
        Az3 = Az2 - 180
    End If
    
    AdjustedAzi = Az3

End Function

Private Function LatDep(gdist, Az, LD)

    Lat = gdist * Cos(Az * Pi / 180)
    Dep = gdist * Sin(Az * Pi / 180)
    
    Select Case UCase$(LD)
        Case "L"
            LatDep = Lat
        Case "D"
            LatDep = Dep
    End Select

End Function

Private Function CorrLatDep(ELat, EDep, sumdist, gdist, LD)

    corrlat = (ELat / sumdist) * gdist * -1
    corrdep = (EDep / sumdist) * gdist * -1
    
    Select Case UCase$(LD)
        Case "L"
            CorrLatDep = corrlat
        Case "D"
            CorrLatDep = corrdep
    End Select

End Function

Private Function AdjustedEN(E0, N0, Dep, Lat, corrdep, corrlat, EN)

    E1 = E0 + Dep + corrdep
    N1 = N0 + Lat + corrlat
    
    Select Case UCase$(EN)
    Case "E"
        AdjustedEN = E1
    Case "N"
        AdjustedEN = N1
    End Select

End Function

Private Function CorrAngle(NumAngle, ErAngle)

    CorrAngle = (ErAngle / NumAngle) * 3600 * -1

End Function
'-------------------------------- End Private Function --------------------------------'

'-------------------------------- Summary data to Close Traverse --------------------------------'
Sub SumDataToCloseTRAV()

    Sheets("SUM-DATA").Select

    '1. Information
    LoopName = Range("C3").Value
    Location = Range("C4").Value
    CalDate = Range("C5").Value
    CalBy = Range("C6").Value
    Instrument = Range("C7").Value
    SN = Range("C8").Value

    '2. Number of Station
    numOfSTA = Range("C12").Value + 1

    '3. Fixed Control Points and Scale Factor
    Range("C16").Select

    Dim StaFix(3)
    Dim EFix(3)
    Dim NFix(3)
    For i = 0 To 3

        StaFix(i) = ActiveCell.Offset(i, 0)
        EFix(i) = ActiveCell.Offset(i, 1)
        NFix(i) = ActiveCell.Offset(i, 2)

    Next

    Datum = Range("N16").Value
    a_semi = Range("N17").Value
    df = Range("N18").Value
    LGSF = Range("N19").Value

    '4. Observation Data
    Range("A25").Select

    Dim Sta() As Variant
    Dim dd() As Variant
    Dim mm() As Variant
    Dim ss() As Variant
    Dim MeanDist() As Variant

    ReDim Sta(numOfSTA)
    ReDim dd(numOfSTA)
    ReDim mm(numOfSTA)
    ReDim ss(numOfSTA)
    ReDim MeanDist(numOfSTA)

    For i = 0 To numOfSTA

        Sta(i) = ActiveCell.Offset(i, 1)
        dd(i) = ActiveCell.Offset(i, 2)
        mm(i) = ActiveCell.Offset(i, 3)
        ss(i) = ActiveCell.Offset(i, 4)
        MeanDist(i) = ActiveCell.Offset(i, 7)

    Next
    
    '------------ Close Traverse Report  ------------'
    Sheets("CLOSE TRAVERSE").Select
    
    'Ceate Rows
    For i = 0 To numOfSTA - 2
        Rows("19").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    
    'Information
    Range("E5").Value = LoopName
    Range("E6").Value = CalBy
    Range("P5").Value = Location
    Range("P6").Value = Instrument
    Range("V5").Value = CalDate
    Range("V6").Value = SN

    'Fixed Control Points
    Range("E9").Select
    For w = 0 To 3
        ActiveCell.Offset(w, 0).Value = StaFix(w)
        ActiveCell.Offset(w, 1).Value = EFix(w)
        ActiveCell.Offset(w, 3).Value = NFix(w)
    Next

    'Coordinate Reference System
    Range("M9").Value = Datum
    Range("M10").Value = a_semi
    Range("M11").Value = df
    Range("K13").Value = LGSF
    
    'Number of Station
    Range("R8").Value = numOfSTA - 1
    
    'Traverse Table
    Range("C18").Select

    'First Row Data
    ActiveCell.Offset(0, 0).Value = 1
    ActiveCell.Offset(0, 1).Value = Sta(0)

    'Second Row Data
    ActiveCell.Offset(1, 0).Value = 2
    ActiveCell.Offset(1, 1).Value = Sta(1)
    ActiveCell.Offset(1, 2).Value = dd(1)
    ActiveCell.Offset(1, 3).Value = mm(1)
    ActiveCell.Offset(1, 4).Value = ss(1)

    'Last Row Data
    ActiveCell.Offset(numOfSTA, 0).Value = numOfSTA + 1
    ActiveCell.Offset(numOfSTA, 1).Value = Sta(numOfSTA)

    'Middle Row Data
    For w = 2 To numOfSTA - 1
        ActiveCell.Offset(w, 0).Value = w + 1
        ActiveCell.Offset(w, 1).Value = Sta(w)
        ActiveCell.Offset(w, 2).Value = dd(w)
        ActiveCell.Offset(w, 3).Value = mm(w)
        ActiveCell.Offset(w, 4).Value = ss(w)
        ActiveCell.Offset(w, 12).Value = MeanDist(w - 1)
    Next
    
    MsgBox "GET SUMMARY DATA Completely!"
    
    Sheets("CLOSE TRAVERSE").Select
    Range("E9").Select
    
End Sub


'-------------------------------- Close Traverse Computation --------------------------------'

Sub CloseTraverseComp()
    Sheets("CLOSE TRAVERSE").Select
    
    '------------ Traverse Data ------------'
    '1. Number of Station
    numOfSTA = Range("R8").Value + 1

    '2. Fixed Control Points
    Range("E9").Select
    
    Dim StaFix(3)
    Dim EFix(3)
    Dim NFix(3)
    
    For i = 0 To 3
        StaFix(i) = ActiveCell.Offset(i, 0)
        EFix(i) = ActiveCell.Offset(i, 1)
        NFix(i) = ActiveCell.Offset(i, 3)
    Next
    TravMethod = Range("Z12").Value
    
    '3. Scale Factor
    LGSF = Range("K13").Value

    '4. Observation Data
    Range("C18").Select

    Dim Sta() As Variant
    Dim dd() As Variant
    Dim mm() As Variant
    Dim ss() As Variant
    Dim ObsAng() As Variant
    Dim MeanDist() As Variant
    Dim GridDist() As Variant

    ReDim Sta(numOfSTA)
    ReDim dd(numOfSTA)
    ReDim mm(numOfSTA)
    ReDim ss(numOfSTA)
    ReDim ObsAng(numOfSTA)
    ReDim MeanDist(numOfSTA)
    ReDim GridDist(numOfSTA)

    sumGridDist = 0
    For i = 0 To numOfSTA
        Sta(i) = ActiveCell.Offset(i, 1)
        dd(i) = ActiveCell.Offset(i, 2)
        mm(i) = ActiveCell.Offset(i, 3)
        ss(i) = ActiveCell.Offset(i, 4)
        ObsAng(i) = DMStoDeg(dd(i), mm(i), ss(i))
        MeanDist(i) = ActiveCell.Offset(i, 12)
        GridDist(i) = MeanDist(i) * LGSF
        sumGridDist = sumGridDist + GridDist(i)
    Next

    '------------ Compute Traverse  ------------'
    '1. Sum Observed Angle (deg.)
    sumOBSAng = 0
    For i = 1 To numOfSTA - 1
        sumOBSAng = sumOBSAng + ObsAng(i)
    Next

    '2. Fixed Azimuth of Start (deg.)
    FixedAzStart = DirecDistAz(EFix(0), NFix(0), EFix(1), NFix(1), "A")

    '3. Observed Azimuth of End (deg.)
    'ObsAzEnd = AziClosure(FixedAzStart, sumOBSAng, numOfSTA - 1)
    If AziClosure(FixedAzStart, sumOBSAng, numOfSTA - 1) < 0 Then
        'วัดมุมภายใน บรรจบคู่เดิม
        ObsAzEnd = 360 + AziClosure(FixedAzStart, sumOBSAng, numOfSTA - 1)
    Else
        'วัดมุมภายนอก บรรจบคู่เดิม หรือวัดมุม ไปบรรจบอีกคู่
        ObsAzEnd = AziClosure(FixedAzStart, sumOBSAng, numOfSTA - 1)
    End If

    '4. Fixed Azimuth of End (deg.)
    If TravMethod = True Then
        FixedAzEnd = DirecDistAz(EFix(2), NFix(2), EFix(3), NFix(3), "A")
    Else
        FixedAzEnd = 0
    End If

    '5. Azimuth Closure (deg.)
    If TravMethod = True Then
        AzClosure = ObsAzEnd - FixedAzEnd
    Else
        AzClosure = 0
    End If
    
    '6. Correction Angle (sec.)
    If TravMethod = True Then
        CoorAng = CorrAngle(numOfSTA - 1, AzClosure)
    Else
        CoorAng = 0
    End If

    '7. Azimuth and Adjusted Azimuth
    Dim AdjAng() As Variant
    Dim AdjAz() As Variant
    
    ReDim AdjAng(numOfSTA)
    ReDim AdjAz(numOfSTA)
    
    AdjAng(1) = ObsAng(1) + CoorAng / 3600
    AdjAz(1) = FixedAzStart
    For i = 2 To numOfSTA
        AdjAng(i) = ObsAng(i) + CoorAng / 3600
        AdjAz(i) = AdjustedAzi(AdjAz(i - 1), ObsAng(i - 1), CoorAng)
    Next

    '8. Departure (m.), Latitude (m.), Correction and Accurasy
    Dim Dep() As Variant
    Dim Lat() As Variant

    ReDim Dep(numOfSTA)
    ReDim Lat(numOfSTA)

    sumDep = 0
    sumLat = 0
    For i = 2 To numOfSTA - 1
        Dep(i) = LatDep(GridDist(i), AdjAz(i), "D")
        Lat(i) = LatDep(GridDist(i), AdjAz(i), "L")
        sumDep = sumDep + Dep(i)
        sumLat = sumLat + Lat(i)
    Next

    FixedDep = EFix(2) - EFix(1)
    FixedLat = NFix(2) - NFix(1)
    ErrorDep = sumDep - FixedDep
    ErrorLat = sumLat - FixedLat
    Linear = Sqr(ErrorDep ^ 2 + ErrorLat ^ 2)
    Acc = sumGridDist / Linear

    '9. Coorection of Departure (m.) and Latitude (m.)
    Dim CoorDep() As Variant
    Dim CoorLat() As Variant

    ReDim CoorDep(numOfSTA)
    ReDim CoorLat(numOfSTA)

    For i = 2 To numOfSTA - 1
        CoorDep(i) = CorrLatDep(ErrorLat, ErrorDep, sumGridDist, GridDist(i), "D")
        CoorLat(i) = CorrLatDep(ErrorLat, ErrorDep, sumGridDist, GridDist(i), "L")
    Next

    '10. Adjusted Coordinates
    Dim AdjE() As Variant
    Dim AdjN() As Variant

    ReDim AdjE(numOfSTA)
    ReDim AdjN(numOfSTA)

    AdjE(0) = EFix(0)
    AdjN(0) = NFix(0)
    AdjE(1) = EFix(1)
    AdjN(1) = NFix(1)
    AdjE(numOfSTA) = EFix(3)
    AdjN(numOfSTA) = NFix(3)
    For i = 2 To numOfSTA - 1
        AdjE(i) = AdjustedEN(AdjE(i - 1), AdjN(i - 1), Dep(i), Lat(i), CoorDep(i), CoorLat(i), "E")
        AdjN(i) = AdjustedEN(AdjE(i - 1), AdjN(i - 1), Dep(i), Lat(i), CoorDep(i), CoorLat(i), "N")
    Next

    '------------ Print Traverse Report  ------------'
    Range("R9").Select
    Dim TraArr1() As Variant
    TraArr1 = Array(sumOBSAng, FixedAzStart, ObsAzEnd, _
                    FixedAzEnd, AzClosure, CoorAng / 3600)
    For w = LBound(TraArr1) To UBound(TraArr1)
        ActiveCell.Offset(w, 0).Value = DegtoDMSStr(TraArr1(w))
    Next

    Range("U8").Select
    Dim TraArr2() As Variant
    TraArr2 = Array(sumDep, FixedDep, ErrorDep, Linear, sumGridDist)
    For w = LBound(TraArr2) To UBound(TraArr2)
        ActiveCell.Offset(w, 0).Value = TraArr2(w)
    Next

    Range("W8").Select
    Dim TraArr3() As Variant
    TraArr3 = Array(sumLat, FixedLat, ErrorLat)
    For w = LBound(TraArr3) To UBound(TraArr3)
        ActiveCell.Offset(w, 0).Value = TraArr3(w)
    Next
    Range("W11").Value = Acc

    'Traverse Table
    Range("C18").Select

    'First Row Data
    ActiveCell.Offset(0, 18).Value = AdjE(0)
    ActiveCell.Offset(0, 19).Value = AdjN(0)
    ActiveCell.Offset(0, 20).Value = Sta(0)

    'Second Row Data
    ActiveCell.Offset(1, 5).Value = CoorAng
    ActiveCell.Offset(1, 6).Value = DegtoDMS(AdjAng(1), "D")
    ActiveCell.Offset(1, 7).Value = DegtoDMS(AdjAng(1), "M")
    ActiveCell.Offset(1, 8).Value = DegtoDMS(AdjAng(1), "S")
    ActiveCell.Offset(1, 9).Value = DegtoDMS(AdjAz(1), "D")
    ActiveCell.Offset(1, 10).Value = DegtoDMS(AdjAz(1), "M")
    ActiveCell.Offset(1, 11).Value = DegtoDMS(AdjAz(1), "S")
    ActiveCell.Offset(1, 18).Value = AdjE(1)
    ActiveCell.Offset(1, 19).Value = AdjN(1)
    ActiveCell.Offset(1, 20).Value = Sta(1)

    'Last Row Data
    ActiveCell.Offset(numOfSTA, 9).Value = DegtoDMS(AdjAz(numOfSTA), "D")
    ActiveCell.Offset(numOfSTA, 10).Value = DegtoDMS(AdjAz(numOfSTA), "M")
    ActiveCell.Offset(numOfSTA, 11).Value = DegtoDMS(AdjAz(numOfSTA), "S")
    ActiveCell.Offset(numOfSTA, 18).Value = AdjE(numOfSTA)
    ActiveCell.Offset(numOfSTA, 19).Value = AdjN(numOfSTA)
    ActiveCell.Offset(numOfSTA, 20).Value = Sta(numOfSTA)

    'Middle Row Data
    For w = 2 To numOfSTA - 1
        ActiveCell.Offset(w, 5).Value = CoorAng
        ActiveCell.Offset(w, 6).Value = DegtoDMS(AdjAng(w), "D")
        ActiveCell.Offset(w, 7).Value = DegtoDMS(AdjAng(w), "M")
        ActiveCell.Offset(w, 8).Value = DegtoDMS(AdjAng(w), "S")
        ActiveCell.Offset(w, 9).Value = DegtoDMS(AdjAz(w), "D")
        ActiveCell.Offset(w, 10).Value = DegtoDMS(AdjAz(w), "M")
        ActiveCell.Offset(w, 11).Value = DegtoDMS(AdjAz(w), "S")
        ActiveCell.Offset(w, 12).Value = MeanDist(w)
        ActiveCell.Offset(w, 13).Value = GridDist(w)
        ActiveCell.Offset(w, 14).Value = Dep(w)
        ActiveCell.Offset(w, 15).Value = Lat(w)
        ActiveCell.Offset(w, 16).Value = CoorDep(w)
        ActiveCell.Offset(w, 17).Value = CoorLat(w)
        ActiveCell.Offset(w, 18).Value = AdjE(w)
        ActiveCell.Offset(w, 19).Value = AdjN(w)
        ActiveCell.Offset(w, 20).Value = Sta(w)
    Next

    Sheets("CLOSE TRAVERSE").Select
    Range("E9").Select
    MsgBox "Traverse Computation Complete!"

End Sub

Sub ClearCloseTrav()

    Sheets("CLOSE TRAVERSE").Select

    numOfSTA = Range("R8").Value

    'Information
    Range("E5:E6").Select
    Selection.ClearContents
    Range("P5:P6").Select
    Selection.ClearContents
    Range("V5:V6").Select
    Selection.ClearContents

    'Fixed Control Points
    Range("E9:I12").Select
    Selection.ClearContents

    'Coordinate Reference System
    Range("M9:M11").Select
    Selection.ClearContents
    Range("K13").Select
    Selection.ClearContents

    Range("R9:R14").Select
    Selection.ClearContents

    Range("U8:U12").Select
    Selection.ClearContents
    Range("W8:W11").Select
    Selection.ClearContents

    'Traverse Table
    Range("C18:W18").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    For i = 0 To numOfSTA - 1
        Rows("20").Select
        Selection.Delete Shift:=xlUp
    Next

    Range("R8").Select
    Selection.ClearContents
    
    Sheets("CLOSE TRAVERSE").Select
    Range("E9").Select

End Sub
