Attribute VB_Name = "OpenTraverse"
' Topic; Open Traverse Program
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

Private Function ObsAzimuth(Az1, A)

    Az2 = Az1 + A
    
    If Az2 - 180 < 0 Then
        Az3 = Az2 - 180 + 360
    ElseIf Az2 - 180 > 360 Then
        Az3 = Az2 - 360 - 180
    Else
        Az3 = Az2 - 180
    End If
    
    ObsAzimuth = Az3

End Function

'2 Points 3D Coordinates
Private Function Coor3Dby2Points(Estn, Nstn, Zstn, HI, Ei, Ni, Zi, HP, HDSDZAAZ)
    
    ' Hor.Dist, Slop Dist, Zenith Angle, STN --> Next Point
    HDist = Sqr((Ei - Estn) ^ 2 + (Ni - Nstn) ^ 2)
    DiffEL = (Zi + HP) - (Zstn + HI)
    SDist = Sqr(DiffEL ^ 2 + HDist ^ 2)
    Qv = RadtoDeg(Atn(DiffEL / HDist))
    Za = 90 - Qv
    
    ' Azimuth, STN --> Next Point
     AZi = DirecDistAz(Estn, Nstn, Ei, Ni, "A")
      
    Select Case UCase$(HDSDZAAZ)
        Case "HD"
            Coor3Dby2Points = HDist
        Case "SD"
            Coor3Dby2Points = SDist
        Case "ZA"
            Coor3Dby2Points = Za
        Case "AZ"
            Coor3Dby2Points = AZi
    End Select
    
End Function

'Compute 3D coordinate
Private Function compute3DNEZ(Estn, Nstn, Zstn, AZi, HD, Zenth, HI, HP, NEZ)
    'คำนวนค่าพิกัด N,E ที่จุดใดๆ
    Ni = Nstn + HD * Cos(DegtoRad(AZi))
    Ei = Estn + HD * Sin(DegtoRad(AZi))

    'คำนวนค่าระดับที่จุดใดๆ
    ang = 90 - Zenth
    VD = HD * Tan(DegtoRad(ang))
    Zi = (Zstn + HI + VD) - HP
    
    Select Case UCase$(NEZ)
        Case "N"
        compute3DNEZ = Ni
        Case "E"
        compute3DNEZ = Ei
        Case "Z"
        compute3DNEZ = Zi
    End Select
    
End Function
'-------------------------------- End Private Function --------------------------------'

'-------------------------------- Summary data to Open Traverse --------------------------------'
Sub SumDataToOpenTRAV()

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
    Dim ZFix(3)
    Dim HIFix(3)
    For i = 0 To 3
        StaFix(i) = ActiveCell.Offset(i, 0)
        EFix(i) = ActiveCell.Offset(i, 1)
        NFix(i) = ActiveCell.Offset(i, 2)
        ZFix(i) = ActiveCell.Offset(i, 3)
        HIFix(i) = ActiveCell.Offset(i, 4)
    Next

    Datum = Range("N16").Value
    a_semi = Range("N17").Value
    df = Range("N18").Value
    LGSF = Range("N19").Value

    '4. Observation Data
    Range("L25").Select

    Dim Sta() As Variant
    Dim HAdd() As Variant
    Dim HAmm() As Variant
    Dim HAss() As Variant
    Dim ZAdd() As Variant
    Dim ZAmm() As Variant
    Dim ZAss() As Variant
    Dim HorDist() As Variant
    Dim HI() As Variant
    Dim HP() As Variant

    ReDim Sta(numOfSTA)
    ReDim HAdd(numOfSTA)
    ReDim HAmm(numOfSTA)
    ReDim HAss(numOfSTA)
    ReDim ZAdd(numOfSTA)
    ReDim ZAmm(numOfSTA)
    ReDim ZAss(numOfSTA)
    ReDim HorDist(numOfSTA)
    ReDim HI(numOfSTA)
    ReDim HP(numOfSTA)

    For i = 0 To numOfSTA
        Sta(i) = ActiveCell.Offset(i, 1)
        HAdd(i) = ActiveCell.Offset(i, 2)
        HAmm(i) = ActiveCell.Offset(i, 3)
        HAss(i) = ActiveCell.Offset(i, 4)
        ZAdd(i) = ActiveCell.Offset(i, 5)
        ZAmm(i) = ActiveCell.Offset(i, 6)
        ZAss(i) = ActiveCell.Offset(i, 7)
        HorDist(i) = ActiveCell.Offset(i, 8)
        HI(i) = ActiveCell.Offset(i, 10)
        HP(i) = ActiveCell.Offset(i, 11)
    Next

    '------------ Open Traverse Report  ------------'
    Sheets("OPEN TRAV-3D").Select

    'Ceate Rows
    For i = 0 To numOfSTA - 2
        Rows("18").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next

    'Information
    Range("E5").Value = LoopName
    Range("E6").Value = CalBy
    Range("O5").Value = Location
    Range("O6").Value = Instrument
    Range("T5").Value = CalDate
    Range("T6").Value = SN

    'Fixed Control Points
    Range("F9").Select
    For w = 0 To 1
        ActiveCell.Offset(w, 0).Value = StaFix(w)
        ActiveCell.Offset(w, 1).Value = EFix(w)
        ActiveCell.Offset(w, 3).Value = NFix(w)
        ActiveCell.Offset(w, 5).Value = ZFix(w)
        ActiveCell.Offset(w, 7).Value = HIFix(w)
    Next

    'Coordinate Reference System
    Range("Q9").Value = Datum
    Range("Q10").Value = a_semi
    Range("Q11").Value = df
    Range("O13").Value = LGSF

    'Number of Station
    Range("U8").Value = numOfSTA - 1

    'Traverse Table
    Range("C17").Select

    'First Row Data
    ActiveCell.Offset(0, 0).Value = 1
    ActiveCell.Offset(0, 1).Value = Sta(0)

    'Second Row Data
    ActiveCell.Offset(1, 0).Value = 2
    ActiveCell.Offset(1, 1).Value = Sta(1)
    ActiveCell.Offset(1, 2).Value = HAdd(1)
    ActiveCell.Offset(1, 3).Value = HAmm(1)
    ActiveCell.Offset(1, 4).Value = HAss(1)
    ActiveCell.Offset(1, 13).Value = HI(1)

    'Middle Row Data
    For w = 2 To numOfSTA
        ActiveCell.Offset(w, 0).Value = w + 1
        ActiveCell.Offset(w, 1).Value = Sta(w)
        ActiveCell.Offset(w, 2).Value = HAdd(w)
        ActiveCell.Offset(w, 3).Value = HAmm(w)
        ActiveCell.Offset(w, 4).Value = HAss(w)
        ActiveCell.Offset(w, 5).Value = ZAdd(w - 1)
        ActiveCell.Offset(w, 6).Value = ZAmm(w - 1)
        ActiveCell.Offset(w, 7).Value = ZAss(w - 1)
        ActiveCell.Offset(w, 11).Value = HorDist(w - 1)
        ActiveCell.Offset(w, 13).Value = HI(w)
        ActiveCell.Offset(w, 14).Value = HP(w - 1)
    Next

    MsgBox "GET SUMMARY DATA Completely!"

    Sheets("OPEN TRAV-3D").Select
    Range("F9").Select
    
End Sub

'-------------------------------- Open Traverse Computation --------------------------------'

Sub OpenTraverseComp()
    Sheets("OPEN TRAV-3D").Select
    
    '------------ Traverse Data ------------'
    '1. Number of Station
    numOfSTA = Range("U8").Value + 1

    '2. Fixed Control Points
    Range("F9").Select

    Dim StaFix(1)
    Dim EFix(1)
    Dim NFix(1)
    Dim ZFix(1)
    Dim HIFix(1)
    
    For i = 0 To 1
        StaFix(i) = ActiveCell.Offset(i, 0)
        EFix(i) = ActiveCell.Offset(i, 1)
        NFix(i) = ActiveCell.Offset(i, 3)
        ZFix(i) = ActiveCell.Offset(i, 5)
        HIFix(i) = ActiveCell.Offset(i, 7)
    Next
    
    '3. Scale Factor
    LGSF = Range("O13").Value

    '4. Observation Data
    Range("C17").Select

    Dim Sta() As Variant
    Dim HAdd() As Variant
    Dim HAmm() As Variant
    Dim HAss() As Variant
    Dim ObsHAng() As Variant
    Dim ZAdd() As Variant
    Dim ZAmm() As Variant
    Dim ZAss() As Variant
    Dim ObsZAng() As Variant
    Dim HorDist() As Variant
    Dim GridDist() As Variant
    Dim HI() As Variant
    Dim HP() As Variant

    ReDim Sta(numOfSTA)
    ReDim HAdd(numOfSTA)
    ReDim HAmm(numOfSTA)
    ReDim HAss(numOfSTA)
    ReDim ObsHAng(numOfSTA)
    ReDim ZAdd(numOfSTA)
    ReDim ZAmm(numOfSTA)
    ReDim ZAss(numOfSTA)
    ReDim ObsZAng(numOfSTA)
    ReDim HorDist(numOfSTA)
    ReDim GridDist(numOfSTA)
    ReDim HI(numOfSTA)
    ReDim HP(numOfSTA)

    sumGridDist = 0
    For i = 0 To numOfSTA
        Sta(i) = ActiveCell.Offset(i, 1)
        HAdd(i) = ActiveCell.Offset(i, 2)
        HAmm(i) = ActiveCell.Offset(i, 3)
        HAss(i) = ActiveCell.Offset(i, 4)
        ObsHAng(i) = DMStoDeg(HAdd(i), HAmm(i), HAss(i))
        ZAdd(i) = ActiveCell.Offset(i, 5)
        ZAmm(i) = ActiveCell.Offset(i, 6)
        ZAss(i) = ActiveCell.Offset(i, 7)
        ObsZAng(i) = DMStoDeg(ZAdd(i), ZAmm(i), ZAss(i))
        HorDist(i) = ActiveCell.Offset(i, 11)
        GridDist(i) = HorDist(i) * LGSF
        HI(i) = ActiveCell.Offset(i, 13)
        HP(i) = ActiveCell.Offset(i, 14)
        
        sumGridDist = sumGridDist + GridDist(i)
    Next

    '------------ Compute Traverse  ------------'
    '1. Fixed Azimuth of Start (deg.)
    FixedAzStart = DirecDistAz(EFix(0), NFix(0), EFix(1), NFix(1), "A")
    
    '2. Zenith Angle of Backsight (deg.)
    ZAofBS = Coor3Dby2Points(EFix(1), NFix(1), ZFix(1), HIFix(1), EFix(0), NFix(0), ZFix(0), HIFix(0), "ZA")
    
    '3. Horzontal Dist. of Backsight (m.)
    HorDistOfBS = Coor3Dby2Points(EFix(1), NFix(1), ZFix(1), HIFix(1), EFix(0), NFix(0), ZFix(0), HIFix(0), "HD")
    
    '4. Slope Dist. of Backsight (deg.)
    SlopeDistOfBS = Coor3Dby2Points(EFix(1), NFix(1), ZFix(1), HIFix(1), EFix(0), NFix(0), ZFix(0), HIFix(0), "SD")
   
    '5. Azimuth
    Dim ObsAz() As Variant
    ReDim ObsAz(numOfSTA)

    ObsAz(1) = FixedAzStart
    For i = 2 To numOfSTA
        ObsAz(i) = ObsAzimuth(ObsAz(i - 1), ObsHAng(i - 1))
    Next

    '6. 3D Coordinates
    Dim CoorE() As Variant
    Dim CoorN() As Variant
    Dim CoorZ() As Variant

    ReDim CoorE(numOfSTA)
    ReDim CoorN(numOfSTA)
    ReDim CoorZ(numOfSTA)

    CoorE(0) = EFix(0)
    CoorN(0) = NFix(0)
    CoorZ(0) = ZFix(0)
    CoorE(1) = EFix(1)
    CoorN(1) = NFix(1)
    CoorZ(1) = ZFix(1)
    For i = 2 To numOfSTA
        CoorE(i) = compute3DNEZ(CoorE(i - 1), CoorN(i - 1), CoorZ(i - 1), ObsAz(i), GridDist(i), ObsZAng(i), HI(i - 1), HP(i), "E")
        CoorN(i) = compute3DNEZ(CoorE(i - 1), CoorN(i - 1), CoorZ(i - 1), ObsAz(i), GridDist(i), ObsZAng(i), HI(i - 1), HP(i), "N")
        CoorZ(i) = compute3DNEZ(CoorE(i - 1), CoorN(i - 1), CoorZ(i - 1), ObsAz(i), GridDist(i), ObsZAng(i), HI(i - 1), HP(i), "Z")
    Next

    '------------ Print Traverse Report  ------------'
    Range("U9").Select
    Dim TraArr1() As Variant
    TraArr1 = Array(DegtoDMSStr(FixedAzStart), DegtoDMSStr(ZAofBS), HorDistOfBS, SlopeDistOfBS, sumGridDist)
    For w = LBound(TraArr1) To UBound(TraArr1)
        ActiveCell.Offset(w, 0).Value = TraArr1(w)
    Next

    'Traverse Table
    Range("C17").Select

    'First Row Data
    ActiveCell.Offset(0, 15).Value = CoorE(0)
    ActiveCell.Offset(0, 16).Value = CoorN(0)
    ActiveCell.Offset(0, 17).Value = CoorZ(0)
    ActiveCell.Offset(0, 18).Value = Sta(0)

    'Second Row Data
    ActiveCell.Offset(1, 8).Value = DegtoDMS(ObsAz(1), "D")
    ActiveCell.Offset(1, 9).Value = DegtoDMS(ObsAz(1), "M")
    ActiveCell.Offset(1, 10).Value = DegtoDMS(ObsAz(1), "S")
    ActiveCell.Offset(1, 15).Value = CoorE(1)
    ActiveCell.Offset(1, 16).Value = CoorN(1)
    ActiveCell.Offset(1, 17).Value = CoorZ(1)
    ActiveCell.Offset(1, 18).Value = Sta(1)

    'Middle Row Data
    For w = 2 To numOfSTA
        ActiveCell.Offset(w, 8).Value = DegtoDMS(ObsAz(w), "D")
        ActiveCell.Offset(w, 9).Value = DegtoDMS(ObsAz(w), "M")
        ActiveCell.Offset(w, 10).Value = DegtoDMS(ObsAz(w), "S")
        ActiveCell.Offset(w, 12).Value = GridDist(w)
        ActiveCell.Offset(w, 15).Value = CoorE(w)
        ActiveCell.Offset(w, 16).Value = CoorN(w)
        ActiveCell.Offset(w, 17).Value = CoorZ(w)
        ActiveCell.Offset(w, 18).Value = Sta(w)
    Next

    Sheets("OPEN TRAV-3D").Select
    Range("F9").Select
    MsgBox "Traverse Computation Complete!"

End Sub

Sub ClearOpenTrav()

    Sheets("OPEN TRAV-3D").Select

    numOfSTA = Range("U8").Value

    'Information
    Range("E5:E6").Select
    Selection.ClearContents
    Range("O5:O6").Select
    Selection.ClearContents
    Range("T5:T6").Select
    Selection.ClearContents

    'Fixed Control Points
    Range("F9:N10").Select
    Selection.ClearContents

    'Coordinate Reference System
    Range("Q9:Q11").Select
    Selection.ClearContents
    Range("O13").Select
    Selection.ClearContents

    Range("U9:U13").Select
    Selection.ClearContents

    'Traverse Table
    Range("C17:U17").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    For i = 0 To numOfSTA - 1
        Rows("19").Select
        Selection.Delete Shift:=xlUp
    Next

    Range("U8").Select
    Selection.ClearContents
    
    Sheets("OPEN TRAV-3D").Select
    Range("F9").Select

End Sub
