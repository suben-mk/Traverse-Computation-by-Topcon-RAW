Attribute VB_Name = "TopconToStarNet"
' Topic; Topcon RAW (BFFB) to StarNet Program Rev.15
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 06/06/2025
'

'--------------------------------------- Topcon Total Station ---------------------------------------'
Sub ImportTextFileTS()

    Dim fileToOpen As Variant
    Dim fileFilterPattern As String
    Dim wsMaster As Worksheet
    Dim wbTextImport As Workbook
    
    Application.ScreenUpdating = False
    fileFilterPattern = "Topcon Files (*.txt; *.csv; *.cs1), *.txt; *.csv; *.cs1"
    fileToOpen = Application.GetOpenFilename(fileFilterPattern)
    
    If fileToOpen = False Then
        MsgBox "Please select TOPCON RAW file!"
        Exit Sub
    Else
        Workbooks.OpenText _
                    Filename:=fileToOpen, _
                    StartRow:=2, _
                    DataType:=xlDelimited, _
                    Comma:=True
                    
        Set wbTextImport = ActiveWorkbook
        Set wsMaster = ThisWorkbook.Worksheets("TOPCON-TS RAW")
            wbTextImport.Worksheets(1).Range("A1").CurrentRegion.Copy
            Workbooks(1).Activate
            Range("A3").PasteSpecial Paste:=xlPasteValues
            wbTextImport.Close
     End If

    Application.ScreenUpdating = True
    
    MsgBox "TOPCON-TS RAW was Imported!"
    Sheets("TOPCON-TS RAW").Select
    Range("A3").Select
End Sub

Sub TransformTS_RAW()

    num = ThisWorkbook.Sheets("TOPCON-TS RAW").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count - 1
    
    Sheets("TOPCON-TS RAW").Select
    Range("A2").Select
    
    Dim TsType() As Variant
    Dim Inst() As Variant
    Dim HI() As Variant
    Dim TarPnt() As Variant
    Dim HT() As Variant
    Dim HorAng() As Variant
    Dim HorDist() As Variant
    Dim ZenithAng() As Variant
    Dim SlopeDist() As Variant
    Dim Prism() As Variant
    Dim Code() As Variant
    
    ReDim TsType(num)
    ReDim Inst(num)
    ReDim HI(num)
    ReDim TarPnt(num)
    ReDim HT(num)
    ReDim HorAng(num)
    ReDim HorDist(num)
    ReDim ZenithAng(num)
    ReDim SlopeDist(num)
    ReDim Prism(num)
    ReDim Code(num)
    
    For i = 1 To num
        TsType(i) = ActiveCell.Offset(i, 0)
        Inst(i) = ActiveCell.Offset(i, 1)
        HI(i) = ActiveCell.Offset(i, 2)
        TarPnt(i) = ActiveCell.Offset(i, 3)
        HT(i) = ActiveCell.Offset(i, 4)
        HorAng(i) = ActiveCell.Offset(i, 5)
        HorDist(i) = ActiveCell.Offset(i, 6)
        ZenithAng(i) = ActiveCell.Offset(i, 7)
        SlopeDist(i) = ActiveCell.Offset(i, 8)
        Prism(i) = ActiveCell.Offset(i, 9)
        Code(i) = ActiveCell.Offset(i, 10)
    Next
    
    w = 1
    For i = 1 To num
        Select Case TsType(i)
            Case "BKB"
                    Sheets("CLEAN-TS RAW").Select
                    Range("A2").Select
                    
                    If TsType(i + 1) = "BS" And TsType(i + 2) = "BKB" Then
                        'Do not thing
                    Else
                        Dim BKBArray() As Variant
                        BKBArray = Array(w, TsType(i), Inst(i), HI(i), TarPnt(i), HT(i), HorAng(i), HorDist(i), ZenithAng(i), SlopeDist(i), Prism(i), Code(i))
                        For u = 0 To 10
                            ActiveCell.Offset(w, u).Value = BKBArray(u)
                        Next
                        w = w + 1
                    End If
                    
             Case "BS"
                    Sheets("CLEAN-TS RAW").Select
                    Range("A2").Select
                    
                    If IsEmpty(HorDist(i)) = True Then
                        'Do not thing
                    ElseIf TsType(i - 1) = "BKB" And TsType(i + 1) = "BS" Then
                        'Do not thing
                    ElseIf TsType(i - 1) = "BKB" And TsType(i + 1) = "BKB" Then
                        'Do not thing
                    Else
                        Dim BSArray() As Variant
                        BSArray = Array(w, TsType(i), Inst(i), HI(i), TarPnt(i), HT(i), HorAng(i), HorDist(i), ZenithAng(i), SlopeDist(i), Prism(i), Code(i))
                        For u = 0 To 10
                            ActiveCell.Offset(w, u).Value = BSArray(u)
                        Next
                        w = w + 1
                    End If
                    
             Case "SS"
                    Sheets("CLEAN-TS RAW").Select
                    Range("A2").Select
                    
                    If IsEmpty(HorDist(i)) = True Then
                        'Do not thing
                    Else
                        Dim SSArray() As Variant
                        SSArray = Array(w, TsType(i), Inst(i), HI(i), TarPnt(i), HT(i), HorAng(i), HorDist(i), ZenithAng(i), SlopeDist(i), Prism(i), Code(i))
                        For u = 0 To 10
                            ActiveCell.Offset(w, u).Value = SSArray(u)
                        Next
                        w = w + 1
                    End If
                    
        End Select
    Next
    
    MsgBox "CLEAN-TS RAW was Completed!"
    Sheets("TOPCON-TS RAW").Select
    Range("A3").Select

End Sub

'Convert Degrees to D.MMSS
Function DegtoDmmss(deg)

        dd = Int(deg)
        mm = Int((deg - Int(deg)) * 60)
        ss = (((deg - Int(deg)) * 60) - Int((deg - Int(deg)) * 60)) * 60
        
        DegtoDmmss = dd + mm / 100 + Round(ss, 2) / 10000

End Function

'Convert D.MMSS to Degrees
Function DmmsstoDeg(d_mmss)

        dd = Int(d_mmss)
        mm = Int((d_mmss - Int(d_mmss)) * 100) / 60
        ss = ((((d_mmss - Int(d_mmss)) * 100) - Int((d_mmss - Int(d_mmss)) * 100)) * 100) / 3600
    
        DmmsstoDeg = dd + mm + ss

End Function

'Convert D.MMSS to dd-mm-ss
Function DmmsstoDd_mm_ss(d_mmss)

        dd = Int(d_mmss)
        mm = Int((d_mmss - Int(d_mmss)) * 100)
        ss = (((d_mmss - Int(d_mmss)) * 100) - Int((d_mmss - Int(d_mmss)) * 100)) * 100
    
        DmmsstoDd_mm_ss = Format(dd, "000") & "-" & Format(mm, "00") & "-" & Format(ss, "00.00")

End Function

' Horizontal Angle @1SET
Function LRHorAng(BL, FL, FR, BR, LR)

    HorL = DmmsstoDeg(FL) - DmmsstoDeg(BL)
    HorR = DmmsstoDeg(FR) - DmmsstoDeg(BR)
    
    If HorL < 0 Then
        HorL = HorL + 360
    Else
        HorL = HorL
    End If
    
    If HorR < 0 Then
        HorR = HorR + 360
    Else
        HorR = HorR
    End If
    
    Select Case UCase$(LR)
        Case "L"
            LRHorAng = DegtoDmmss(HorL)
        Case "R"
            LRHorAng = DegtoDmmss(HorR)
    End Select
    
End Function

' Zenith Angle @1SET
Function LRZenithAng(BL, FL, FR, BR, BSFS)
    
    ZAArray1 = Array(BL, FL, FR, BR)
    
    Dim ZAArray2() As Variant
    ReDim ZAArray2(3)
    For i = 0 To 3
        If ZAArray1(i) > 180 Then
            ZAArray2(i) = 360 - DmmsstoDeg(ZAArray1(i))
        Else
            ZAArray2(i) = DmmsstoDeg(ZAArray1(i))
        End If
    Next
    
    Select Case UCase$(BSFS)
        Case "BS"
            LRZenithAng = DegtoDmmss((ZAArray2(0) + ZAArray2(3)) / 2)
        Case "FS"
            LRZenithAng = DegtoDmmss((ZAArray2(1) + ZAArray2(2)) / 2)
    End Select
    
End Function

' Zenith Angle @1SIDE
Function OneSideZenithAng(ZAng)

    If ZAng > 180 Then
        SideZenithAng = 360 - DmmsstoDeg(ZAng)
    Else
        SideZenithAng = DmmsstoDeg(ZAng)
    End If
    
    OneSideZenithAng = DegtoDmmss(SideZenithAng)
    
End Function


Sub TopconTS_RAWToStarNet()
On Error GoTo line2

    num = ThisWorkbook.Sheets("CLEAN-TS RAW").Range("B:B").Cells.SpecialCells(xlCellTypeConstants).Count - 1

    Sheets("CLEAN-TS RAW").Select
    Range("A2").Select
    
    OCC = WorksheetFunction.CountIf(Range("B:B"), "BKB") * 2

    Dim TsType() As Variant
    Dim Inst() As Variant
    Dim HI() As Variant
    Dim TarPnt() As Variant
    Dim HT() As Variant
    Dim HorAng() As Variant
    Dim ZenithAng() As Variant
    Dim SlopeDist() As Variant
    
    ReDim TsType(num)
    ReDim Inst(num)
    ReDim HI(num)
    ReDim TarPnt(num)
    ReDim HT(num)
    ReDim HorAng(num)
    ReDim ZenithAng(num)
    ReDim SlopeDist(num)
    
    For i = 1 To num
        TsType(i) = ActiveCell.Offset(i, 1)
        Inst(i) = ActiveCell.Offset(i, 2)
        HI(i) = ActiveCell.Offset(i, 3)
        TarPnt(i) = ActiveCell.Offset(i, 4)
        HT(i) = ActiveCell.Offset(i, 5)
        HorAng(i) = ActiveCell.Offset(i, 6)
        ZenithAng(i) = ActiveCell.Offset(i, 8)
        SlopeDist(i) = ActiveCell.Offset(i, 9)
    Next


'------------------------------- RAW To StarNet Option 1 -------------------------------'

    Sheets("TRAV STARNET-3D").Select
    Range("A1").Select
    
    Range("A1").Value = "# Topcon TXT to Star*Net by VBA Excel"
    Range("A3").Value = "# Job  : "
    Range("A4").Value = "# Date : "
    Range("A5").Value = "# Time : "
    Range("A6").Value = "# Instrument model : "
    Range("A7").Value = "# Serial number : "
    
    Range("A9").Value = ".Units METERS"
    Range("A10").Value = ".Units DMS"
    Range("A11").Value = ".Order AtFromTo"
    Range("A12").Value = ".Separator -"
    Range("A13").Value = ".Delta Off"
    Range("A14").Value = ".3D"
    Range("A15").Value = "#.SCALE 1.000000000000"
    
    Range("A17").Value = "# Fixed Control Point"
    Range("A18").Value = "#C    ! ! !"
    Range("A19").Value = "#C    ! ! !"
    
    Range("A21").Value = "# Observed Angle and Distance Data"
    
    w = 1
    For i = 1 To num - 1
        Select Case TsType(i)
            Case "BKB"
                    Sheets("TRAV STARNET-3D").Select
                    Range("A21").Select
                    If i = 1 Then
                        ActiveCell.Offset(w, 0).Value = "# OCC:" & Inst(i) & " - " & "BS:" & TarPnt(i + 1) & " - " & "FS:" & TarPnt(i + 3)
                        ActiveCell.Offset(w + 1, 0).Value = "DB"
                        ActiveCell.Offset(w + 1, 1).Value = Inst(i)
                        ActiveCell.Offset(w + 1, 7).Value = "# OCC"
                        w = w + 2
                    ElseIf i > 1 Then
                        ActiveCell.Offset(w, 0).Value = "DE"
                        ActiveCell.Offset(w + 2, 0).Value = "# OCC:" & Inst(i) & " - " & "BS:" & TarPnt(i + 1) & " - " & "FS:" & TarPnt(i + 3)
                        ActiveCell.Offset(w + 3, 0).Value = "DB"
                        ActiveCell.Offset(w + 3, 1).Value = Inst(i)
                        ActiveCell.Offset(w + 3, 7).Value = "# OCC"
                        w = w + 4
                    End If
            Case "BS"
                    Sheets("TRAV STARNET-3D").Select
                    Range("A21").Select
                                                
                        ActiveCell.Offset(w, 0).Value = "DM"
                        ActiveCell.Offset(w, 1).Value = TarPnt(i)
                        ActiveCell.Offset(w, 2).Value = DmmsstoDd_mm_ss(HorAng(i))
                        ActiveCell.Offset(w, 3).Value = SlopeDist(i)
                        ActiveCell.Offset(w, 4).Value = DmmsstoDd_mm_ss(OneSideZenithAng(ZenithAng(i)))
                        ActiveCell.Offset(w, 5).Value = Format(HI(i), "0.0000") & "/" & Format(HT(i), "0.0000")
                        ActiveCell.Offset(w, 7).Value = "# BS"
                    w = w + 1
            Case "SS"
                    Sheets("TRAV STARNET-3D").Select
                    Range("A21").Select
                    
                    ActiveCell.Offset(w, 0).Value = "DM"
                    ActiveCell.Offset(w, 1).Value = TarPnt(i)
                    ActiveCell.Offset(w, 2).Value = DmmsstoDd_mm_ss(HorAng(i))
                    ActiveCell.Offset(w, 3).Value = SlopeDist(i)
                    ActiveCell.Offset(w, 4).Value = DmmsstoDd_mm_ss(OneSideZenithAng(ZenithAng(i)))
                    ActiveCell.Offset(w, 5).Value = Format(HI(i), "0.0000") & "/" & Format(HT(i), "0.0000")
                    ActiveCell.Offset(w, 7).Value = "# FS"
                    w = w + 1
        End Select
    Next
    
    Sheets("TRAV STARNET-3D").Select
    Range("A21").Select
    
    ActiveCell.Offset(w, 0).Value = "DM"
    ActiveCell.Offset(w, 1).Value = TarPnt(num)
    ActiveCell.Offset(w, 2).Value = DmmsstoDd_mm_ss(HorAng(num))
    ActiveCell.Offset(w, 3).Value = SlopeDist(num)
    ActiveCell.Offset(w, 4).Value = DmmsstoDd_mm_ss(OneSideZenithAng(ZenithAng(num)))
    ActiveCell.Offset(w, 5).Value = Format(HI(num), "0.0000") & "/" & Format(HT(num), "0.0000")
    ActiveCell.Offset(w, 7).Value = "# BS"
    ActiveCell.Offset(w + 1, 0).Value = "DE"

'------------------------------- RAW To StarNet Option 2 -------------------------------'

    Sheets("TRAV STARNET-3D").Select
    Range("M1").Select
    
    Range("M1").Value = "# Topcon TXT to Star*Net by VBA Excel"
    Range("M3").Value = "# Job  : "
    Range("M4").Value = "# Date : "
    Range("M5").Value = "# Time : "
    Range("M6").Value = "# Instrument model : "
    Range("M7").Value = "# Serial number : "
    
    Range("M9").Value = ".Units METERS"
    Range("M10").Value = ".Units DMS"
    Range("M11").Value = ".Order AtFromTo"
    Range("M12").Value = ".Separator -"
    Range("M13").Value = ".Delta Off"
    Range("M14").Value = ".3D"
    Range("M15").Value = "#.SCALE 1.000000000000"
    
    Range("M17").Value = "# Fixed Control Point"
    Range("M18").Value = "#C    ! ! !"
    Range("M19").Value = "#C    ! ! !"
    
    Range("M21").Value = "# Observed Angle and Distance Data"
    
    d = 1
    w = 1
    For i = 1 To OCC
        Select Case TsType(d)
            Case "BKB"
                    Sheets("TRAV STARNET-3D").Select
                    Range("M21").Select
                    If d = 1 Then
                        ActiveCell.Offset(w, 0).Value = "# OCC:" & Inst(d) & " - " & "BS:" & TarPnt(d + 1) & " - " & "FS:" & TarPnt(d + 3)
                        d = d + 1
                        w = w + 1
                    ElseIf d > 1 Then
                        ActiveCell.Offset(w + 1, 0).Value = "# OCC:" & Inst(d) & " - " & "BS:" & TarPnt(d + 1) & " - " & "FS:" & TarPnt(d + 3)
                        d = d + 1
                        w = w + 2
                    End If
            Case "BS"
                    SetNo = 1
                    Do Until TsType(d) = "BKB"
                        Sheets("TRAV STARNET-3D").Select
                        Range("M21").Select
                                                
                        Dim BSSB_TsType(0 To 3) As Variant
                        Dim BSSB_Inst(0 To 3) As Variant
                        Dim BSSB_HI(0 To 3) As Variant
                        Dim BSSB_TarPnt(0 To 3) As Variant
                        Dim BSSB_HT(0 To 3) As Variant
                        Dim BSSB_HorAng(0 To 3) As Variant
                        Dim BSSB_ZenithAng(0 To 3) As Variant
                        Dim BSSB_SlopeDist(0 To 3) As Variant
                        
                        For u = 0 To 3
                            BSSB_TsType(u) = TsType(d + u)
                            BSSB_Inst(u) = Inst(d + u)
                            BSSB_HI(u) = HI(d + u)
                            BSSB_TarPnt(u) = TarPnt(d + u)
                            BSSB_HT(u) = HT(d + u)
                            BSSB_HorAng(u) = HorAng(d + u)
                            BSSB_ZenithAng(u) = ZenithAng(d + u)
                            BSSB_SlopeDist(u) = SlopeDist(d + u)
                        Next
                        
                        At_Int = BSSB_Inst(0)
                        From_BS = BSSB_TarPnt(0)
                        To_BS = BSSB_TarPnt(1)
                        
                        HorL = LRHorAng(BSSB_HorAng(0), BSSB_HorAng(1), BSSB_HorAng(2), BSSB_HorAng(3), "L")
                        HorR = LRHorAng(BSSB_HorAng(0), BSSB_HorAng(1), BSSB_HorAng(2), BSSB_HorAng(3), "R")
                        HorLR = DegtoDmmss((DmmsstoDeg(HorL) + DmmsstoDeg(HorR)) / 2)
                        
                        SlopeDistBS = (BSSB_SlopeDist(0) + BSSB_SlopeDist(3)) / 2
                        SlopeDistSS = (BSSB_SlopeDist(1) + BSSB_SlopeDist(2)) / 2
                        
                        ZAngBS = LRZenithAng(BSSB_ZenithAng(0), BSSB_ZenithAng(1), BSSB_ZenithAng(2), BSSB_ZenithAng(3), "BS")
                        ZAngFS = LRZenithAng(BSSB_ZenithAng(0), BSSB_ZenithAng(1), BSSB_ZenithAng(2), BSSB_ZenithAng(3), "FS")
                        
                        HI_Int = (BSSB_HI(0) + BSSB_HI(1) + BSSB_HI(2) + BSSB_HI(3)) / 4
                        HT_BS = (BSSB_HT(0) + BSSB_HT(3)) / 2
                        HT_FS = (BSSB_HT(1) + BSSB_HT(2)) / 2
                        
                        ActiveCell.Offset(w, 0).Value = "M"
                        ActiveCell.Offset(w, 1).Value = At_Int & "-" & From_BS & "-" & To_BS
                        ActiveCell.Offset(w, 2).Value = DmmsstoDd_mm_ss(HorLR)
                        ActiveCell.Offset(w, 3).Value = SlopeDistSS
                        ActiveCell.Offset(w, 4).Value = DmmsstoDd_mm_ss(ZAngFS)
                        ActiveCell.Offset(w, 5).Value = Format(HI_Int, "0.0000") & "/" & Format(HT_FS, "0.0000")
                        ActiveCell.Offset(w, 7).Value = "#" & " Set " & SetNo
            
                        d = d + 4
                        w = w + 1
                        SetNo = SetNo + 1
                    Loop
        End Select
    Next

line2:          'error-handler #2
    MsgBox "TRAV STARNET-3D was Completed!"
    Sheets("CLEAN-TS RAW").Select
    Range("A3").Select
End Sub

Sub ClearAll_TS()

    Sheets("TOPCON-TS RAW").Select
    Range("A3:K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("CLEAN-TS RAW").Select
    Range("A3:M3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("TRAV STARNET-3D").Select
    Range("A:H").Select
    Selection.ClearContents
    
    Sheets("TRAV STARNET-3D").Select
    Range("M:T").Select
    Selection.ClearContents
    
    Sheets("TOPCON-TS RAW").Select
    Range("A3").Select
End Sub
