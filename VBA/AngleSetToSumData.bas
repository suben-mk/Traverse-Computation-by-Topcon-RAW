Attribute VB_Name = "AngleSetToSumData"
' Topic; Angle Set to Summary Data
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 06/06/2025
'

'Convert D.MMSS to deg min sec
Function DmmsstoDd_mm_ss(d_mmss, DMS)

        dd = Int(d_mmss)
        mm = Int((d_mmss - Int(d_mmss)) * 100)
        ss = (((d_mmss - Int(d_mmss)) * 100) - Int((d_mmss - Int(d_mmss)) * 100)) * 100
        
        Select Case UCase$(DMS)
         Case "D"
                 DmmsstoDd_mm_ss = dd
         Case "M"
                 DmmsstoDd_mm_ss = mm
         Case "S"
                 DmmsstoDd_mm_ss = ss
        End Select

End Function

Sub AngSetToSumData()
    Sheets("SUM-DATA").Select
    OCC = Range("C12").Value
    
    If OCC = 0 Then
        MsgBox "Please input NUMBER OF STATION!"
        Exit Sub
    Else
        For i = 1 To OCC
            Sheets(CStr(i)).Select
            
            Inst = Range("BG44").Value
            HI = Range("BH44").Value
            BSPnt = Range("BG43").Value
            BSHT = Range("BH43").Value
            FSPnt = Range("BG45").Value
            FSHT = Range("BH45").Value
            HorAng = Range("BI44").Value
            ZenithAng = Range("BJ45").Value
            BSHorDist = Range("BK43").Value
            BSSlopeDist = Range("BL43").Value
            FSHorDist = Range("BK45").Value
            FSSlopeDist = Range("BL45").Value
            
            Range("A2").Select
            
            Sheets("SUM-DATA").Select
            'CLOSE TRAVERSE
            Range("A25").Select
            'No.
            ActiveCell.Offset(i - 1, 0).Value = i
            ActiveCell.Offset(i, 0).Value = i + 1
            ActiveCell.Offset(i + 1, 0).Value = i + 2
            
            'Station
            ActiveCell.Offset(i - 1, 1).Value = BSPnt
            ActiveCell.Offset(i, 1).Value = Inst
            ActiveCell.Offset(i + 1, 1).Value = FSPnt
            
            'Hor. Angle
            ActiveCell.Offset(i, 2).Value = DmmsstoDd_mm_ss(HorAng, "D") 'deg
            ActiveCell.Offset(i, 3).Value = DmmsstoDd_mm_ss(HorAng, "M") 'min
            ActiveCell.Offset(i, 4).Value = DmmsstoDd_mm_ss(HorAng, "S") 'sec
            
            'Hor. Dist.
            ActiveCell.Offset(i, 5).Value = BSHorDist
            ActiveCell.Offset(i, 6).Value = FSHorDist
            ActiveCell.Offset(i, 7).Value = "=(R[0]C[-1]+R[1]C[-2])/2"
            ActiveCell.Offset(i, 8).Value = "=R[0]C[-2]-R[1]C[-3]"
            
            'OPEN TRAVERSE-3D
            Range("L25").Select
            'No.
            ActiveCell.Offset(i - 1, 0).Value = i
            ActiveCell.Offset(i, 0).Value = i + 1
            ActiveCell.Offset(i + 1, 0).Value = i + 2
            
            'Station
            ActiveCell.Offset(i - 1, 1).Value = BSPnt
            ActiveCell.Offset(i, 1).Value = Inst
            ActiveCell.Offset(i + 1, 1).Value = FSPnt
            
            'Hor. Angle
            ActiveCell.Offset(i, 2).Value = DmmsstoDd_mm_ss(HorAng, "D") 'deg
            ActiveCell.Offset(i, 3).Value = DmmsstoDd_mm_ss(HorAng, "M") 'min
            ActiveCell.Offset(i, 4).Value = DmmsstoDd_mm_ss(HorAng, "S") 'sec
            
            'Zenith Angle
            ActiveCell.Offset(i, 5).Value = DmmsstoDd_mm_ss(ZenithAng, "D") 'deg
            ActiveCell.Offset(i, 6).Value = DmmsstoDd_mm_ss(ZenithAng, "M") 'min
            ActiveCell.Offset(i, 7).Value = DmmsstoDd_mm_ss(ZenithAng, "S") 'sec
            
            'Hor. Dist.
            ActiveCell.Offset(i, 8).Value = FSHorDist
            
            'Slope Dist.
            ActiveCell.Offset(i, 9).Value = FSSlopeDist
            
            'Height
            ActiveCell.Offset(i, 10).Value = HI
            ActiveCell.Offset(i, 11).Value = FSHT
        Next
    End If
    
    'CLOSE TRAVERSE
    Sheets("SUM-DATA").Select
    Range("A25").Select
    ActiveCell.Offset(OCC, 7).Value = "=R[0]C[-1]"
    ActiveCell.Offset(OCC, 8).Value = ""
    
    MsgBox "SUMMARY DATA was Completed!"
    Sheets("SUM-DATA").Select
    Range("C3").Select
    
End Sub


Sub ClearSumData()
    Sheets("SUM-DATA").Select
    
    'Information
    Range("C3:C8").Select
    Selection.ClearContents

    'Number of Station
    Range("C12").Select
    Selection.ClearContents

    'Fixed Control Points and Scale Factor
    Range("C16:G19").Select
    Selection.ClearContents
    Range("N16").Select
    Selection.ClearContents
    
    'Observation Data
    Range("A25:W25").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Sheets("SUM-DATA").Select
    Range("C3").Select
End Sub
