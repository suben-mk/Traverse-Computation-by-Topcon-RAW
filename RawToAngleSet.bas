Attribute VB_Name = "RawToAngleSet"
' Topic; Cleaned Raw Data to Angle Set
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 06/06/2025
'

Sub CleanedRawToAngSet()
    
    Sheets("CLEAN-TS RAW").Select
    Range("A2").Select
    
    Sheets("ANG-SET").Visible = True
    
    OCC = WorksheetFunction.CountIf(Range("B:B"), "BKB")

    Dim TsType() As Variant
    Dim FwBw() As Variant
    Dim numOfSet() As Variant

    ReDim TsType(OCC)
    ReDim FwBw(OCC)
    ReDim numOfSet(OCC)
    
    SheetNo = 1
    w = 0
    For i = 1 To OCC
        TsType(i) = ActiveCell.Offset(i + w, 1)
        FwBw(i) = ActiveCell.Offset(i + w, 11)
        numOfSet(i) = ActiveCell.Offset(i + w, 12)
        
        If TsType(i) = "BKB" And FwBw(i) = "FW" Then
            Sheets("ANG-SET").Copy After:=Sheets(Sheets.Count)
            Sheets("ANG-SET (2)").name = SheetNo
            
            Sheets("CLEAN-TS RAW").Select
            Range("A2").Select
            Range(ActiveCell.Offset((i + w) + 1, 1), ActiveCell.Offset((i + w) + (numOfSet(i) * 4), 9)).Select
            Selection.Copy
            
            Sheets(CStr(SheetNo)).Select
            Range("BP3").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Range("A2").Select
            
            Sheets("CLEAN-TS RAW").Select
            Application.CutCopyMode = False
            Range("A2").Select
            
        ElseIf TsType(i) = "BKB" And FwBw(i) = "BW" Then
            SheetNo = SheetNo - 1
            
            Sheets("CLEAN-TS RAW").Select
            Range("A2").Select
            Range(ActiveCell.Offset((i + w) + 1, 1), ActiveCell.Offset((i + w) + (numOfSet(i) * 4), 9)).Select
            Selection.Copy
            
            Sheets(CStr(SheetNo)).Select
            Range("CB3").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Range("A2").Select
            
            Sheets("CLEAN-TS RAW").Select
            Application.CutCopyMode = False
            Range("A2").Select
            
        End If
        
        SheetNo = SheetNo + 1
        w = w + (numOfSet(i) * 4)
    Next
    
    Sheets("ANG-SET").Visible = False
    
    MsgBox "ANGLE SET COLECTION was Completed!"
    Sheets("CLEAN-TS RAW").Select
    Range("A3").Select
End Sub
