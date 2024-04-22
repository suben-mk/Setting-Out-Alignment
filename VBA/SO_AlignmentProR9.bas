Attribute VB_Name = "SO_AlignmentProR9"
' Topic; Setting Out Alignment Program
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 19/10/2023
'
 Option Base 1
 Const Pi As Single = 3.141592654
 
'----------------General Private Function----------------'

'Convert Degrees to Radian.
Private Function DegtoRad(d)

    DegtoRad = d * (Pi / 180)

End Function

'Convert Radian to Degrees.
Private Function RadtoDeg(r)

    RadtoDeg = r * (180 / Pi)

End Function

'Compute Northing and Easting by Local Coordinate (Y, X) , Coordinate of Center and Azimuth.
Private Function CoorYXtoNE(ECL, NCL, AZCL, Y, x, EN)

    Ei = ECL + Y * Sin(DegtoRad(AZCL)) + x * Sin(DegtoRad(90 + AZCL))
    Ni = NCL + Y * Cos(DegtoRad(AZCL)) + x * Cos(DegtoRad(90 + AZCL))
    
    Select Case UCase$(EN)
     Case "E"
             CoorYXtoNE = Ei
     Case "N"
             CoorYXtoNE = Ni
  End Select
End Function 'Coordinate Y,X to N, E

'Compute Distance and Azimuth from 2 Points.
Private Function DirecDistAz(EStart, NStart, EEnd, NEnd, DA)

    De = EEnd - EStart: dN = NEnd - NStart
    Distance = Sqr(De ^ 2 + dN ^ 2)
    
    If dN <> 0 Then Q = RadtoDeg(Atn(De / dN))
      If dN = 0 Then
        If De > 0 Then
          Azi = 90
        ElseIf De < 0 Then
          Azi = 270
        Else
          Azi = False
      End If
      
    ElseIf dN > 0 Then
      If De > 0 Then
          Azi = Q
      ElseIf De < 0 Then
          Azi = 360 + Q
      End If
      
    ElseIf dN < 0 Then
          Azi = 180 + Q
    End If
    
    Select Case UCase$(DA)
      Case "D"
          DirecDistAz = Distance
      Case "A"
          DirecDistAz = Azi
    End Select

End Function 'DirecDistAz

'Check Azimuth
Private Function ModAzi(AziA)
      Dim k As Integer
      k = Int(AziA / 360)
      If AziA >= 360 * k Then
         ModAzi = AziA - (360 * k)
      ElseIf AziA < 0 Then
         ModAzi = (360 * k) + AziA
      End If
End Function 'ModAZi

'Convert Degrees, Minuses, Second to D° M' S"
Private Function DegtoDMSStr(deg)

        DD = Int(deg)
        mm = Int((deg - Int(deg)) * 60)
        SS = (((deg - Int(deg)) * 60) - Int((deg - Int(deg)) * 60)) * 60
        
        DegtoDMSStr = " " & DD & ChrW(&HB0) & " " & mm & "' " & Round(SS, 2) & """"

End Function 'DegtoDMSStr
'----------------End General Private Function----------------'

'-----------------Private Function of Spiral Curve-----------------'

Private Function Spiral(Ls, Rc, d, QKPXYT)
'1.Spiral Angle and Circular Angle (rad.)
    Qs = Ls / (2 * Rc)

'2.Offset Xs, Ys
    C1 = 1 / 3: C2 = -1 / 10: C3 = -1 / 42: C4 = 1 / 216: C5 = 1 / 1320: C6 = -1 / 9360: C7 = -1 / 75600: C8 = 1 / 685440
    Xs = Ls * (1 + (C2 * Qs ^ 2) + (C4 * Qs ^ 4) + (C6 * Qs ^ 6) + (C8 * Qs ^ 8))
    Ys = Ls * ((C1 * Qs) + (C3 * Qs ^ 3) + (C5 * Qs ^ 5) + (C7 * Qs ^ 7))

'3.Offset from PCO Tangent to New Curve (m.)
    P = Ys - Rc * (1 - Cos(Qs))

'4.Distance from PCO Tangent to New Curve (m.)
    k = Xs - Rc * Sin(Qs)

'5.Tangent Line from TS to PI or ST to PI (m.)
    Ts = (Rc + P) * Tan(DegtoRad(d / 2)) + k

    Select Case UCase$(QKPXYT)
      Case "Q"
          Spiral = Qs
      Case "K"
          Spiral = k
      Case "P"
          Spiral = P
      Case "X"
          Spiral = Xs
      Case "Y"
          Spiral = Ys
      Case "T"
          Spiral = Ts
    End Select 'PXYT
End Function 'Spiral parameter

Private Function SpiralIN(Ls1, Ls2, Rc, d, XYT)

    If Ls1 = Ls2 Then
    
        Xs1 = Spiral(Ls1, Rc, d, "X")
        Ys1 = Spiral(Ls1, Rc, d, "Y")
        Ts1 = Spiral(Ls1, Rc, d, "T")
    Else
        P1 = Spiral(Ls1, Rc, d, "P")
        K1 = Spiral(Ls1, Rc, d, "K")
        P2 = Spiral(Ls2, Rc, d, "P")
        
        Xs1 = Spiral(Ls1, Rc, d, "X")
        Ys1 = Spiral(Ls1, Rc, d, "Y")
        Ts1 = K1 + Rc * Tan(DegtoRad(d / 2)) + P2 / Sin(DegtoRad(d)) - P1 / Tan(DegtoRad(d))
    End If
    
    Select Case UCase$(XYT)
      Case "X"
          SpiralIN = Xs1
      Case "Y"
          SpiralIN = Ys1
      Case "T"
          SpiralIN = Ts1
    End Select 'XYT
End Function

Function SpiralOUT(Ls1, Ls2, Rc, d, XYT)

    If Ls1 = Ls2 Then
    
        Xs2 = Spiral(Ls2, Rc, d, "X")
        Ys2 = Spiral(Ls2, Rc, d, "Y")
        Ts2 = Spiral(Ls2, Rc, d, "T")
    Else
        P1 = Spiral(Ls1, Rc, d, "P")
        P2 = Spiral(Ls2, Rc, d, "P")
        K2 = Spiral(Ls2, Rc, d, "K")
        
        Xs2 = Spiral(Ls2, Rc, d, "X")
        Ys2 = Spiral(Ls2, Rc, d, "Y")
        Ts2 = K2 + Rc * Tan(DegtoRad(d / 2)) + P1 / Sin(DegtoRad(d)) - P2 / Tan(DegtoRad(d))
    End If
    
    Select Case UCase$(XYT)
      Case "X"
          SpiralOUT = Xs2
      Case "Y"
          SpiralOUT = Ys2
      Case "T"
          SpiralOUT = Ts2
    End Select 'XYT
End Function
'-----------------End Private Function of Spiral Curve-----------------'

'-----------------Setting out Horizontal Alignment Computation-----------------'

Sub HorAlignmentComp()
    
    Dim lastRow As Long
    lastRow = ThisWorkbook.Sheets("HIP DATA").Cells(Rows.Count, 1).End(xlUp).Row
    totalHip = lastRow - 3 'Total HIP curve
    MsgBox "TOTAL HIP =" & " " & totalHip
    
    Alignment_Name = Range("B1") 'Alignment name
    HorSOName = "HOR-SETTING OUT" 'Sheet name for setting out
    HorArrName = "HOR-ARRAY" 'Sheet name for array
    
    Sheets.Add(After:=Sheets("HIP DATA")).Name = HorSOName
    Sheets.Add(After:=Sheets(HorSOName)).Name = HorArrName

'--------------------------Format Table (Setting Out) -----------------------'
    
    Sheets(HorSOName).Select
    
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.RowHeight = 20
    Columns("B:B").Select
    Selection.ColumnWidth = 15
    Columns("C:C").Select
    Selection.ColumnWidth = 8
    Columns("D:D").Select
    Selection.ColumnWidth = 11
    Columns("E:G").Select
    Selection.ColumnWidth = 15
    Columns("H:H").Select
    Selection.ColumnWidth = 6
    Columns("I:I").Select
    Selection.ColumnWidth = 15
    Columns("J:J").Select
    Selection.ColumnWidth = 6
    Columns("K:K").Select
    Selection.ColumnWidth = 15
    Columns("L:L").Select
    Selection.ColumnWidth = 6
    Columns("M:M").Select
    Selection.ColumnWidth = 4.11
    Columns("N:Q").Select
    Selection.ColumnWidth = 9
    Columns("R:R").Select
    Selection.ColumnWidth = 20
    Range("C4:H4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("I4:Q4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("G5:H5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("I5:J5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("K5:M5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("N5:Q5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B3:R3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C2:E2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Rows("3:3").Select
    Selection.RowHeight = 30
    Range("B3:R3").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
 
'--------------------------Head Table (Setting Out)-----------------------'

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "ALIGNMENT NAME :"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = Alignment_Name
    
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "SETTING OUT DATA - HORIZONTAL ALIGNMENT"
    
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "HORIZONTAL MAIN POINTS"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "HORIZONTAL ELEMENTS"
    
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "HIP NO."
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "POINT"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "CHAINAGE"
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "EASTING"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "NORTHING"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "AZIMUTH"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "AZIMUTH TANGENT"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "DEFLECTION ANGLE"
    Range("N5").Select
    ActiveCell.FormulaR1C1 = "CURVE DATA"
    Range("R5").Select
    ActiveCell.FormulaR1C1 = "REMARKS"
    
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "(M.)"
    Range("E6").Select
    ActiveCell.FormulaR1C1 = "(M.)"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "(M.)"
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "D" & ChrW(&HB0) & " " & "M' S.SS"""
    Range("H6").Select
    ActiveCell.FormulaR1C1 = "(DEG.)"
    Range("I6").Select
    ActiveCell.FormulaR1C1 = "D" & ChrW(&HB0) & " " & "M' S.SS"""
    Range("J6").Select
    ActiveCell.FormulaR1C1 = "(DEG.)"
    Range("K6").Select
    ActiveCell.FormulaR1C1 = "D" & ChrW(&HB0) & " " & "M' S.SS"""
    Range("L6").Select
    ActiveCell.FormulaR1C1 = "(DEG.)"
    Range("M6").Select
    ActiveCell.FormulaR1C1 = "LT/RT"
    Range("N6").Select
    ActiveCell.FormulaR1C1 = "RADIUS (M.)"
    Range("O6").Select
    ActiveCell.FormulaR1C1 = "Ls IN (M.)"
    Range("P6").Select
    ActiveCell.FormulaR1C1 = "Lc (M.)"
    Range("Q6").Select
    ActiveCell.FormulaR1C1 = "Ls OUT (M.)"

    Range("B2").Select
    Selection.Font.Bold = True
    Range("B3:R6").Select
    Selection.Font.Bold = True
    
'--------------------------Format Table (Array) -----------------------'
    
    Sheets(HorArrName).Select
    
    Cells.Select
    Selection.RowHeight = 30
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Columns("B:B").Select
    Selection.ColumnWidth = 25
    Columns("C:D").Select
    Selection.ColumnWidth = 15
    Columns("E:J").Select
    Selection.ColumnWidth = 20
    Columns("K:K").Select
    Selection.ColumnWidth = 15
    Columns("L:L").Select
    Selection.ColumnWidth = 30
    Range("C2:E2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Range("B3:L3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Rows("3:3").Select
    Selection.RowHeight = 40
    ActiveWindow.Zoom = 70
    Range("A1").Select
    
'--------------------------Head Table (Array) -----------------------'

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "ALIGNMENT NAME :"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = Alignment_Name
    
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "HORIZONTAL ALIGNMENT DATA"
    
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "HIP NO."
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "MAIN POINT"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "LOOP NO."
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "CH.START (M.)"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "CH.END (M.)"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "E.START (M.)"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "N.START (M.)"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "AZIMUTH (DEG.)"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "RADIUS (M.)"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "CURVE TYPE"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "REMARK"
    
    Range("L5").Select
    ActiveCell.FormulaR1C1 = "T = Straight Line"
    Range("L6").Select
    ActiveCell.FormulaR1C1 = "SPIN = Spiral Curve In"
    Range("L7").Select
    ActiveCell.FormulaR1C1 = "C = Circular Curve"
    Range("L8").Select
    ActiveCell.FormulaR1C1 = "SPOT = Spiral Curve Out"
    
    Range("B2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("B3:L3").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 13
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("B4:L4").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("L5:L8").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B2").Select

'--------------------------Beginning of Point-----------------------'
    
    Sheets("HIP DATA").Select 'Select sheet name "HIP DATA"
    
    'totalHip = Range("D2") 'Total HIP curve
    HipNo = Range("A4") 'HIP name of BOP
    CBP = Range("B2") 'Chainage of BOP
    EBP = Range("B4") 'Easting of of BOP
    NBP = Range("C4") 'Northing of BOP
    EPINext = Range("B5") 'Next HIP curve
    NPINext = Range("C5") 'Next HIP curve
    AzT2 = ModAzi(DirecDistAz(EBP, NBP, EPINext, NPINext, "A")) 'Azimuth of BOP
    AzT2DMS = DegtoDMSStr(AzT2) 'Convert Deg. to D° M' S"
    
    'Print BOP Setting Out data
    Sheets(HorSOName).Select 'Select sheet of setting out alignment computation
    Range("B7").Select 'Start 0,0 at B7 for index
    
    Dim BOPValue() As Variant
    Dim BOPFormat() As Variant
    BOPValue = Array(HipNo, "BOP", CBP, EBP, NBP, AzT2DMS, AzT2) 'Create array value
    BOPFormat = Array("@", "@", "0+000.000", "0.0000", "0.0000", "@", "0.000") 'Create array format
    For u = LBound(BOPValue) To UBound(BOPValue) 'LBound() is min index and UBound() is max index
        ActiveCell.Offset(0, u - 1).Value = BOPValue(u)
        ActiveCell.Offset(0, u - 1).NumberFormat = BOPFormat(u)
    Next
    ActiveCell.Offset(0, -1).Value = "i=0" & "," & "j=0" 'Print index of i, j

    'Print BOP Array data
    Sheets(HorArrName).Select
    Range("B5").Select 'Start 0,0 at B5 for index
    
    Dim BOPArrValue() As Variant
    Dim BOPArrFormat() As Variant
    BOPArrValue = Array(HipNo, "BOP", "=R[1]C + 1", CBP, "=R[1]C[-1]", EBP, NBP, AzT2, 0, "T")
    BOPArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@")
    For u = LBound(BOPArrValue) To UBound(BOPArrValue)
        ActiveCell.Offset(0, u - 1).Value = BOPArrValue(u)
        ActiveCell.Offset(0, u - 1).NumberFormat = BOPArrFormat(u)
    Next
    ActiveCell.Offset(0, -1).Value = "i=0" & "," & "k=0"

'--------------------------Starting of Curve-----------------------'

    Sheets(HorSOName).Select 'Select sheet of setting out alignment computation
    
    j = 1 'Index of setting out alignment computation
    k = 1 'Index of alignment Array
    For i = 1 To totalHip - 2 'Skip BOP and EOP (-2)

        Sheets("HIP DATA").Select
        Range("A4").Select 'Start 0,0 at A4 for index
    
        HipNo = ActiveCell.Offset(i, 0) 'HIP name
        EPI = ActiveCell.Offset(i, 1) 'HIP Easting
        NPI = ActiveCell.Offset(i, 2) 'HIP Northing
        EPIBack = ActiveCell.Offset(i - 1, 1) 'Back HIP Easting
        NPIBack = ActiveCell.Offset(i - 1, 2) 'Back HIP Northing
        EPINext = ActiveCell.Offset(i + 1, 1) 'Next HIP Easting
        NPINext = ActiveCell.Offset(i + 1, 2) 'Next HIP Northing
        Radius = ActiveCell.Offset(i, 3) 'Radius or curve
        Ls1 = ActiveCell.Offset(i, 4) 'Spiral length in
        Ls2 = ActiveCell.Offset(i, 5) 'Spiral length out
        AzT1 = ModAzi(DirecDistAz(EPIBack, NPIBack, EPI, NPI, "A")) 'Azimuth of tangent 1 (Back PI to PI)
        AzT1DMS = DegtoDMSStr(AzT1) 'Convert Deg. to D° M' S"
        AzT2 = ModAzi(DirecDistAz(EPI, NPI, EPINext, NPINext, "A")) 'Azimuth of tangent 2 (PI to next PI)
        
        'Left turn (LT) or Rigth turn (RT) and Deflection angle
        Dim DefAngle As Double
        Delta = AzT2 - AzT1
        If Abs(Delta) > 180 Then
            DefAngle = Delta - Sgn(Delta) * 360
            DefAngleDMS = DegtoDMSStr(Abs(DefAngle))
        Else
            DefAngle = Delta
            DefAngleDMS = DegtoDMSStr(Abs(DefAngle))
        End If
        
        Dim TurnLR As String
        If DefAngle < 0 Then
            TurnLR = "LT"
        Else
            TurnLR = "RT"
        End If
        
        'Selecte Curve Type
        Dim CurveType As String
        If Radius > 0 Then
            If Ls1 > 0 And Ls2 > 0 Then
                CurveType = "SPIRAL"    'Spiral Curve
            ElseIf Ls1 = 0 And Ls2 = 0 Then
                CurveType = "CIRCULAR"  'Circular Curve
            Else
                CurveType = False
            End If
        ElseIf Radius = 0 Then
            CurveType = "NOCURVE"  'PI No Curve
        Else
            CurveType = False
        End If
        
        '---------------Select Case of Curve type---------------'
        
        Select Case CurveType
            '---------------PI No Curve---------------'
            Case "NOCURVE"
                Sheets(HorSOName).Select
                Range("B7").Select
                CBP = ActiveCell.Offset(j - 1, 2) 'Chainage of BOP for next curve
                EBP = ActiveCell.Offset(j - 1, 3) 'Easting of BOP for next curve
                NBP = ActiveCell.Offset(j - 1, 4) 'Northing of BOP for next curve
                CPI = CBP + DirecDistAz(EPI, NPI, EBP, NBP, "D") 'Chainage of PI no curve
                
                'Print PI no curve data
                Dim PINoCValue() As Variant
                Dim PINoCFormat() As Variant
                PINoCValue = Array(HipNo, "PI", CPI, EPI, NPI, "", "", AzT1DMS, AzT1, DefAngleDMS & " " & TurnLR, Abs(DefAngle), TurnLR, Radius, Ls1, 0, Ls2)
                PINoCFormat = Array("@", "@", "0+000.000", "0.0000", "0.0000", "", "", "@", "0.000", "@", "0.000", "@", "0.000", "0.000", "0.000", "0.000")
                For u = LBound(PINoCValue) To UBound(PINoCValue)
                    ActiveCell.Offset(j, u - 1).Value = PINoCValue(u)
                    ActiveCell.Offset(j, u - 1).NumberFormat = PINoCFormat(u)
                Next
                ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
                j = j + 1 'Next index
            
                'Print PI no curve Array data
                Sheets(HorArrName).Select
                Range("B5").Select 'Start 0,0 at B5 for index
                
                Dim PINoCArrValue() As Variant
                Dim PINoCArrFormat() As Variant
                PINoCArrValue = Array(HipNo, "PI", "=R[1]C + 1", CPI, "=R[1]C[-1]", EPI, NPI, AzT2, 0, "T")
                PINoCArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@")
                For u = LBound(PINoCArrValue) To UBound(PINoCArrValue)
                    ActiveCell.Offset(k, u - 1).Value = PINoCArrValue(u)
                    ActiveCell.Offset(k, u - 1).NumberFormat = PINoCArrFormat(u)
                Next
                ActiveCell.Offset(k, -1).Value = i & "," & k
                k = k + 1
            
            '---------------Circular Curve---------------'
            Case "CIRCULAR"
                'Circular curve parameter
                Lc = Radius * DegtoRad(Abs(DefAngle))
                Tc = Radius * Tan(DegtoRad(Abs(DefAngle) / 2))
                
                'PC point data
                EPC = EPI - Tc * Sin(DegtoRad(AzT1))
                NPC = NPI - Tc * Cos(DegtoRad(AzT1))
                AzPC = AzT1
                AzPCDMS = DegtoDMSStr(AzPC)
                
                'PT point data
                EPT = EPI + Tc * Sin(DegtoRad(AzT2))
                NPT = NPI + Tc * Cos(DegtoRad(AzT2))
                AzPT = AzT2
                AzPTDMS = DegtoDMSStr(AzPT)
                
                Sheets(HorSOName).Select
                Range("B7").Select 'Start 0,0 at B7 for index
                CBP = ActiveCell.Offset(j - 1, 2) 'Chainage of BOP for next curve
                EBP = ActiveCell.Offset(j - 1, 3) 'Easting of BOP for next curve
                NBP = ActiveCell.Offset(j - 1, 4) 'Northing of BOP for next curve
                CPI = CBP + DirecDistAz(EPI, NPI, EBP, NBP, "D") 'Chainage of PI
                CPC = CPI - Tc 'Chainage of PC
                CPT = CPC + Lc 'Chainage of PT
                
                'Print PC data
                Dim PCValue() As Variant
                Dim PCFormat() As Variant
                PCValue = Array("", "PC", CPC, EPC, NPC, AzPCDMS, AzPC)
                PCFormat = Array("", "@", "0+000.000", "0.0000", "0.0000", "@", "0.000")
                For u = LBound(PCValue) To UBound(PCValue)
                    ActiveCell.Offset(j, u - 1).Value = PCValue(u)
                    ActiveCell.Offset(j, u - 1).NumberFormat = PCFormat(u)
                Next
                ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
                
                'Print PI data
                j = j + 1 'Next index
                Dim PICValue() As Variant
                Dim PICFormat() As Variant
                PICValue = Array(HipNo, "PI", CPI, EPI, NPI, "", "", AzT1DMS, AzT1, DefAngleDMS & " " & TurnLR, Abs(DefAngle), TurnLR, Radius, Ls1, Lc, Ls2)
                PICFormat = Array("@", "@", "0+000.000", "0.0000", "0.0000", "", "", "@", "0.000", "@", "0.000", "@", "0.000", "0.000", "0.000", "0.000")
                For u = LBound(PICValue) To UBound(PICValue)
                    ActiveCell.Offset(j, u - 1).Value = PICValue(u)
                    ActiveCell.Offset(j, u - 1).NumberFormat = PICFormat(u)
                Next
                ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
                
                'Print PT data
                j = j + 1 'Next index
                Dim PTValue() As Variant
                Dim PTFormat() As Variant
                PTValue = Array("", "PT", CPT, EPT, NPT, AzPTDMS, AzPT)
                PTFormat = Array("", "@", "0+000.000", "0.0000", "0.0000", "@", "0.000")
                For u = LBound(PTValue) To UBound(PTValue)
                    ActiveCell.Offset(j, u - 1).Value = PTValue(u)
                    ActiveCell.Offset(j, u - 1).NumberFormat = PTFormat(u)
                Next
                ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
                
                j = j + 1 'Next index
            
                'Print PC Array data
                Sheets(HorArrName).Select
                Range("B5").Select 'Start 0,0 at B5 for index
                
                Dim PCArrValue() As Variant
                Dim PCArrFormat() As Variant
                PCArrValue = Array(HipNo, "PC", "=R[1]C + 1", CPC, "=R[1]C[-1]", EPC, NPC, AzPC, Radius * Sgn(DefAngle), "C")
                PCArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@")
                For u = LBound(PCArrValue) To UBound(PCArrValue)
                    ActiveCell.Offset(k, u - 1).Value = PCArrValue(u)
                    ActiveCell.Offset(k, u - 1).NumberFormat = PCArrFormat(u)
                Next
                ActiveCell.Offset(k, -1).Value = i & "," & k
            
                'Print PT Array data
                k = k + 1
                Dim PTArrValue() As Variant
                Dim PTArrFormat() As Variant
                PTArrValue = Array("", "PT", "=R[1]C + 1", CPT, "=R[1]C[-1]", EPT, NPT, AzPT, 0, "T")
                PTArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@")
                For u = LBound(PTArrValue) To UBound(PTArrValue)
                    ActiveCell.Offset(k, u - 1).Value = PTArrValue(u)
                    ActiveCell.Offset(k, u - 1).NumberFormat = PTArrFormat(u)
                Next
                ActiveCell.Offset(k, -1).Value = i & "," & k
                
                k = k + 1
            
            '---------------Spiral Curve---------------'
            Case "SPIRAL"
                
                'Spiral In cumputation
                'Parameter
                Ts1 = SpiralIN(Ls1, Ls2, Radius, Abs(DefAngle), "T")
                Xs1 = SpiralIN(Ls1, Ls2, Radius, Abs(DefAngle), "X")
                Ys1 = SpiralIN(Ls1, Ls2, Radius, Abs(DefAngle), "Y")
                Qs1 = Spiral(Ls1, Radius, Abs(DefAngle), "Q")
                
                'TS point data
                ETS = EPI - Ts1 * Sin(DegtoRad(AzT1))
                NTS = NPI - Ts1 * Cos(DegtoRad(AzT1))
                AzTS = AzT1
                AzTSDMS = DegtoDMSStr(AzTS)
                
                'SC point data
                ESC = CoorYXtoNE(ETS, NTS, AzTS, Xs1, Ys1 * Sgn(DefAngle), "E")
                NSC = CoorYXtoNE(ETS, NTS, AzTS, Xs1, Ys1 * Sgn(DefAngle), "N")
                AzSC = AzTS + RadtoDeg(Qs1) * Sgn(DefAngle)
                AzSCDMS = DegtoDMSStr(AzSC)
                
                'Spiral Out cumputation
                'Parameter
                Ts2 = SpiralOUT(Ls1, Ls2, Radius, Abs(DefAngle), "T")
                Xs2 = SpiralOUT(Ls1, Ls2, Radius, Abs(DefAngle), "X")
                Ys2 = SpiralOUT(Ls1, Ls2, Radius, Abs(DefAngle), "Y")
                Qs2 = Spiral(Ls2, Radius, Abs(DefAngle), "Q")
                
                'ST point data
                EST = EPI + Ts2 * Sin(DegtoRad(AzT2))
                NST = NPI + Ts2 * Cos(DegtoRad(AzT2))
                AzST = AzT2
                AzSTDMS = DegtoDMSStr(AzST)
                                
                'CS point data
                ECS = CoorYXtoNE(EST, NST, AzST, -Xs2, Ys2 * Sgn(DefAngle), "E")
                NCS = CoorYXtoNE(EST, NST, AzST, -Xs2, Ys2 * Sgn(DefAngle), "N")
                AzCS = AzST - RadtoDeg(Qs2) * Sgn(DefAngle)
                AzCSDMS = DegtoDMSStr(AzCS)
                
                'Circular Parameter
                Qc = DegtoRad(Abs(DefAngle)) - (Qs1 + Qs2)
                Lc = Radius * Qc
                                
                Sheets(HorSOName).Select
                Range("B7").Select 'Start 0,0 at B7 for index
                CBP = ActiveCell.Offset(j - 1, 2) 'Chainage of BOP for next curve
                EBP = ActiveCell.Offset(j - 1, 3) 'Easting of BOP for next curve
                NBP = ActiveCell.Offset(j - 1, 4) 'Northing of BOP for next curve
                CPI = CBP + DirecDistAz(EPI, NPI, EBP, NBP, "D") 'Chainage of PI
                CTS = CPI - Ts1 'Chainage of TS
                Csc = CTS + Ls1 'Chainage of SC
                CCS = Csc + Lc 'Chainage of CS
                CST = CCS + Ls2 'Chainage of ST
                
                'Print TS data
                Dim TSValue() As Variant
                Dim TSFormat() As Variant
                TSValue = Array("", "TS", CTS, ETS, NTS, AzTSDMS, AzTS)
                TSFormat = Array("", "@", "0+000.000", "0.0000", "0.0000", "@", "0.000")
                For u = LBound(TSValue) To UBound(TSValue)
                    ActiveCell.Offset(j, u - 1).Value = TSValue(u)
                    ActiveCell.Offset(j, u - 1).NumberFormat = TSFormat(u)
                Next
                ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
                
                'Print SC data
                j = j + 1 'Next index
                Dim SCValue() As Variant
                Dim SCFormat() As Variant
                SCValue = Array("", "SC", Csc, ESC, NSC, AzSCDMS, AzSC)
                SCFormat = Array("", "@", "0+000.000", "0.0000", "0.0000", "@", "0.000")
                For u = LBound(SCValue) To UBound(SCValue)
                    ActiveCell.Offset(j, u - 1).Value = SCValue(u)
                    ActiveCell.Offset(j, u - 1).NumberFormat = SCFormat(u)
                Next
                ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
                
                'Print SC data
                j = j + 1 'Next index
                Dim PISValue() As Variant
                Dim PISormat() As Variant
                PISValue = Array(HipNo, "PI", CPI, EPI, NPI, "", "", AzT1DMS, AzT1, DefAngleDMS & " " & TurnLR, Abs(DefAngle), TurnLR, Radius, Ls1, Lc, Ls2)
                PISFormat = Array("@", "@", "0+000.000", "0.0000", "0.0000", "", "", "@", "0.000", "@", "0.000", "@", "0.000", "0.000", "0.000", "0.000")
                For u = LBound(PISValue) To UBound(PISValue)
                    ActiveCell.Offset(j, u - 1).Value = PISValue(u)
                    ActiveCell.Offset(j, u - 1).NumberFormat = PISFormat(u)
                Next
                ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
                
                'Print CS data
                j = j + 1 'Next index
                Dim CSValue() As Variant
                Dim CSFormat() As Variant
                CSValue = Array("", "CS", CCS, ECS, NCS, AzCSDMS, AzCS)
                CSFormat = Array("", "@", "0+000.000", "0.0000", "0.0000", "@", "0.000")
                For u = LBound(CSValue) To UBound(CSValue)
                    ActiveCell.Offset(j, u - 1).Value = CSValue(u)
                    ActiveCell.Offset(j, u - 1).NumberFormat = CSFormat(u)
                Next
                ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
                
                'Print ST data
                j = j + 1 'Next index
                Dim STValue() As Variant
                Dim STFormat() As Variant
                STValue = Array("", "ST", CST, EST, NST, AzSTDMS, AzST)
                STFormat = Array("", "@", "0+000.000", "0.0000", "0.0000", "@", "0.000")
                For u = LBound(STValue) To UBound(STValue)
                    ActiveCell.Offset(j, u - 1).Value = STValue(u)
                    ActiveCell.Offset(j, u - 1).NumberFormat = STFormat(u)
                Next
                ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
                
                j = j + 1 'Next index
                
                'Print TS Array data
                Sheets(HorArrName).Select
                Range("B5").Select 'Start 0,0 at B5 for index
                
                Dim TSArrValue() As Variant
                Dim TSArrFormat() As Variant
                TSArrValue = Array(HipNo, "TS", "=R[1]C + 1", CTS, "=R[1]C[-1]", ETS, NTS, AzTS, Radius * Sgn(DefAngle), "SPIN")
                TSArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@")
                For u = LBound(TSArrValue) To UBound(TSArrValue)
                    ActiveCell.Offset(k, u - 1).Value = TSArrValue(u)
                    ActiveCell.Offset(k, u - 1).NumberFormat = TSArrFormat(u)
                Next
                ActiveCell.Offset(k, -1).Value = i & "," & k
                
                'Print SC Array data
                k = k + 1
                Dim SCArrValue() As Variant
                Dim SCArrFormat() As Variant
                SCArrValue = Array("", "SC", "=R[1]C + 1", Csc, "=R[1]C[-1]", ESC, NSC, AzSC, Radius * Sgn(DefAngle), "C")
                SCArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@")
                For u = LBound(SCArrValue) To UBound(SCArrValue)
                    ActiveCell.Offset(k, u - 1).Value = SCArrValue(u)
                    ActiveCell.Offset(k, u - 1).NumberFormat = SCArrFormat(u)
                Next
                ActiveCell.Offset(k, -1).Value = i & "," & k
                
                'Print CS Array data
                k = k + 1
                Dim CSArrValue() As Variant
                Dim CSArrFormat() As Variant
                CSArrValue = Array("", "CS", "=R[1]C + 1", CCS, "=R[1]C[-1]", ECS, NCS, AzCS, Radius * Sgn(DefAngle), "SPOT")
                CSArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@")
                For u = LBound(CSArrValue) To UBound(CSArrValue)
                    ActiveCell.Offset(k, u - 1).Value = CSArrValue(u)
                    ActiveCell.Offset(k, u - 1).NumberFormat = CSArrFormat(u)
                Next
                ActiveCell.Offset(k, -1).Value = i & "," & k
                
                'Print ST Array data
                k = k + 1
                Dim STArrValue() As Variant
                Dim STArrFormat() As Variant
                STArrValue = Array("", "ST", "=R[1]C + 1", CST, "=R[1]C[-1]", EST, NST, AzST, 0, "T")
                STArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@")
                For u = LBound(STArrValue) To UBound(STArrValue)
                    ActiveCell.Offset(k, u - 1).Value = STArrValue(u)
                    ActiveCell.Offset(k, u - 1).NumberFormat = STArrFormat(u)
                Next
                ActiveCell.Offset(k, -1).Value = i & "," & k
                
                k = k + 1
            End Select
    Next
        
'--------------------------Ending of Point-----------------------'
    
    Sheets("HIP DATA").Select 'Select sheet name "HIP DATA"
    Range("A4").Select 'Start 0,0 at A4 for index
    
    v = totalHip - 1 'Index last HIP data
    
    HipNo = ActiveCell.Offset(v, 0) 'HIP name of EOP
    EEP = ActiveCell.Offset(v, 1) 'Easting of EOP
    NEP = ActiveCell.Offset(v, 2) 'Northing of EOP
    EPIBack = ActiveCell.Offset(v - 1, 1) 'Back HIP curve
    NPIBack = ActiveCell.Offset(v - 1, 2) 'Back HIP curve
    AzT1 = ModAzi(DirecDistAz(EPIBack, NPIBack, EEP, NEP, "A")) 'Azimuth of EOP
    AzT1DMS = DegtoDMSStr(AzT1) 'Convert Deg. to D° M' S"
 
    Sheets(HorSOName).Select 'Select sheet of setting out alignment computation
    Range("B7").Select 'Start 0,0 at B7 for index
    CBP = ActiveCell.Offset(j - 1, 2) 'Chainage of BOP for next curve
    EBP = ActiveCell.Offset(j - 1, 3) 'Easting of BOP for next curve
    NBP = ActiveCell.Offset(j - 1, 4) 'Northing of BOP for next curve
    CEP = CBP + DirecDistAz(EEP, NEP, EBP, NBP, "D") 'Chainage of EOP
    
    'Print EOP data
    Dim EOPValue() As Variant
    Dim EOPFormat() As Variant
    EOPValue = Array(HipNo, "EOP", CEP, EEP, NEP, AzT1DMS, AzT1)
    EOPFormat = Array("@", "@", "0+000.000", "0.0000", "0.0000", "@", "0.000")
    For u = LBound(BOPValue) To UBound(BOPValue)
        ActiveCell.Offset(j, u - 1).Value = EOPValue(u)
        ActiveCell.Offset(j, u - 1).NumberFormat = EOPFormat(u)
    Next
    ActiveCell.Offset(j, -1).Value = v & "," & j 'Print index of v, j
        
        
    'Print EOP Array data
    Sheets(HorArrName).Select
    Range("B5").Select 'Start 0,0 at B5 for index
    
    Dim EOPArrValue() As Variant
    Dim EOPArrFormat() As Variant
    EOPArrValue = Array(HipNo, "EOP", "=R[1]C + 1", CEP, "=R[0]C[-1]+0.002", EEP, NEP, AzT1, 0, "T")
    EOPArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@")
    For u = LBound(EOPArrValue) To UBound(EOPArrValue)
        ActiveCell.Offset(k, u - 1).Value = EOPArrValue(u)
        ActiveCell.Offset(k, u - 1).NumberFormat = EOPArrFormat(u)
    Next
    ActiveCell.Offset(k, -1).Value = v & "," & k
        
    Sheets(HorSOName).Select
    Range("B7").Select
    ActiveWindow.Zoom = 90
    
    Sheets("HIP DATA").Select
    Range("A4").Select
    MsgBox "Horizontal Alignment Complete!"
    
End Sub


'-----------------Setting out Vertical Alignment Computation-----------------'

Sub VerAlignmentComp()

    Dim lastRow As Long
    lastRow = ThisWorkbook.Sheets("VIP DATA").Cells(Rows.Count, 1).End(xlUp).Row
    totalVip = lastRow - 3 'Total VIP curve
    MsgBox "TOTAL VIP =" & " " & totalVip

    Alignment_Name = Range("B1") 'Alignment name
    VerSOName = "VER-SETTING OUT" 'Sheet name for setting out
    VerArrName = "VER-ARRAY" 'Sheet name for array
    
    Sheets.Add(After:=Sheets("VIP DATA")).Name = VerSOName
    Sheets.Add(After:=Sheets(VerSOName)).Name = VerArrName


'--------------------------Format Table (Setting Out) -----------------------'
    
    Sheets(VerSOName).Select
    
    Cells.Select
    Selection.RowHeight = 20
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("B:B").Select
    Selection.ColumnWidth = 15
    Columns("C:C").Select
    Selection.ColumnWidth = 8
    Columns("D:K").Select
    Selection.ColumnWidth = 15
    Columns("I:I").Select
    Selection.ColumnWidth = 20
    Range("C2:E2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Range("B3:K3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C4:E4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("F4:J4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Rows("3:3").Select
    Selection.RowHeight = 30
    Range("B3:K3").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("A1").Select
    ActiveWindow.Zoom = 90
    
'--------------------------Head Table (Setting Out) -----------------------'

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "ALIGNMENT NAME :"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = Alignment_Name
    
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "SETTING OUT DATA - VERTICAL ALIGNMENT"
    
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "VERTICAL MAIN POINTS"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "VERTICAL ELEMENTS"
    
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "VIP NO."
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "POINT"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "CHAINAGE"
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "ELEVATION"
    Range("F5").Select
    ActiveCell.FormulaR1C1 = "GRADIENT 1"
    Range("G5").Select
    ActiveCell.FormulaR1C1 = "GRADIENT 2"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "LVC"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "LVC 1"
    Range("J5").Select
    ActiveCell.FormulaR1C1 = "LVC 2"
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "REMARK"
    
    Range("D6").Select
    ActiveCell.FormulaR1C1 = "(M.)"
    Range("E6").Select
    ActiveCell.FormulaR1C1 = "(M.)"
    Range("F6").Select
    ActiveCell.FormulaR1C1 = "(%)"
    Range("G6").Select
    ActiveCell.FormulaR1C1 = "(%)"
    Range("H6").Select
    ActiveCell.FormulaR1C1 = "(M.)"
    Range("I6").Select
    ActiveCell.FormulaR1C1 = "(M.)"
    Range("J6").Select
    ActiveCell.FormulaR1C1 = "(M.)"
    
    Range("B2").Select
    Selection.Font.Bold = True
    Range("B3:K6").Select
    Selection.Font.Bold = True
    Range("B2").Select

'--------------------------Format Table (Array) -----------------------'
    
    Sheets(VerArrName).Select
    
    Cells.Select
    Selection.RowHeight = 30
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("B:B").Select
    Selection.ColumnWidth = 25
    Columns("C:D").Select
    Selection.ColumnWidth = 15
    Columns("E:M").Select
    Selection.ColumnWidth = 20
    Columns("N:N").Select
    Selection.ColumnWidth = 30
    Range("C2:E2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Range("B2").Select
    Selection.Font.Bold = True
    Range("B3:N3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Font
        .Name = "Arial"
        .Size = 13
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("B3:N4").Select
    Selection.Font.Bold = True
    Rows("3:3").Select
    Selection.RowHeight = 40
    
    '--------------------------Head Table (Array) -----------------------'

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "ALIGNMENT NAME :"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = Alignment_Name
    
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "VERTICAL ALIGNMENT DATA"
    
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "VIP NO."
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "MAIN POINT"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "LOOP NO."
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "CH.START (M.)"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "CH.END (M.)"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "ELEVATION (M.)"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "GRADIENT 1 (%)"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "GRADIENT 2 (%)"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "LVC (M.)"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "LVC 1 (M.)"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "LVC 2 (M.)"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "CURVE TYPE"
    Range("N4").Select
    ActiveCell.FormulaR1C1 = "REMARK"
    
    Range("N5").Select
    ActiveCell.FormulaR1C1 = "T = Tangent"
    Range("N6").Select
    ActiveCell.FormulaR1C1 = "S = Symmetric Curve"
    Range("N7").Select
    ActiveCell.FormulaR1C1 = "U = Unsymmetric Curve"
    Range("N5:N7").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("B2").Select
    ActiveWindow.Zoom = 70

'--------------------------Beginning of Point-----------------------'

    Sheets("VIP DATA").Select 'Select sheet name "VIP DATA"
    
    'totalVip = Range("D2") 'Total VIP curve
    VipNo = Range("A4") 'VIP name of BOP
    CBP = Range("B2") 'Chainage of BOP
    ELBP = Range("C4") 'Elevation of of BOP
    CPINext = Range("B5") 'Next VIP curve
    ELPINext = Range("C5") 'Next VIP curve
    
    'Gradient (%)
    GBP2 = ((ELPINext - ELBP) / (CPINext - CBP)) * 100
    GBP1 = GBP2

    Sheets(VerSOName).Select 'Select sheet of setting out alignment computation
    Range("B7").Select 'Start 0,0 at B7 for index
    
    'Print BOP data
    Dim BOPValue() As Variant
    Dim BOPFormat() As Variant
    BOPValue = Array(VipNo, "BOP", CBP, ELBP, GBP1, GBP2) 'Create array value
    BOPFormat = Array("@", "@", "0+000.000", "0.000", "0.000", "0.000") 'Create array format
    For u = LBound(BOPValue) To UBound(BOPValue) 'LBound() is min index and UBound() is max index
        ActiveCell.Offset(0, u - 1).Value = BOPValue(u)
        ActiveCell.Offset(0, u - 1).NumberFormat = BOPFormat(u)
    Next
    ActiveCell.Offset(0, -1).Value = "i=0" & "," & "j=0" 'Print index of i, j

    'Print BOP Array data
    Sheets(VerArrName).Select
    Range("B5").Select 'Start 0,0 at B5 for index
    
    Dim BOPArrValue() As Variant
    Dim BOPArrFormat() As Variant
    BOPArrValue = Array(VipNo, "BOP", "=R[1]C + 1", CBP, "=R[1]C[-1]", ELBP, GBP1, GBP2, "", "", "", "T") 'Create array value
    BOPArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "0.000", "0.000", "@") 'Create array format
    For u = LBound(BOPArrValue) To UBound(BOPArrValue) 'LBound() is min index and UBound() is max index
        ActiveCell.Offset(0, u - 1).Value = BOPArrValue(u)
        ActiveCell.Offset(0, u - 1).NumberFormat = BOPArrFormat(u)
    Next
    ActiveCell.Offset(0, -1).Value = "i=0" & "," & "k=0" 'Print index of i,k

'--------------------------Starting of Curve-----------------------'

    Sheets(VerSOName).Select 'Select sheet of setting out alignment computation
    
    j = 1 'Index of setting out alignment computation
    k = 1 'Index of alignment Array
    For i = 1 To totalVip - 2 'Skip BOP and EOP (-2)

        Sheets("VIP DATA").Select
        Range("A4").Select 'Start 0,0 at A4 for index

        VipNo = ActiveCell.Offset(i, 0) 'VIP name
        CPI = ActiveCell.Offset(i, 1) 'Chainage of PVI
        ELPI = ActiveCell.Offset(i, 2) 'Elevation of PVI
        LVC = ActiveCell.Offset(i, 3) 'Length of vertical curve (Symmetry)
        LVC1 = ActiveCell.Offset(i, 4) 'Length 1 of vertical curve (Unsymmetry)
        LVC2 = ActiveCell.Offset(i, 5) 'Length 2 of vertical curve (Unsymmetry)
        CPIBack = ActiveCell.Offset(i - 1, 1) 'Chainage of PVI back
        ELPIBack = ActiveCell.Offset(i - 1, 2) 'Elevation of PVI back
        CPINext = ActiveCell.Offset(i + 1, 1) 'Chainage of PVI next
        ELPINext = ActiveCell.Offset(i + 1, 2) 'Elevation of PVI next

        If LVC <> 0 And LVC1 = 0 And LVC2 = 0 Then 'Symmetric Curve
        
            G1 = ((ELPI - ELPIBack) / (CPI - CPIBack)) * 100 'Gradient 1 (%)
            G2 = ((ELPINext - ELPI) / (CPINext - CPI)) * 100 'Gradient 2 (%)
            CPVC = CPI - LVC / 2 'Chainage of PVC
            ELPVC = ELPI - (G1 / 100) * (LVC / 2) 'Elevation of PVC
            CPVT = CPVC + LVC 'Chainage of PVT
            ELPVT = ELPI + (G2 / 100) * (LVC / 2) 'Elevation of PVT
            CurveType = "S"
            
        ElseIf LVC = 0 And LVC1 <> 0 And LVC2 <> 0 Then 'Symmetric Curve
        
            G1 = ((ELPI - ELPIBack) / (CPI - CPIBack)) * 100 'Gradient 1 (%)
            G2 = ((ELPINext - ELPI) / (CPINext - CPI)) * 100 'Gradient 2 (%)
            CPVC = CPI - LVC1 'Chainage of PVC
            ELPVC = ELPI - (G1 / 100) * LVC1 'Elevation of PVC
            CPVT = CPVC + LVC1 + LVC2 'Chainage of PVT
            ELPVT = ELPI + (G2 / 100) * LVC2 'Elevation of PVT
            CurveType = "U"
        
        End If
        
        Sheets(VerSOName).Select 'Select sheet of setting out alignment computation
        Range("B7").Select 'Start 0,0 at B7 for index
        
        'Print PVC data
        Dim PVCValue() As Variant
        Dim PVCFormat() As Variant
        PVCValue = Array("", "PVC", CPVC, ELPVC, G1, G2) 'Create array value
        PVCFormat = Array("", "@", "0+000.000", "0.000", "0.000", "0.000") 'Create array format
        For u = LBound(PVCValue) To UBound(PVCValue) 'LBound() is min index and UBound() is max index
            ActiveCell.Offset(j, u - 1).Value = PVCValue(u)
            ActiveCell.Offset(j, u - 1).NumberFormat = PVCFormat(u)
        Next
        ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
        
        'Print PVI data
        j = j + 1
        Dim PVIValue() As Variant
        Dim PVIFormat() As Variant
        PVIValue = Array(VipNo, "PVI", CPI, ELPI, "", "", LVC, LVC1, LVC2) 'Create array value
        PVIFormat = Array("@", "@", "0+000.000", "0.000", "0.000", "0.000", "0.000", "0.000", "0.000") 'Create array format
        For u = LBound(PVIValue) To UBound(PVIValue) 'LBound() is min index and UBound() is max index
            ActiveCell.Offset(j, u - 1).Value = PVIValue(u)
            ActiveCell.Offset(j, u - 1).NumberFormat = PVIFormat(u)
        Next
        ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j

        'Print PVT data
        j = j + 1
        Dim PVTValue() As Variant
        Dim PVTFormat() As Variant
        PVTValue = Array("", "PVT", CPVT, ELPVT, G2, G2) 'Create array value
        PVTFormat = Array("", "@", "0+000.000", "0.000", "0.000", "0.000") 'Create array format
        For u = LBound(PVTValue) To UBound(PVTValue) 'LBound() is min index and UBound() is max index
            ActiveCell.Offset(j, u - 1).Value = PVTValue(u)
            ActiveCell.Offset(j, u - 1).NumberFormat = PVTFormat(u)
        Next
        ActiveCell.Offset(j, -1).Value = i & "," & j 'Print index of i, j
        
        j = j + 1
    
        'Print PVC Array data
        Sheets(VerArrName).Select
        Range("B5").Select 'Start 0,0 at B5 for index
        
        Dim PVCArrValue() As Variant
        Dim PVCArrFormat() As Variant
        PVCArrValue = Array(VipNo, "PVC", "=R[1]C + 1", CPVC, "=R[1]C[-1]", ELPVC, G1, G2, LVC, LVC1, LVC2, CurveType) 'Create array value
        PVCArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "0.000", "0.000", "@") 'Create array format
        For u = LBound(PVCArrValue) To UBound(PVCArrValue) 'LBound() is min index and UBound() is max index
            ActiveCell.Offset(k, u - 1).Value = PVCArrValue(u)
            ActiveCell.Offset(k, u - 1).NumberFormat = PVCArrFormat(u)
        Next
        ActiveCell.Offset(k, -1).Value = i & "," & k 'Print index of i,k
    
        'Print PVT Array data
        Sheets(VerArrName).Select
        Range("B5").Select 'Start 0,0 at B5 for index
        
        k = k + 1
        Dim PVTArrValue() As Variant
        Dim PVTArrFormat() As Variant
        PVTArrValue = Array("", "PVT", "=R[1]C + 1", CPVT, "=R[1]C[-1]", ELPVT, G2, G2, "", "", "", "T") 'Create array value
        PVTArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "0.000", "0.000", "@") 'Create array format
        For u = LBound(PVTArrValue) To UBound(PVTArrValue) 'LBound() is min index and UBound() is max index
            ActiveCell.Offset(k, u - 1).Value = PVTArrValue(u)
            ActiveCell.Offset(k, u - 1).NumberFormat = PVTArrFormat(u)
        Next
        ActiveCell.Offset(k, -1).Value = i & "," & k 'Print index of i,k
        
        k = k + 1
        
    Next
    
'--------------------------Ending of Point-----------------------'
    
    Sheets("VIP DATA").Select 'Select sheet name "VIP DATA"
    Range("A4").Select 'Start 0,0 at A4 for index
    
    v = totalVip - 1 'Index last VIP data

    VipNo = ActiveCell.Offset(v, 0) 'VIP name
    CEP = ActiveCell.Offset(v, 1) 'Chainage of end point
    ELEP = ActiveCell.Offset(v, 2) 'Elevation of end point
    CPIBack = ActiveCell.Offset(v - 1, 1) 'Chainage of PVI back
    ELPIBack = ActiveCell.Offset(v - 1, 2) 'Elevation of PVI back
    
    'Gradient 1&2 (%)
    GEP1 = ((ELEP - ELPIBack) / (CEP - CPIBack)) * 100
    GEP2 = GEP1

    Sheets(VerSOName).Select 'Select sheet of setting out alignment computation
    Range("B7").Select 'Start 0,0 at B7 for index
        
    'Print EOP data
    Dim EOPValue() As Variant
    Dim EOPFormat() As Variant
    EOPValue = Array(VipNo, "EOP", CEP, ELEP, GEP1, GEP2) 'Create array value
    EOPFormat = Array("@", "@", "0+000.000", "0.000", "0.000", "0.000") 'Create array format
    For u = LBound(EOPValue) To UBound(EOPValue) 'LBound() is min index and UBound() is max index
        ActiveCell.Offset(j, u - 1).Value = EOPValue(u)
        ActiveCell.Offset(j, u - 1).NumberFormat = EOPFormat(u)
    Next
    ActiveCell.Offset(j, -1).Value = v & "," & j 'Print index of i, j
       
    'Print EOP Array data
    Sheets(VerArrName).Select
    Range("B5").Select 'Start 0,0 at B5 for index
    
    Dim EOPArrValue() As Variant
    Dim EOPArrFormat() As Variant
    EOPArrValue = Array(VipNo, "EOP", "=R[1]C + 1", CEP, "=R[0]C[-1]+0.002", ELEP, GEP1, GEP2, "", "", "", "T") 'Create array value
    EOPArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "0.000", "0.000", "@") 'Create array format
    For u = LBound(PVTArrValue) To UBound(PVTArrValue) 'LBound() is min index and UBound() is max index
        ActiveCell.Offset(k, u - 1).Value = EOPArrValue(u)
        ActiveCell.Offset(k, u - 1).NumberFormat = EOPArrFormat(u)
    Next
    ActiveCell.Offset(k, -1).Value = v & "," & k 'Print index of i,k
       
    Sheets(VerSOName).Select
    Range("B7").Select
    ActiveWindow.Zoom = 90
    
    Sheets("VIP DATA").Select
    Range("A4").Select
    MsgBox "Vertical Alignment Complete!"
    
End Sub


