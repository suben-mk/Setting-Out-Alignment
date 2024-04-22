Attribute VB_Name = "Tunnel_AlignmentProR9"
' Topic; Tuunel Alignment Program
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 19/10/2023
'

Sub Tunnel_Alignment()

    Dim lastRow As Long
    lastRow = ThisWorkbook.Sheets("TUNNEL OFFSET DATA").Cells(Rows.Count, 2).End(xlUp).Row
    totalPoint = lastRow - 3 'Total Point Offset
    MsgBox "TOTAL POINT OF TUNNEL OFFSET =" & " " & totalPoint

    Alignment_Name = Range("B1") 'Alignment name
    TunArrName = "TUA-ARRAY"
    
    Sheets.Add(After:=Sheets("TUNNEL OFFSET DATA")).Name = TunArrName
    
'--------------------------Format Table (Array) -----------------------'
    
    Sheets(TunArrName).Select
    
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
    Columns("K:L").Select
    Selection.ColumnWidth = 15
    Columns("M:M").Select
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
    Range("B3:M3").Select
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
    ActiveCell.FormulaR1C1 = "TUNNEL ALIGNMENT DATA"
    
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
    ActiveCell.FormulaR1C1 = "HOR.OS START (M.)"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "HOR.OS END (M.)"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "VER.OS START (M.)"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "VER.OS END (M.)"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "HOR. TYPE"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "VER. TYPE"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "REMARK"
    
    Range("M5").Select
    ActiveCell.FormulaR1C1 = "V = Vary"
    Range("M6").Select
    ActiveCell.FormulaR1C1 = "N = Normal"
    
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
    Range("B3:M3").Select
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
    Range("B4:M4").Select
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
    Range("M5:M6").Select
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

'--------------------------Tunnel Alignment Array -----------------------'

    Sheets("TUNNEL OFFSET DATA").Select
    
    'totalPoint = Range("D2")
    j = 0
    For i = 0 To totalPoint - 2
        
        Sheets("TUNNEL OFFSET DATA").Select
        Range("A4").Select 'Start 0,0 at A4 for index
        
        HipNo = ActiveCell.Offset(i, 0)
        Pnt = ActiveCell.Offset(i, 1)
        CH = ActiveCell.Offset(i, 2)
        HorOS1 = ActiveCell.Offset(i, 3)
        HorOS2 = ActiveCell.Offset(i + 1, 3)
        VerOS1 = ActiveCell.Offset(i, 4)
        VerOS2 = ActiveCell.Offset(i + 1, 4)
        
        'Tunnel Offset Type
        If HorOS1 = HorOS2 Then
            HorType = "N"
        ElseIf HorOS1 <> HorOS2 Then
            HorType = "V"
        Else
            HorType = False
        End If

        If VerOS1 = VerOS2 Then
            VerType = "N"
        ElseIf VerOS1 <> VerOS2 Then
            VerType = "V"
        Else
            VerType = False
        End If

        
        'Print Array data
        Sheets(TunArrName).Select
        Range("B5").Select
        
        Dim TunArrValue() As Variant
        Dim TunArrFormat() As Variant
        TunArrValue = Array(HipNo, Pnt, "=R[1]C + 1", CH, "=R[1]C[-1]", HorOS1, "=R[1]C[-1]", VerOS1, "=R[1]C[-1]", HorType, VerType)
        TunArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@", "@")
        For u = LBound(TunArrValue) To UBound(TunArrValue)
            ActiveCell.Offset(j, u).Value = TunArrValue(u)
            ActiveCell.Offset(j, u).NumberFormat = TunArrFormat(u)
        Next
        ActiveCell.Offset(j, -1).Value = i & "," & j
        
        j = j + 1
        
    Next
        
'--------------------------End Point Array-----------------------'
        
    Sheets("TUNNEL OFFSET DATA").Select
    Range("A4").Select 'Start 0,0 at A4 for index
    
    k = totalPoint - 1
    
    HipNo = ActiveCell.Offset(k, 0)
    PEP = ActiveCell.Offset(k, 1)
    CEP = ActiveCell.Offset(k, 2)
    HorOSEP = ActiveCell.Offset(k, 3)
    VerOSEP = ActiveCell.Offset(k, 4)

    'Print Array data
    Sheets(TunArrName).Select
    Range("B5").Select
    
    Dim EPArrValue() As Variant
    Dim EPArrFormat() As Variant
    EPArrValue = Array(HipNo, "EOP", "=R[1]C + 1", CEP, "=R[0]C[-1]+0.002", HorOSEP, "=R[0]C[-1]", VerOSEP, "=R[0]C[-1]", "N", "N")
    EPArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "0.000", "0.000", "@", "@")
    For u = LBound(EPArrValue) To UBound(EPArrValue)
        ActiveCell.Offset(j, u).Value = EPArrValue(u)
        ActiveCell.Offset(j, u).NumberFormat = EPArrFormat(u)
    Next
    ActiveCell.Offset(j, -1).Value = k & "," & j

    Sheets("TUNNEL OFFSET DATA").Select
    Range("A4").Select
    MsgBox "Tunnel Alignment Complete!"

End Sub
