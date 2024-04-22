Attribute VB_Name = "Cant_ProR9"
' Topic; Cant Program
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 19/10/2023
'

Sub Cant()

    Dim lastRow As Long
    lastRow = ThisWorkbook.Sheets("CANT DATA").Cells(Rows.Count, 2).End(xlUp).Row
    totalPoint = lastRow - 3 'Total Point Offset
    MsgBox "TOTAL POINT OF CANT =" & " " & totalPoint

    Alignment_Name = Range("B1") 'Alignment name
    CantArrName = "CANT-ARRAY"
    
    Sheets.Add(After:=Sheets("CANT DATA")).Name = CantArrName
    
'--------------------------Format Table (Array) -----------------------'
    
    Sheets(CantArrName).Select
    
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
    Columns("E:H").Select
    Selection.ColumnWidth = 20
    Columns("I:I").Select
    Selection.ColumnWidth = 15
    Columns("J:J").Select
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
    Range("B3:J3").Select
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
    ActiveCell.FormulaR1C1 = "CANT DATA"
    
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
    ActiveCell.FormulaR1C1 = "CANT START (MM.)"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "CANT END (MM.)"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "TYPE"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "REMARK"
    
    Range("J5").Select
    ActiveCell.FormulaR1C1 = "V = Vary"
    Range("J6").Select
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
    Range("B3:J3").Select
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
    Range("B4:J4").Select
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
    Range("J5:J6").Select
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

'--------------------------Cant Array -----------------------'

    Sheets("CANT DATA").Select
    
    'totalPoint = Range("D2")
    j = 0
    For i = 0 To totalPoint - 2
        
        Sheets("CANT DATA").Select
        Range("A4").Select 'Start 0,0 at A4 for index
        
        HipNo = ActiveCell.Offset(i, 0)
        Pnt = ActiveCell.Offset(i, 1)
        CH = ActiveCell.Offset(i, 2)
        Cant1 = ActiveCell.Offset(i, 3)
        Cant2 = ActiveCell.Offset(i + 1, 3)
        
        'Cant Type
        If Cant1 = Cant2 Then
            CantType = "N"
        ElseIf Cant1 <> Cant2 Then
            CantType = "V"
        Else
            CantType = False
        End If
        
        'Print Array data
        Sheets(CantArrName).Select
        Range("B5").Select
        
        Dim CantArrValue() As Variant
        Dim CantArrFormat() As Variant
        CantArrValue = Array(HipNo, Pnt, "=R[1]C + 1", CH, "=R[1]C[-1]", Cant1, "=R[1]C[-1]", CantType)
        CantArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0", "0", "@")
        For u = LBound(CantArrValue) To UBound(CantArrValue)
            ActiveCell.Offset(j, u).Value = CantArrValue(u)
            ActiveCell.Offset(j, u).NumberFormat = CantArrFormat(u)
        Next
        ActiveCell.Offset(j, -1).Value = i & "," & j
        
        j = j + 1
        
    Next
        
'--------------------------End Point Array-----------------------'
        
    Sheets("CANT DATA").Select
    Range("A4").Select 'Start 0,0 at A4 for index
    
    k = totalPoint - 1
    
    HipNo = ActiveCell.Offset(k, 0)
    PEP = ActiveCell.Offset(k, 1)
    CEP = ActiveCell.Offset(k, 2)
    CantEP = ActiveCell.Offset(k, 3)

    'Print Array data
    Sheets(CantArrName).Select
    Range("B5").Select
    
    Dim EPArrValue() As Variant
    Dim EPArrFormat() As Variant
    EPArrValue = Array(HipNo, "EOP", "=R[1]C + 1", CEP, "=R[0]C[-1]+0.002", CantEP, "=R[0]C[-1]", "N")
    EPArrFormat = Array("@", "@", "0", "0+000.000", "0+000.000", "0", "0", "@")
    For u = LBound(EPArrValue) To UBound(EPArrValue)
        ActiveCell.Offset(j, u).Value = EPArrValue(u)
        ActiveCell.Offset(j, u).NumberFormat = EPArrFormat(u)
    Next
    ActiveCell.Offset(j, -1).Value = k & "," & j

    Sheets("CANT DATA").Select
    Range("A4").Select
    MsgBox "Cant Complete!"

End Sub

