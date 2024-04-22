Attribute VB_Name = "CrossFall_ProR9"
' Topic; Cross Fall (%) Program
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 19/10/2023
'

Sub CrossFall()

    Dim lastRow As Long
    lastRow = ThisWorkbook.Sheets("X-FALL DATA").Cells(Rows.Count, 2).End(xlUp).Row
    totalSlope = lastRow - 3 'Total Crown Slope
    MsgBox "TOTAL CROWN SLOPE =" & " " & totalSlope

    Alignment_Name = Range("B1") 'Alignment name
    XFallArrName = "XFALL-ARRAY"
    
    Sheets.Add(After:=Sheets("X-FALL DATA")).Name = XFallArrName
    
'--------------------------Format Table (Array) -----------------------'
    
    Sheets(XFallArrName).Select
    
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
    Columns("C:C").Select
    Selection.ColumnWidth = 15
    Columns("D:I").Select
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
    Range("B3:I3").Select
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
    ActiveCell.FormulaR1C1 = "CROSS FALL DATA"
    
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "CROWN NAME"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "LOOP NO."
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "CH.START (M.)"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "CH.END (M.)"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "X-FALL.START (%)"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "X-FALL.END (%)"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "TYPE"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "REMARK"
    
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "V = Vary"
    Range("I6").Select
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
    Range("B3:I3").Select
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
    Range("B4:I4").Select
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
    Range("I5:I6").Select
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

    Sheets("X-FALL DATA").Select
    
    j = 0
    For i = 0 To totalSlope - 2
        
        Sheets("X-FALL DATA").Select
        Range("A4").Select 'Start 0,0 at A4 for index
        
        Name = ActiveCell.Offset(i, 0)
        CH = ActiveCell.Offset(i, 1)
        XFall1 = ActiveCell.Offset(i, 2)
        XFall2 = ActiveCell.Offset(i + 1, 2)
        
        'X-Fall Type
        If XFall1 = XFall2 Then
            XFallType = "N"
        ElseIf XFall1 <> XFall2 Then
            XFallType = "V"
        Else
            XFallType = False
        End If
        
        'Print Array data
        Sheets(XFallArrName).Select
        Range("B5").Select
        
        Dim XFallArrValue() As Variant
        Dim XFallArrFormat() As Variant
        XFallArrValue = Array(Name, "=R[1]C + 1", CH, "=R[1]C[-1]", XFall1, "=R[1]C[-1]", XFallType)
        XFallArrFormat = Array("@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "@")
        For u = LBound(XFallArrValue) To UBound(XFallArrValue)
            ActiveCell.Offset(j, u).Value = XFallArrValue(u)
            ActiveCell.Offset(j, u).NumberFormat = XFallArrFormat(u)
        Next
        ActiveCell.Offset(j, -1).Value = i & "," & j
        
        j = j + 1
        
    Next
        
'--------------------------End Point Array-----------------------'
        
    Sheets("X-FALL DATA").Select
    Range("A4").Select 'Start 0,0 at A4 for index
    
    k = totalSlope - 1
    
    EP_Name = ActiveCell.Offset(k, 0)
    EP_CH = ActiveCell.Offset(k, 1)
    EP_XFall1 = ActiveCell.Offset(k, 2)

    'Print Array data
    Sheets(XFallArrName).Select
    Range("B5").Select
    
    Dim EPArrValue() As Variant
    Dim EPArrFormat() As Variant
    EPArrValue = Array(EP_Name, "=R[1]C + 1", EP_CH, "=R[0]C[-1]+0.002", EP_XFall1, "=R[0]C[-1]", "N")
    EPArrFormat = Array("@", "0", "0+000.000", "0+000.000", "0.000", "0.000", "@")
    For u = LBound(EPArrValue) To UBound(EPArrValue)
        ActiveCell.Offset(j, u).Value = EPArrValue(u)
        ActiveCell.Offset(j, u).NumberFormat = EPArrFormat(u)
    Next
    ActiveCell.Offset(j, -1).Value = k & "," & j

    Sheets("X-FALL DATA").Select
    Range("A4").Select
    MsgBox "Cross Fall (%) Complete!"

End Sub
