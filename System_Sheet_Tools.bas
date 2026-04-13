Attribute VB_Name = "System_Sheet_Tools"
Sub SystemRow()
Dim STL As Worksheet: Set STL = ActiveWorkbook.Worksheets("SYSTEM_TEMPLATE_LOOKUP")
Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    aCell = "B" & ActiveCell.Row
    InsertRow = ActiveCell.Row
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    'Copy Template Row
    excArray = IsInArray(asName.Name, ExcSheets)
    If excArray = False Then
        If asName.Range(aCell).HasFormula = True Then
            STL.Visible = xlSheetVisible
            STL.Select
            Rows("6:6").Select
            Application.CutCopyMode = False
            Selection.Copy
            asName.Select
            Rows(InsertRow).Select
            Selection.Insert Shift:=xlDown
            asName.Range(aCell).Select
            Application.CutCopyMode = False
'            STL.Visible = xlSheetVeryHidden
        End If
    ElseIf asName.Name = "PROJECT_EQUIPMENT_LIST" Then
        LastRow = asName.Cells.Find(What:="//", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
        If LastRow < 3 Then
            LastRow = 3
        End If
        Rows(LastRow + 1).Select
        Selection.EntireRow.Hidden = False
        Application.CutCopyMode = False
        Selection.Copy
        If InsertRow > LastRow Then
            Rows(LastRow - 1).Select
        Else
            Rows(InsertRow).Select
        End If
        Selection.Insert Shift:=xlDown
        asName.Range(aCell).Select
        Application.CutCopyMode = False
        Rows(LastRow + 2).Select
        Selection.EntireRow.Hidden = True
        asName.Range(aCell).Select
    Else
        MsgBox "Rows can only be added to the PROJECT_EQUIPMENT_LIST and System Tabs."
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    

End Sub
Sub ofciRow()
Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
Dim myRange As Range
Dim myRangeAddr As String


    InsertRow = ActiveCell.Row
    asName.Select
    Set myRange = Selection
    myRangeAddr = Selection.Areas(1).Address(False, False)
    myrangeRow = Selection.Rows(1).Row
    myrangeRowCount = Selection.Rows.Count
    aCell = "B" & myrangeRow & ":B" & myrangeRow + myrangeRowCount - 1
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    'Copy Template Row
    excArray = IsInArray(asName.Name, ExcSheets)
    If excArray = False Then
        If asName.Range(aCell).HasFormula = True Then

            ofciAsk.Show
            If ofciAsk.OB1 = True Then
                ofciText = "OFE"
            ElseIf ofciAsk.OB2 = True Then
                ofciText = "OFCI"
            ElseIf ofciAsk.OB3 = True Then
                ofciText = "OFOI"
            ElseIf ofciAsk.OB4 = True Then
                ofciText = ofciAsk.TB1.Value
            End If
            
            asName.Range("G" & myrangeRow & ":H" & myrangeRow + myrangeRowCount - 1).Value = ofciText
            
            asName.Range("G" & myrangeRow & ":H" & myrangeRow + myrangeRowCount - 1).Select
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            asName.Range(aCell).Select
        End If
        Unload ofciAsk
    Else
        MsgBox "The OFCI tool can only be used on System Tabs."
    End If
End Sub
Sub standRow()
Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
Dim myRange As Range

    
    InsertRow = ActiveCell.Row
    asName.Select
    Set myRange = Selection
    myRangeAddr = Selection.Areas(1).Address(False, False)
    myrangeRow = Selection.Rows(1).Row
    myrangeRowCount = Selection.Rows.Count
    aCell = "G" & myrangeRow & ":H" & myrangeRow + myrangeRowCount - 1
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    'Copy Template Row
    excArray = IsInArray(asName.Name, ExcSheets)
    If excArray = False Then
        For Each c In asName.Range(aCell).Cells
            nowcell = c.Address
            asName.Range(nowcell).Select
            If Selection.Interior.Color = 11309970 Then
                cleango = False
                Exit For
            End If
            If Range(nowcell).HasFormula = False Then
                cleango = True
            Else
                cleango = False
            End If
        Next
'        If asName.Range(aCell).HasFormula = False Then
        If cleango = True Then
            asName.Range("G" & myrangeRow & ":G" & myrangeRow + myrangeRowCount - 1).Formula = "=IF(ISNUMBER(J" & InsertRow & ")=TRUE,J" & InsertRow & "*K" & InsertRow & ",J" & InsertRow & ")"
            asName.Range("H" & myrangeRow & ":H" & myrangeRow + myrangeRowCount - 1).Formula = "=IF(ISNUMBER(G" & InsertRow & ")=TRUE,F" & InsertRow & "*G" & InsertRow & ",G" & InsertRow & ")"
            asName.Range("G" & myrangeRow & ":H" & myrangeRow + myrangeRowCount - 1).Select
            asName.Range(aCell).Select
        End If

    Else
        MsgBox "The OFCI Undo tool can only be used on System Tabs."
    End If

End Sub
Sub noteRow()
Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    aCell = "B" & ActiveCell.Row
    InsertRow = ActiveCell.Row
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    'Copy Template Row
    excArray = IsInArray(asName.Name, ExcSheets)
    If excArray = False Then
        asName.Rows(InsertRow - 1).Select
        If ActiveCell.DisplayFormat.Interior.Color = 14270668 Then
            If asName.Range(aCell).HasFormula = True Then
                asName.Rows(InsertRow).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Range("A" & InsertRow & ":H" & InsertRow).Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Selection.Borders(xlInsideVertical).LineStyle = xlNone
                Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                Selection.Interior.Color = 11309970
                Range("I" & InsertRow).Select
                ActiveCell.Interior.Color = 13288897
            End If
        Else
            If asName.Range(aCell).HasFormula = True Then
                asName.Rows(InsertRow).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Range("A" & InsertRow & ":H" & InsertRow).Select
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Selection.Borders(xlInsideVertical).LineStyle = xlNone
                Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                Selection.Interior.Color = 11309970
                Range("I" & InsertRow).Select
                ActiveCell.Interior.Color = 13288897
                asName.Range("A" & InsertRow).Select
                Selection.Font.Bold = True
            End If
        End If
    Else
        MsgBox "Rows can only be added to the PROJECT_EQUIPMENT_LIST and System Tabs."
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
