Attribute VB_Name = "Track_Changes"
Sub trackC()
Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet

    aCell = "B" & ActiveCell.Row
    InsertRow = ActiveCell.Row
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    'Change Highlight Color
    excArray = IsInArray(asName.Name, ExcSheets)
    If excArray = False Then
        If asName.Range(aCell).HasFormula = True Then
            trackChangeAsk.Show
            asName.Range("AE" & InsertRow).Value = trackChangeAsk.ComboBox1.Value
            asName.Range("AH" & InsertRow).Value = asName.Range("F" & InsertRow).Value
            asName.Range("F" & InsertRow).Value = trackChangeAsk.TextBox1.Value
            
            asName.Range("A" & InsertRow & ":H" & InsertRow).Select
            If asName.Range("F" & InsertRow).Value < asName.Range("AH" & InsertRow).Value Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            ElseIf asName.Range("F" & InsertRow).Value > asName.Range("AH" & InsertRow).Value Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
           ElseIf asName.Range("F" & InsertRow).Value = asName.Range("AH" & InsertRow).Value Then
                With Selection.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
            asName.Range("A" & InsertRow).Select
            Unload trackChangeAsk
        End If
    End If

End Sub
Sub trackIssue()
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Dim LastRow As Range
    Dim c As Range
    
    Set LastRow = asName.Cells.Find(What:="TOTAL INSTALLED COST", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    InsertRow = 5
    Do Until InsertRow = LastRow.Row
        If asName.Range("AE" & InsertRow).Value = revAsk.ComboBox1.Value Then
            asName.Range("A" & InsertRow & ":H" & InsertRow).Select
            If asName.Range("F" & InsertRow).Value < asName.Range("AH" & InsertRow).Value Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            ElseIf asName.Range("F" & InsertRow).Value > asName.Range("AH" & InsertRow).Value Then
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            ElseIf asName.Range("F" & InsertRow).Value = asName.Range("AH" & InsertRow).Value Then
                With Selection.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        ElseIf IsEmpty(asName.Range("F" & InsertRow).Value) = True Then
            If asName.Range("F" & InsertRow).Value = 0 Then
                asName.Rows(InsertRow & ":" & InsertRow).Select
                Selection.EntireRow.Hidden = True
            End If
        End If
        InsertRow = InsertRow + 1
    Loop
End Sub
