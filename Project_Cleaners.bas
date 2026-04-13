Attribute VB_Name = "Project_Cleaners"
Sub CleanSheet(EPC)
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")
    Dim ECsheet As Worksheet
    Dim sec() As Long
    Dim k As Integer
    Dim LastRow As Range
    Dim c As Range
    Dim n As Long
    Dim m As Integer
    Dim cell
    Dim rng As Range
    
    i = 5
    sCount = 0
    'Find Array Size
    Do While endCellfound = False
        asName.Range("A" & i).Select
        cellColor = ActiveCell.Interior.Color
        If cellColor = 14270668 Then
            sCount = sCount + 1
        ElseIf cellColor = 14277081 Or cellColor = 13288897 Then
            endCellfound = True
        End If
        i = i + 1
    Loop
    
    'Size array
    ReDim sec(sCount)
    
    i = 5
    secPOS = 0
    endCellfound = False
    'Populate Array
    Do While endCellfound = False
        asName.Range("A" & i).Select
        cellColor = ActiveCell.Interior.Color
        If cellColor = 14270668 Then
            sec(secPOS) = ActiveCell.Row
            secPOS = secPOS + 1
        ElseIf cellColor = 14277081 Or cellColor = 13288897 Then
            endCellfound = True
            sec(secPOS) = ActiveCell.Row
            
        End If
        i = i + 1
    Loop
    
    ''Check endCell found
    If endCellfound = True Then
        'Highlight change rows and remove old 0 rows
        Call trackIssue
        
        If EPC = False Then
            ''Clean formulas from parts list
            Set LastRow = asName.Cells.Find(What:="TOTAL INSTALLED COST", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
            Set sublastrow = asName.Cells.Find(What:="//", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
            asName.Range("A1:A3").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            asName.Range("A6:G" & LastRow.Row).Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
        ElseIf EPC = True Then
            ''Clean formulas from parts list
            Set LastRow = asName.Cells.Find(What:="TOTAL INSTALLED COST", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
            If psSheet.Range("P3").Value = True Then
                Set sublastrow = asName.Cells.Find(What:="LICENSE", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                sOset = 1
            ElseIf psSheet.Range("P3").Value = False Then
                Set sublastrow = asName.Cells.Find(What:="//", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                sOset = 0
            End If
            asName.Range("A1:A3").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            asName.Range("A6:F" & LastRow.Row).Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
            'Find if Equiement Cost sheet is present
            wbhere = WorksheetExists("Equipment Cost")
            
            ''Clean budget cost and set formulas
            If wbhere = True Then
                Set ECsheet = ActiveWorkbook.Worksheets("Equipment Cost")
                Set rng = asName.Range("G6:G" & sec(UBound(sec)))
                For Each cell In rng
                    If cell.HasFormula = True Then
                        cell.Formula = "=IF($A" & cell.Row & "="""","""",INDEX('" & ECsheet.Name & "'!$A$2:$D$1001,MATCH($A" & cell.Row & ",'" & ECsheet.Name & "'!$A$2:$A$1001,0),4))"
                    End If
                Next cell
            Else
                Set rng = asName.Range("G6:G" & sec(UBound(sec)))
                For Each cell In rng
                    If cell.HasFormula = True Then
                        cell.ClearContents
                    End If
                Next cell
            End If
            asName.Range("A1").Select
            
            For k = LBound(sec) To UBound(sec) - 1
                asName.Range("I" & sec(k)).Formula = "=SUM(H" & sec(k) & ":H" & sec(k + 1) - 1 & ")"
            Next k
        
            'Update Cost Summary
            asName.Range("D" & sec(UBound(sec))).Select
            asName.Range("D" & sec(UBound(sec))).Value = "COST SUMMARY"
            asName.Range("I" & sec(UBound(sec)) + 1).Formula = "=SUM(I" & sec(0) & ":I" & sec(UBound(sec)) & ")"
            asName.Range("I" & sec(UBound(sec)) + 4 + sOset & ":I" & LastRow.Row - 4).Value = "0"
            'Delete Estimate Items
            asName.Rows(LastRow.Row - 12 - sOset & ":" & LastRow.Row - 11 - sOset).Delete
            asName.Rows(LastRow.Row - 2 & ":" & LastRow.Row - 1).Delete
            
            If psSheet.Range("P3").Value = True Then
                asName.Range("I" & LastRow.Row).Formula = "=SUM(I" & sec(UBound(sec)) + 1 & ",I" & sec(UBound(sec)) + 2 & ",I" & sec(UBound(sec)) + 10 & ")"
                asName.Range("I" & sec(UBound(sec)) + 10).Formula = "=SUM(I" & sec(UBound(sec)) + 3 & ":I" & LastRow.Row - 2 & ")"
            Else
                asName.Range("I" & LastRow.Row).Formula = "=SUM(I" & sec(UBound(sec)) + 1 & ",I" & sec(UBound(sec)) + 10 & ")"
                asName.Range("I" & sec(UBound(sec)) + 10).Formula = "=SUM(I" & sec(UBound(sec)) + 2 & ":I" & LastRow.Row - 2 & ")"
            End If


'            asName.Range("D" & UBound(sec)).ClearContents
            
            
            ''Hide Notes
            If PubAsk.OptionButton1.Value = True Then
                Columns("D:F").Select
                Selection.EntireColumn.Hidden = False
            ElseIf PubAsk.OptionButton2.Value = True Then
                Columns("E:E").Select
                Selection.EntireColumn.Hidden = True
            End If
        ElseIf EPC = False Then
            ''Clean formulas from parts list
            Set LastRow = asName.Cells.Find(What:="TOTAL INSTALLED COST", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
            Set sublastrow = asName.Cells.Find(What:="//", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
            asName.Range("A1:A3").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            asName.Range("A6:G" & LastRow.Row).Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            
            'Hide Notes
            If noteStat = True Then
                Columns("D:F").Select
                Selection.EntireColumn.Hidden = False
            ElseIf noteStat = False Then
                Columns("E:E").Select
                Selection.EntireColumn.Hidden = True
            End If
            
            'Remove Cost
            If costStat = True Then
                Columns("F:I").Select
                Selection.EntireColumn.Hidden = False
            ElseIf costStat = False Then
                asName.Range("I5:I" & LastRow.Row).Select
                Selection.Copy
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                
                asName.Range("H" & sec(UBound(sec)) & ":H" & LastRow.Row).Select
                Selection.Cut
                asName.Range("F" & sec(UBound(sec)) & ":F" & LastRow.Row).Select
                asName.Paste
                asName.Range("I:I").Select
                Selection.Copy
                asName.Range("G:G").Select
                asName.Paste
                Columns("H:I").Select
                Selection.EntireColumn.Delete
            End If
            
'            'Hide Equipment
'            If equipStat = True Then
'                Columns("B:E").Select
'                Selection.EntireColumn.Hidden = False
'            ElseIf equipStat = False Then
'                Columns("C:D").Select
'                Selection.EntireColumn.Delete
'            End If
'
            'Remove labor breakout
            If laborStat = True Then
                Rows(LastRow.Row - 10 & ":" & LastRow.Row).Select
                Selection.EntireRow.Hidden = False
            ElseIf laborStat = False Then
                asName.Range(LastRow.Row - 3 & ":" & LastRow.Row - 3).Select
                Selection.Copy
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                Rows(LastRow.Row - 10 & ":" & LastRow.Row - 4).Select
                Selection.EntireRow.Delete
            End If
        End If
        
        ''Clean sheet
        asName.Range("J:AL").Delete
        asName.Range("A" & sec(UBound(sec))).ClearContents
        
        'AutoFit Rows
        Rows("1:" & LastRow.Row).Select
        Selection.Rows.AutoFit
        Range("A6").Select
        
        ''Remove Page Brakes/Set View
        ActiveWindow.View = xlPageBreakPreview

        'colapse blank sections
        For m = 5 To sublastrow.Row - 1
            
            secArray = IsInArray(CStr(m), sec)
            If secArray = True Then
                If IsEmpty(ActiveWorkbook.ActiveSheet.Range("A" & m + 1).Value) = True Then
                    ActiveWorkbook.ActiveSheet.Rows(m).ShowDetail = False
                End If
            End If
        Next m
        
        ActiveSheet.PageSetup.LeftFooter = "&""Veranda""&8" & Range("A1") & Chr(13) & Range("A3")
        
        asName.Range("A1").Select
    End If
        
    
End Sub
Sub cleanWorkbook()
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim tempSheet As Worksheet: Set tempSheet = ActiveWorkbook.Worksheets("_TEMP")
    
    'Clean Workbook
    For Each ws In Worksheets
        excArray = IsInArray(ws.Name, ExcSheets)
        If excArray = True Then
            If ws.Name = "Summary" Then
                'Using this to keep this worksheet in BOM and EoPC
'                ws.Protect Password:=pw, UserInterfaceOnly:=True
            ElseIf ws.Name = "Issuances" Then
                'Using this to keep this worksheet in BOM and EoPC
'                ws.Protect Password:=pw, UserInterfaceOnly:=True
            ElseIf ws.Name = "Revision List" Then
                'Using this to keep this worksheet in BOM and EoPC
'                ws.Protect Password:=pw, UserInterfaceOnly:=True
            ElseIf ws.Name = "Equipment Cost" Then
                'Using this to keep this worksheet in BOM and EoPC
'                ws.Protect Password:=pw, UserInterfaceOnly:=True
            Else
                If ws.Name = "DATA_HOLD" Then
                    dataHold.Visible = xlSheetVisible
                End If
                If ws.Name = "_TEMP" Then
                    tempSheet.Visible = xlSheetVisible
                End If
                If ws.Name = "PROJECT_SETTINGS" Then
                    ActiveWorkbook.Worksheets("PROJECT_SETTINGS").Visible = xlSheetVeryHidden
                Else
                    ws.Delete
                End If
            End If
        End If
    Next ws
End Sub

