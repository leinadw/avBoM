Attribute VB_Name = "Presentaion_Modes"
Sub presentationmode()
    Dim ws As Worksheet
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    
        'Setup Excluded Sheet Array
        Call excSetup

        'Get Worksheet List
        Call sheetList

        'Select AV systems
        PresAsk.Show

        'Set pubSheet Array
        sysCount = dataHold.Range("B" & dataHold.Rows.Count).End(xlUp).Row
        
        ReDim presSheets(sysCount)
         
        ac = 0
        For i = 1 To sysCount
            presSheets(ac) = dataHold.Range("B" & i).Value
            ac = ac + 1
        Next i

        For Each ws In Worksheets
            excArray = IsInArray(ws.Name, ExcSheets)
            If excArray = False Then
                presArray = IsInArray(ws.Name, presSheets)
                If presArray = True Then
                        ws.Select
                        Set asName = ActiveWorkbook.ActiveSheet
                        'Find last rows
                        Set LastRow = asName.Cells.Find(What:="TOTAL INSTALLED COST", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                        Set sublastrow = asName.Cells.Find(What:="//", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                        
                        'Copy text to unhidden cells
                        asName.Range("H" & sublastrow.Row & ":H" & LastRow.Row).Copy
                        asName.Range("F" & sublastrow.Row).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                        
                        'Hide Item Pricing
                        asName.Range("G:H").Select
                        Selection.EntireColumn.Hidden = True
                        
                        If PresAsk.OptionButton2.Value = True Then
                            'Move text to unhidden cells
                            asName.Range("C2:D2").Copy
                            asName.Range("E2").PasteSpecial
                            
                            'Hide MFR and Model
                            asName.Range("C:D").Select
                            Selection.EntireColumn.Hidden = True
                        End If
                        Application.CutCopyMode = False
                End If
            End If
        Next ws
        
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    
        Unload PresAsk
End Sub
Sub workmode()
    Dim ws As Worksheet
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim returnName As Worksheet: Set returnName = ActiveWorkbook.ActiveSheet
    
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        'Setup Excluded Sheet Array
        Call excSetup
    
        'Set presSheet Array
        sysCount = dataHold.Range("B" & dataHold.Rows.Count).End(xlUp).Row

        ReDim presSheets(sysCount)
        
        ac = 0
        For i = 1 To sysCount
            presSheets(ac) = dataHold.Range("B" & i).Value
            ac = ac + 1
        Next i
        
        For Each ws In Worksheets
            excArray = IsInArray(ws.Name, ExcSheets)
            If excArray = False Then
                presArray = IsInArray(ws.Name, presSheets)
                    If presArray = True Then
                        ws.Select
                        Set asName = ActiveWorkbook.ActiveSheet
                        'Find last rows
                        Set LastRow = asName.Cells.Find(What:="TOTAL INSTALLED COST", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                        Set sublastrow = asName.Cells.Find(What:="//", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                        
                        'unhide hidden rows
                        ActiveSheet.Cells.EntireColumn.Hidden = False
                        
                        asName.Range("E2:F2").ClearContents
                        asName.Range("F" & sublastrow.Row & ":F" & LastRow.Row).ClearContents
                        
                        asName.Range("F" & sublastrow.Row & ":F" & LastRow.Row).Select
                        Selection.Borders(xlEdgeRight).LineStyle = xlNone
                        
                        Application.CutCopyMode = False
                        asName.Range("A6").Select
                    End If
            End If
        Next ws
        
        dataHold.Visible = xlSheetVisible
        dataHold.Activate
        dataHold.Range("B:B").Select
        Selection.ClearContents
        dataHold.Visible = xlSheetVeryHidden
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        returnName.Activate
End Sub
