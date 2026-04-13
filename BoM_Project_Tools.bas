Attribute VB_Name = "BoM_Project_Tools"
Sub pCounts()
    Dim ws As Worksheet
    Dim mSheet As Worksheet
    Dim rSheet As Worksheet
    Dim WBName As Variant
    Set mSheet = ActiveWorkbook.Worksheets("PROJECT_EQUIPMENT_LIST")
    Set rSheet = ActiveWorkbook.Worksheets("Equipment Report")
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")
    Dim rng As Range
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Setup Excluded Sheet Array
    Call excSetup

    'Update Summary Sheet
    Call sumSheetSet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Get worksheet List
    Call sheetList

    'Select AV Systems
    countAsk.Show
    
    'Show hidden systems
    If psSheet.Range("N3").Value <> True Then
        SetWsVisibility 1, 5
    End If
    
    'Unfilter masterSheet
    On Error Resume Next
    mSheet.ShowAllData
    On Error GoTo 0
    
    'Set cutSheet Array
    sysCount = dataHold.Range("B" & dataHold.Rows.Count).End(xlUp).Row
    ReDim cutSheets(sysCount)
    
    
    ac = 0
    For i = 1 To sysCount
        cutSheets(ac) = dataHold.Range("B" & i).Value
        ac = ac + 1
    Next i
    
    'Clean Report
    rSheet.Activate
    rlastrow = rSheet.Range("A" & Rows.Count).End(xlUp).Row
    If rlastrow > 1 Then
        rSheet.Range("A2:D" & rlastrow).Select
    End If
    Selection.ClearContents
    
    mLastRow = mSheet.Range("A" & Rows.Count).End(xlUp).Row
        
    rOset = 0
    cOset = 0
    If psSheet.Range("P6").Value = True Then
        cOset = cOset + 1
    End If
    If psSheet.Range("P3").Value = True Then
        cOset = cOset + 1
    End If


    Call EquipCountSystem(cOset)
    
    'Remove Blanks
    ActiveWorkbook.Worksheets("Equipment Report").Visible = xlSheetVisible
    rSheet.Activate
    rlastrow = rSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    On Error Resume Next
    Set rng = Range("A2:A" & rlastrow).SpecialCells(xlCellTypeBlanks)
    
    If Err.Number = 0 Then
        rng.EntireRow.Delete
    End If
    On Error GoTo 0
    
    'Sort
    rSheet.Activate
    rlastrow = rSheet.Range("A" & Rows.Count).End(xlUp).Row
    Range("A1:D" & rlastrow).Select
    rSheet.Sort.SortFields.Clear
    rSheet.Sort.SortFields.Add2 Key:=Range("B2:B" & rlastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    rSheet.Sort.SortFields.Add2 Key:=Range("C2:C" & rlastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With rSheet.Sort
        .SetRange Range("A1:D" & rlastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Filter
    Range("A1:D" & rlastrow).Select
    If rSheet.AutoFilterMode = True Then
        Selection.AutoFilter
        Selection.AutoFilter
    Else
        Selection.AutoFilter
    End If
    
    'Show Report
    ActiveWorkbook.Worksheets("Equipment Report").Visible = xlSheetVisible
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Unload countAsk

End Sub
Sub EquipCount(i)
    Dim wb As Workbook
    Dim masterSheet As Worksheet
    Dim rSheet As Worksheet
    Dim ws As Worksheet
    Dim curSheet As Worksheet
    Dim sumSheet As Worksheet
    

    'Setup references to workbook and sheet
    Set wb = ActiveWorkbook
    Set masterSheet = wb.Worksheets("PROJECT_EQUIPMENT_LIST")
    Set rSheet = wb.Worksheets("Equipment Report")
    Set sumSheet = wb.Worksheets("Summary")
  
    'Check for Item ID
    If Not IsEmpty(masterSheet.Range("A" & i).Value) Then
        itemID = masterSheet.Range("A" & i).Value
        'Add Make and Model
        rSheet.Range("A" & i).Value = masterSheet.Range("A" & i).Value
        rSheet.Range("B" & i).Value = masterSheet.Range("B" & i).Value
        rSheet.Range("C" & i).Value = masterSheet.Range("C" & i).Value
        
        'Count Through Systems
        For Each ws In ActiveWorkbook.Worksheets
            excArray = IsInArray(ws.Name, ExcSheets)
            If excArray = False Then
                cutArray = IsInArray(ws.Name, cutSheets)
                If cutArray = True Then
                    ws.Select
                    'Set sheet name
                    curName = ActiveSheet.Name
                    Set curSheet = ActiveWorkbook.Worksheets(curName)
                    'Look for itemID
                    On Error Resume Next
                    curSheet.Cells.Find(What:=itemID, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    
                    'Add Item count to Equipment Report
                    If Err.Number = 0 Then
                        systemType = curSheet.Range("A2").Value
                        itemRow = ActiveCell.Row
                        itemCountRoom = wb.Worksheets(curName).Range("F" & itemRow).Value
                        
                        'Room Count
                        sumSheet.Visible = xlSheetVisible
                        sumSheet.Select
                        sumSheet.Range("B4").Activate
                        sumSheet.Cells.Find(What:=systemType, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                        roomRow = ActiveCell.Row
                        RoomCount = sumSheet.Range("G" & roomRow).Value
                        itemCount = itemCountRoom * RoomCount
                        rSheet.Range("C" & i).Value = rSheet.Range("C" & i).Value + itemCount
                    End If
                    On Error GoTo 0
                End If
            End If
            ActiveWorkbook.Activate
        Next ws
    End If
   
End Sub
Sub EquipCountSystem(cOset)
    Dim wb As Workbook
    Dim masterSheet As Worksheet
    Dim rSheet As Worksheet
    Dim ws As Worksheet
    Dim curSheet As Worksheet
    Dim sumSheet As Worksheet
    

    'Setup references to workbook and sheet
    Set wb = ActiveWorkbook
    Set masterSheet = wb.Worksheets("PROJECT_EQUIPMENT_LIST")
    Set rSheet = wb.Worksheets("Equipment Report")
    Set sumSheet = wb.Worksheets("Summary")
    itemRow = 6
    i = 2
  
        
        'Count Through Systems
        For Each ws In ActiveWorkbook.Worksheets
            excArray = IsInArray(ws.Name, ExcSheets)
            If excArray = False Then
                cutArray = IsInArray(ws.Name, cutSheets)
                If cutArray = True Then
                    ws.Select
                    'Set sheet name
                    curName = ActiveSheet.Name
                    Set curSheet = ActiveWorkbook.Worksheets(curName)
                    systemType = curSheet.Range("A2").Value
                    LastRow = curSheet.Range("A" & Rows.Count).End(xlUp).Row
                    'Look for ID on PROJECT_EQUIPMENT_LIST
                    Do Until itemRow > LastRow
                        If curSheet.Range("A" & itemRow).Value <> "//" Then
                            If curSheet.Range("A" & itemRow).Value <> "" Then
                                itemID = curSheet.Range("A" & itemRow).Value
                                itemCountRoom = wb.Worksheets(curName).Range("F" & itemRow).Value
                                'Room Count
                                sumSheet.Visible = xlSheetVisible
                                sumSheet.Select
                                sumSheet.Range("B4").Activate
                                sumSheet.Cells.Find(What:=systemType, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                                roomRow = ActiveCell.Row
                                qtyColLetter = Col_Letter(11 + cOset)
                                RoomCount = sumSheet.Range(qtyColLetter & roomRow).Value
                                If IsNumeric(itemCountRoom) Then
                                    itemCount = itemCountRoom * RoomCount
                                End If
                                
                                'Check EQ Report and update
                                'Look for itemID
                                On Error Resume Next
                                rSheet.Select
                                rSheet.Range("A2").Activate
                                rSheet.Cells.Find(What:=itemID, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                                rRow = ActiveCell.Row
                                
                                If Err.Number = 0 Then
                                    rSheet.Range("D" & rRow).Value = rSheet.Range("D" & rRow).Value + itemCount
                                    On Error GoTo 0
                                Else
                                    On Error GoTo 0
                                    'Add Make and Model
                                    masterSheet.Select
                                    masterSheet.Range("A1").Activate
                                    On Error Resume Next
                                    masterSheet.Range("A:A").Find(What:=itemID, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                                    idRow = ActiveCell.Row
                                    If idRow > 1 Then
                                        rSheet.Range("A" & i).Value = itemID
                                        rSheet.Range("B" & i).Value = masterSheet.Range("B" & idRow).Value
                                        rSheet.Range("C" & i).Value = masterSheet.Range("C" & idRow).Value
                                        rSheet.Range("D" & i).Value = itemCount
                                        i = i + 1
                                    End If
                                    On Error GoTo 0
                                End If
                            End If
                        End If
                    itemRow = itemRow + 1
                    Loop
                End If
                itemRow = 6
            End If
            ActiveWorkbook.Activate
        Next ws
End Sub
Sub ecSetup()
    Dim ws As Worksheet
    Dim mSheet As Worksheet
    Dim rSheet As Worksheet
    Dim WBName As Variant
    Set mSheet = ActiveWorkbook.Worksheets("PROJECT_EQUIPMENT_LIST")
    Set rSheet = ActiveWorkbook.Worksheets("Equipment Cost")
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")
    Dim rng As Range
        
    'Unfilter masterSheet
    On Error Resume Next
    mSheet.ShowAllData
    On Error GoTo 0
    
    'Set cutSheet Array
    sysCount = dataHold.Range("B" & dataHold.Rows.Count).End(xlUp).Row
    ReDim cutSheets(sysCount)
    
    ac = 0
    For i = 1 To sysCount
        cutSheets(ac) = dataHold.Range("B" & i).Value
        ac = ac + 1
    Next i
    
    'Clean Report
    rSheet.Activate
    rlastrow = rSheet.Range("A" & Rows.Count).End(xlUp).Row
    If rlastrow > 1 Then
        rSheet.Range("A2:D" & rlastrow).Select
    End If
    Selection.ClearContents
    
    mLastRow = mSheet.Range("A" & Rows.Count).End(xlUp).Row
        
    rOset = 0
    cOset = 0

    Call ecSetupSystem(cOset)
    
    'Remove Blanks
    ActiveWorkbook.Worksheets("Equipment Report").Visible = xlSheetVisible
    rSheet.Activate
    rlastrow = rSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    On Error Resume Next
    Set rng = Range("A2:A" & rlastrow).SpecialCells(xlCellTypeBlanks)
    
    If Err.Number = 0 Then
        rng.EntireRow.Delete
    End If
    On Error GoTo 0
    
    'Sort
    rSheet.Activate
    rlastrow = rSheet.Range("A" & Rows.Count).End(xlUp).Row
    Range("A1:D" & rlastrow).Select
    rSheet.Sort.SortFields.Clear
    rSheet.Sort.SortFields.Add2 Key:=Range("B2:B" & rlastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    rSheet.Sort.SortFields.Add2 Key:=Range("C2:C" & rlastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With rSheet.Sort
        .SetRange Range("A1:D" & rlastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Filter
    Range("A1:D" & rlastrow).Select
    If rSheet.AutoFilterMode = True Then
        Selection.AutoFilter
        Selection.AutoFilter
    Else
        Selection.AutoFilter
    End If
    
End Sub
Sub ecSetupSystem(cOset)
    Dim wb As Workbook
    Dim masterSheet As Worksheet
    Dim rSheet As Worksheet
    Dim ws As Worksheet
    Dim curSheet As Worksheet
    Dim sumSheet As Worksheet
    

    'Setup references to workbook and sheet
    Set wb = ActiveWorkbook
    Set masterSheet = wb.Worksheets("PROJECT_EQUIPMENT_LIST")
    Set rSheet = wb.Worksheets("Equipment Cost")
    Set sumSheet = wb.Worksheets("Summary")
    itemRow = 6
    i = 2
  
        
        'Count Through Systems
        For Each ws In ActiveWorkbook.Worksheets
            excArray = IsInArray(ws.Name, ExcSheets)
            If excArray = False Then
                cutArray = IsInArray(ws.Name, cutSheets)
                If cutArray = True Then
                    ws.Select
                    'Set sheet name
                    curName = ActiveSheet.Name
                    Set curSheet = ActiveWorkbook.Worksheets(curName)
                    systemType = curSheet.Range("A2").Value
                    LastRow = curSheet.Range("A" & Rows.Count).End(xlUp).Row
                    'Look for ID on PROJECT_EQUIPMENT_LIST
                    Do Until itemRow > LastRow
                        If curSheet.Range("A" & itemRow).Value <> "//" Then
                            If curSheet.Range("A" & itemRow).Value <> "" Then
                                itemID = curSheet.Range("A" & itemRow).Value
                                'Check EQ Report and update
                                'Look for itemID
                                On Error Resume Next
                                rSheet.Select
                                rSheet.Range("A2").Activate
                                rSheet.Cells.Find(What:=itemID, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                                If Err.Number = 0 Then
                                    On Error GoTo 0
                                Else
                                    On Error GoTo 0
                                    'Add Make and Model
                                    masterSheet.Select
                                    masterSheet.Range("A1").Activate
                                    On Error Resume Next
                                    masterSheet.Range("A:A").Find(What:=itemID, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                                    idRow = ActiveCell.Row
                                    If idRow > 1 Then
                                        rSheet.Range("A" & i).Value = itemID
                                        rSheet.Range("B" & i).Value = masterSheet.Range("B" & idRow).Value
                                        rSheet.Range("C" & i).Value = masterSheet.Range("C" & idRow).Value
                                        i = i + 1
                                    End If
                                    On Error GoTo 0
                                End If
                            End If
                        End If
                    itemRow = itemRow + 1
                    Loop
                End If
                itemRow = 6
            End If
            ActiveWorkbook.Activate
        Next ws
End Sub
Sub mArchive()
    
    
    autofolder = ActiveWorkbook.Worksheets("PROJECT_SETTINGS").Range("L3").Value
    
    If autofolder = True Then
        'Base folder set
        aFile = GetLocalPath(ActiveWorkbook.FullName)
        If aFile = "ODyes" Then
            Exit Sub
        End If
        archName.Show
        BaseFolder = Left(aFile, InStrRev(aFile, "\") - 1)
        Call autofoldercheck
    Else
        MsgBox "Auto Archive and Issue Folders set to False.  Manually archive your files."
    End If
    
    Unload archName

End Sub
Sub ExportAsPDF()
Dim FolderPath As String
FolderPath = "C:\Users\Trainee1\Desktop\PDFs"
MkDir FolderPath
      
    ActiveWorkbook.Sheets.Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=FolderPath & "\Sales", _
        openafterpublish:=False, ignoreprintareas:=False
   
MsgBox "All PDF's have been successfully exported."
End Sub

