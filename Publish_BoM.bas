Attribute VB_Name = "Publish_BoM"
Sub PubBoM()
    Dim ws As Worksheet
    Dim WBName As Variant
    Dim wsName As Variant
    Dim iSheet As Worksheet: Set iSheet = ActiveWorkbook.Worksheets("Issuances")
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim sumSheet As Worksheet: Set sumSheet = ActiveWorkbook.Worksheets("Summary")
    Dim revSheet As Worksheet: Set revSheet = ActiveWorkbook.Worksheets("Revision List")
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")
    Dim ECsheet As Worksheet

    
    autofolder = ActiveWorkbook.Worksheets("PROJECT_SETTINGS").Range("L3").Value
    
    If autofolder = True Then
        'Base folder set
        aFile = GetLocalPath(ActiveWorkbook.FullName)
        If aFile = "ODyes" Then
            Exit Sub
        End If
        BaseFolder = Left(aFile, InStrRev(aFile, "\") - 1)
        Call autofoldercheck
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    'Get Worksheet List
    Call sheetList
    
    'Select AV systems
    PubAsk.Show
    
    If dataHold.Range("B1").Value = "" Then
        MsgBox "Please select systems to publish before you can contiunue."
        PubAsk.Show
        If dataHold.Range("B1").Value = "" Then
            Exit Sub
        End If
    End If
    
    'Set pubSheet Array
    sysCount = dataHold.Range("B" & dataHold.Rows.Count).End(xlUp).Row
   
    ReDim pubSheets(sysCount)
    
    ac = 0
    For i = 1 To sysCount
        pubSheets(ac) = dataHold.Range("B" & i).Value
        ac = ac + 1
    Next i
    
    'Show hidden systems
    If psSheet.Range("N3").Value <> True Then
        SetWsVisibility 1, 5
    End If
        
    'Select Issuance and update revisions
    Call revUp
    Call sumSheetSet

    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    
    ''Set publish to BOM
    EPC = True
    
    'Clean worksheets
    For Each ws In Worksheets
        excArray = IsInArray(ws.Name, ExcSheets)
        If excArray = False Then
            pubArray = IsInArray(ws.Name, pubSheets)
            If pubArray = True Then
                ws.Select
                Call CleanSheet(EPC)
            Else
                ws.Delete
            End If
        End If
    Next ws
    
    'Setup Equipment Cost Worksheet
    'Find if Equiement Cost sheet is present
    wbhere = WorksheetExists("Equipment Cost")
            
    ''Clean budget cost and set formulas
    If wbhere = True Then
        Set ECsheet = ActiveWorkbook.Worksheets("Equipment Cost")
        Call ecSetup
    End If
 
    'Update Summary Sheet
    Call sumSheetSet
    
    'set header footer
    If revAsk.ComboBox1.Value = "Add Issuance" Then
        iName = revAsk.TextBox1.Value
    Else
        iName = revAsk.ComboBox1.Value
    End If
    
    sumSheet.Range("A3").Value = iName
    sumSheet.PageSetup.LeftFooter = "&""Veranda""&8" & Range("A1") & Chr(13) & Range("A3")
    iSheet.Range("A3").Value = iName
    iSheet.PageSetup.LeftFooter = "&""Veranda""&8" & Range("A1") & Chr(13) & Range("A3")
    revSheet.Range("A3").Value = iName
    revSheet.PageSetup.LeftFooter = "&""Veranda""&8" & Range("A1") & Chr(13) & Range("A3")
        
    'Clean Workbook
    Call cleanWorkbook
    
    Application.ScreenUpdating = True
    
    Application.CutCopyMode = False
    
    If autofolder = True Then
        exType = "27 41 16 - Appendix A_"
        Call autofoldersave(BaseFolder, exType)
    Else
        WBName = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", InitialFileName:="27 41 16 - Appendix A.xlsx", Title:="Select Location for Appendix A Save")
        If WBName <> False Then
            ActiveWorkbook.SaveAs fileName:=WBName, FileFormat:=51, ConflictResolution:=xlLocalSessionChanges
        End If
    End If
    
    If PubAsk.OptionButton4 = True Then
        PDFName = Left(ActiveWorkbook.FullName, InStrRev(ActiveWorkbook.FullName, ".") - 1)
        Dim shtarr() As Variant
        
        ReDim shtarr(0 To 0)
        For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
        If Len(shtarr(0)) = 0 Then
            shtarr(0) = ws.Name
          Else
            ReDim Preserve shtarr(0 To UBound(shtarr) + 1)
            shtarr(UBound(shtarr)) = ws.Name
          End If
        End If
        Next ws
         
        'Selection addendum
        Dim arrws As Variant
        If Len(shtarr(0)) > 0 Then
        Set arrws = Sheets(shtarr)
        arrws.Select
        End If
        
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=PDFName, openafterpublish:=False, ignoreprintareas:=False
        
        sumSheet.Select
    End If
    
    Unload PubAsk
    Unload revAsk
    
    OneDriveYes = InStr(ActiveWorkbook.FullName, "http")
    If OneDriveYes > 0 Then
        ActiveWorkbook.AutoSaveOn = False
    End If
    
    Application.CutCopyMode = False
            
    Application.DisplayAlerts = True
    
    
    
End Sub

Private Sub contActive_Click()
    textlen = Len(TextBox1.text)
    If textlen > 255 Then
       MsgBox "Your system name must not exceed 31 character and not include any of the following characters: : \ / ?.  Please enter the name again."
    End If
End Sub

Sub pList()
    Dim ws As Worksheet
    Dim mSheet As Worksheet
    Dim pSheet As Worksheet
    Dim WBName As Variant
    Set mSheet = ActiveWorkbook.Worksheets("PROJECT_EQUIPMENT_LIST")
    Set pSheet = ActiveWorkbook.Worksheets("Equpiment Cost")
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")
    Dim rng As Range
    
    'Unfilter mastepSheet
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
    pSheet.Activate
    rlastrow = pSheet.Range("A" & Rows.Count).End(xlUp).Row
    If rlastrow > 1 Then
        pSheet.Range("A2:D" & rlastrow).Select
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

    Call EquipListSystem(cOset)
    
    'Remove Blanks
    ActiveWorkbook.Worksheets("Equipment Report").Visible = xlSheetVisible
    pSheet.Activate
    rlastrow = pSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    On Error Resume Next
    Set rng = Range("A2:A" & rlastrow).SpecialCells(xlCellTypeBlanks)
    
    If Err.Number = 0 Then
        rng.EntireRow.Delete
    End If
    On Error GoTo 0
    
    'Sort
    pSheet.Activate
    rlastrow = pSheet.Range("A" & Rows.Count).End(xlUp).Row
    Range("A1:D" & rlastrow).Select
    pSheet.Sort.SortFields.Clear
    pSheet.Sort.SortFields.Add2 Key:=Range("B2:B" & rlastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    pSheet.Sort.SortFields.Add2 Key:=Range("C2:C" & rlastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With pSheet.Sort
        .SetRange Range("A1:D" & rlastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Filter
    Range("A1:D" & rlastrow).Select
    If pSheet.AutoFilterMode = True Then
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
Sub EquipListSystem(cOset)
    Dim wb As Workbook
    Dim mastepSheet As Worksheet
    Dim pSheet As Worksheet
    Dim ws As Worksheet
    Dim cupSheet As Worksheet
    Dim sumSheet As Worksheet
    

    'Setup references to workbook and sheet
    Set wb = ActiveWorkbook
    Set mastepSheet = wb.Worksheets("PROJECT_EQUIPMENT_LIST")
    Set pSheet = wb.Worksheets("Equipment Cost")
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
                    Set cupSheet = ActiveWorkbook.Worksheets(curName)
                    systemType = cupSheet.Range("A2").Value
                    LastRow = cupSheet.Range("A" & Rows.Count).End(xlUp).Row
                    'Look for ID on PROJECT_EQUIPMENT_LIST
                    Do Until itemRow > LastRow
                        If cupSheet.Range("A" & itemRow).Value <> "//" Then
                            If cupSheet.Range("A" & itemRow).Value <> "" Then
                                itemID = cupSheet.Range("A" & itemRow).Value
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
                                pSheet.Select
                                pSheet.Range("A2").Activate
                                pSheet.Cells.Find(What:=itemID, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                                rRow = ActiveCell.Row
                                
                                If Err.Number = 0 Then
                                    pSheet.Range("D" & rRow).Value = pSheet.Range("D" & rRow).Value + itemCount
                                    On Error GoTo 0
                                Else
                                    On Error GoTo 0
                                    'Add Make and Model
                                    mastepSheet.Select
                                    mastepSheet.Range("A1").Activate
                                    On Error Resume Next
                                    mastepSheet.Range("A:A").Find(What:=itemID, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
                                    idRow = ActiveCell.Row
                                    If idRow > 1 Then
                                        pSheet.Range("A" & i).Value = itemID
                                        pSheet.Range("B" & i).Value = mastepSheet.Range("B" & idRow).Value
                                        pSheet.Range("C" & i).Value = mastepSheet.Range("C" & idRow).Value
                                        pSheet.Range("D" & i).Value = itemCount
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


