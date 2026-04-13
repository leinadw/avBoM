Attribute VB_Name = "DWG_Tools"
Sub dwgPull()
    Dim myRange As Range
    Dim LastRow As Long
    Dim folderName As String
    Dim strFolderExists As String
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim tempSheet As Worksheet: Set tempSheet = wb.Worksheets("_TEMP")
    Dim reportSheet As Worksheet: Set reportSheet = wb.Worksheets("DWG Report")
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim dwgWb As Workbook
    Dim dwgSheet As Worksheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    'Get worksheet List
    Call sheetList
    
    'Select AV Systems
    dwgAskDW.Show
    
    'Set dwgSheet Array
    sysCount = dataHold.Range("B" & dataHold.Rows.Count).End(xlUp).Row
    
    ReDim dwgSheets(sysCount)
    
    ac = 0
    For i = 1 To sysCount
        dwgSheets(ac) = dataHold.Range("B" & i).Value
        ac = ac + 1
    Next i
    
    'Select Project Log
    MsgBox "Select Location of this projects dwg extract"
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Show
        On Error Resume Next
        folderName = .SelectedItems(1)
        DriveLetter = Left(folderName, 2)
        If DriveLetter <> "C:" Then
            dwgLoc = GETNETWORKPATH(DriveLetter) & Right(folderName, Len(folderName) - 2)
        Else
            dwgLoc = folderName
        End If
        dwgFile = Right(folderName, Len(folderName) - InStrRev(folderName, "\"))
        projFolder = Left(dwgLoc, Len(dwgLoc) - (Len(dwgLoc) - InStrRev(dwgLoc, "\", -1, vbBinaryCompare)))
        Err.Clear
        On Error GoTo 0
    End With
    
    'Open Log Workbook
    Workbooks.Open dwgLoc
    Set dwgWb = Workbooks(dwgFile)
    Set dwgSheet = dwgWb.Worksheets("Summary")
    
    'Import Data
    dwgWb.Activate
    dwgSheet.Select
    dwglastrow = dwgSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    
    'filter range where column 2 in range is equal to delete word
    dwgSheet.Range("A1:D" & dwglastrow).AutoFilter Field:=2, Criteria1:=Array( _
        "SMW-AV", "SMW-AV_BORDER", "SMW-AV_C", "SMW-AV_C_FRAME", "SMW-AV_C_INFRA", _
        "SMW-AV_C_SPK", "SMW-AV_FRAME", "SMW-AV_INFRA", "SMW-AV_TAG_LEFT", _
        "SMW-AV_TAG_RIGHT", "SMW-AV_WIRE_FLAG"), Operator:=xlFilterValues
  
    'delete rows that are visible
    dwgSheet.Range("A2:D" & dwglastrow).SpecialCells(xlCellTypeVisible).Delete
  
    'remove filter
    On Error Resume Next
    dwgSheet.ShowAllData
    On Error GoTo 0
    
    
    dwgSheet.Range("A2:D" & dwglastrow).Select
    Selection.Copy
    wb.Activate
    '_TEMP Visible
    tempSheet.Visible = xlSheetVisible
    tempSheet.Activate
    tempSheet.Range("A1").Select
    ActiveSheet.Paste
    dwgWb.Close SaveChanges:=False
    
    'Clean dwgsheet Report
    LastRowA = reportSheet.Range("A" & Rows.Count).End(xlUp).Row
    LastRowD = reportSheet.Range("D" & Rows.Count).End(xlUp).Row
    If LastRowA > 1 And LastRowD Then
        reportSheet.Range("A3:D" & LastRowA).ClearContents
    ElseIf LastRowD > 1 Then
        reportSheet.Range("A2:D" & LastRowD).ClearContents
    End If
    
    rNum = 3
    
    For Each ws In ActiveWorkbook.Worksheets
        excArray = IsInArray(ws.Name, ExcSheets)
        If excArray = False Then
            dwgArray = IsInArray(ws.Name, dwgSheets)
            If dwgArray = True Then
                ws.Select
                Call partIDExtract
            End If
        End If
    Next ws
    
    'Move Systems in DWG not in BoM to report
    templastrow = tempSheet.Range("A" & Rows.Count).End(xlUp).Row
    For n = 1 To templastrow
            reportSheet.Range("A" & rNum).Value = tempSheet.Range("D" & n).Value
            reportSheet.Range("B" & rNum).Value = tempSheet.Range("C" & n).Value
            reportSheet.Range("D" & rNum).Value = tempSheet.Range("A" & n).Value
            tempSheet.Range("A" & n & ":D" & n).Select
            Selection.Clear
            rNum = rNum + 1
    Next n
    
    
    'Clean Up _TEMP and hide
    LastRow = tempSheet.Range("A" & Rows.Count).End(xlUp).Row
    tempSheet.Range("A1:S" & LastRow).ClearContents
    tempSheet.Visible = xlSheetVeryHidden
    
    reportSheet.Activate
    'delete empty rows
    reportlastrow = reportSheet.Range("A" & Rows.Count).End(xlUp).Row
    reportSheet.Range("A2:D" & reportlastrow).SpecialCells(xlCellTypeBlanks).Delete
    
    'highlight mismatch
    Cells.FormatConditions.Delete
    reportSheet.Range("C2:D" & reportlastrow).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$C2<>$D2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    
    reportSheet.Visible = xlSheetVisible
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
Sub partIDExtract()
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Dim endCell As Range: Set endCell = asName.Cells.Find(What:="//", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim reportSheet As Worksheet: Set reportSheet = wb.Worksheets("DWG Report")
    Dim tempSheet As Worksheet: Set tempSheet = wb.Worksheets("_TEMP")
    
    sysType = asName.Range("A2").Value
    
    For i = 6 To endCell.Row
        mfgName = asName.Range("C" & i).Value
        If mfgName <> "" Then
            'Add BoM item to Report
            reportSheet.Range("A" & rNum).Value = asName.Range("A2").Value
            reportSheet.Range("B" & rNum).Value = asName.Range("A" & i).Value
            reportSheet.Range("C" & rNum).Value = asName.Range("F" & i).Value
            
            
            'Add DWG item to Report
            tempSheet.Select
            tempSheet.Range("A1").Select
            templastrow = tempSheet.Range("A" & Rows.Count).End(xlUp).Row
            modType = reportSheet.Range("B" & rNum).Value
        
            For n = 1 To templastrow
                If Cells(n, "D") = sysType And Cells(n, "C") = modType Then
                    reportSheet.Range("D" & rNum).Value = tempSheet.Range("A" & n).Value
                    tempSheet.Range("A" & n & ":D" & n).Select
                    Selection.Clear
                    Exit For
                Else
                    reportSheet.Range("D" & rNum).Value = 0
                End If
            Next n
            rNum = rNum + 1
        End If
    Next i
    
    'items in dwg not in bom
    For n = 1 To templastrow
        If Cells(n, "C") = sysType Then
            reportSheet.Range("A" & rNum).Value = asName.Range("A2").Value
            reportSheet.Range("B" & rNum).Value = tempSheet.Range("B" & n).Value
            reportSheet.Range("C" & rNum).Value = 0
            reportSheet.Range("D" & rNum).Value = tempSheet.Range("A" & n).Value
            tempSheet.Range("A" & n & ":D" & n).Select
            Selection.Clear
            rNum = rNum + 1
        End If
    Next n
End Sub
