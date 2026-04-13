Attribute VB_Name = "Publish_EoPC"
Sub pubEST()
    Dim ws As Worksheet
    Dim WBName As Variant
    Dim estSheets() As Variant
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim iSheet As Worksheet: Set iSheet = ActiveWorkbook.Worksheets("Issuances")
    Dim sumSheet As Worksheet: Set sumSheet = ActiveWorkbook.Worksheets("Summary")
    Dim revSheet As Worksheet: Set revSheet = ActiveWorkbook.Worksheets("Revision List")
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")
    
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
    
    'Prompt for Notes Hide
    budgetAsk.Show
    
    If dataHold.Range("B1").Value = "" Then
        MsgBox "Please select systems to publish before you can contiunue."
        budgetAsk.Show
        If dataHold.Range("B1").Value = "" Then
            Exit Sub
        End If
    End If
    
    'Set pubSheet Array
    sysCount = dataHold.Range("B" & dataHold.Rows.Count).End(xlUp).Row
   
    ReDim estSheets(sysCount)
    ReDim pubSheets(sysCount)
    
    ac = 0
    For i = 1 To sysCount
        estSheets(ac) = dataHold.Range("B" & i).Value
        ac = ac + 1
    Next i
    ac = 0
    For i = 1 To sysCount
        pubSheets(ac) = dataHold.Range("B" & i).Value
        ac = ac + 1
    Next i
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
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
    
    'Set publish to EoPC
    EPC = False
    
    'Clean Worksheets
    For Each ws In Worksheets
        excArray = IsInArray(ws.Name, ExcSheets)
        If excArray = False Then
            estArray = IsInArray(ws.Name, estSheets)
            If estArray = True Then
                ws.Select
                Call CleanSheet(EPC)
            Else
                ws.Delete
            End If
        End If
    Next ws
    
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
    
    'Saving File
    If autofolder = True Then
        exType = "AV EoPC_"
        Call autofoldersave(BaseFolder, exType)
    Else
        WBName = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", InitialFileName:="AV EoPC.xlsx", Title:="Select Location for EoPC Save")
        If WBName <> False Then
            ActiveWorkbook.SaveAs fileName:=WBName, FileFormat:=51, ConflictResolution:=xlLocalSessionChanges
        End If
    End If
    
    If budgetAsk.OptionButton4 = True Then
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
    
    Unload budgetAsk
    Unload revAsk
    
    OneDriveYes = InStr(ActiveWorkbook.FullName, "http")
    If OneDriveYes > 0 Then
        ActiveWorkbook.AutoSaveOn = False
    End If
    
    Application.CutCopyMode = False
    
    Application.DisplayAlerts = True

End Sub
