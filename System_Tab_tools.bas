Attribute VB_Name = "System_Tab_tools"
Sub newSys()
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim STL As Worksheet: Set STL = ActiveWorkbook.Worksheets("SYSTEM_TEMPLATE_LOOKUP")
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")

    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    'Get Worksheet List
    Call sheetList
    
    'Check for template sheet visibility
    If STL.Visible = xlSheetVeryHidden Then
        STL.Visible = xlSheetVisible
        STLvis = False
    Else
        STLvis = True
    End If
    
    'Show userform
    SysAsk.Show
    
    'Show system name dialog box
    sysName.Show
    
    'Show hidden systems
    If psSheet.Range("N3").Value <> True Then
        SetWsVisibility 1, 5
    End If
    
    If sysName.TextBox1.Value = "" Then
        End
    End If
    
    'Check name not used
    imHere = WorksheetExists(sysName.TextBox1.Value)
    If imHere = True Then
        Worksheets(sysName.TextBox1.Value).Visible = xlSheetVisible
        MsgBox sysName.TextBox1.Value & " is already in use.  Pick an unused name."
        sysName.Show
        If sysName.TextBox1.Value = "" Then
            End
        End If
        imHere = WorksheetExists(sysName.TextBox1.Value)
        If imHere = True Then
            End
        End If
    End If
    'Make sheet
    If SysAsk.OptionButton4.Value = True Then
        Sheets("SYSTEM_TEMPLATE_LOOKUP").Copy After:=Sheets(Sheets.Count)
    ElseIf SysAsk.OptionButton3.Value = True Then
        If SysAsk.ComboBox1.Value = "" Then
            MsgBox "You must select a system to copy."
            Unload sysName
            Unload SysAsk
            End
        End If
        Sheets(SysAsk.ComboBox1.Value).Copy After:=Sheets(Sheets.Count)
    End If
    
    'Set worksheet name
    ActiveSheet.Name = sysName.TextBox1.Value
    NewSheet = sysName.TextBox1.Value

        
    'Unload forms
    Unload sysName
    Unload SysAsk
    
    'Update Summary Sheet
    Call sumSheetSet
    
    'Clear temp files
    dataHold.Range("A:A").Clear
    
    'Select New Sheet
    ActiveWorkbook.Worksheets(NewSheet).Activate
    ActiveWorkbook.Worksheets(NewSheet).Range("D2").ClearContents
    
    Unload SysAsk
    Unload sysName
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
Sub deleteSys()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    'Get Worksheet List
    Call sheetList
    
    'Show userform
    sysDelete.Show
  
      
    dSheet = sysDelete.ComboBox1.Value
    
    ActiveWorkbook.Worksheets(dSheet).Delete

    Unload sysDelete
    
    'Update Summary Sheet
    Call sumSheetSet
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
End Sub
