Attribute VB_Name = "Persistant_Var_Tools"
Public noteStat As String
Public costStat As String
Public equipStat As String
Public laborStat As String
Public rNum As String
Public importGo As Boolean
Public issueC As Boolean
Public cutName As String
Public logLoc As String
Public logFile As String
Public projFolder As String
Public pw As String
Public EPC As Boolean
Public ExcSheets() As Variant
Public pubSheets() As Variant
Public cutSheets() As Variant
Public dbMasterFile As Workbook
Private Const mcGWL_STYLE = (-16)
Private Const mcWS_SYSMENU = &H80000




'Windows API calls to handle windows
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Public Sub pwUP()
    pw = 6891
End Sub

Public Sub subRemoveCloseButton(frm As Object)
    Dim lngStyle As Long
    Dim lngHWnd As Long

    lngHWnd = FindWindow(vbNullString, frm.Caption)
    lngStyle = GetWindowLong(lngHWnd, mcGWL_STYLE)

    If lngStyle And mcWS_SYSMENU > 0 Then
        SetWindowLong lngHWnd, mcGWL_STYLE, (lngStyle And Not mcWS_SYSMENU)
    End If

End Sub
Public Sub excSetup()
    
    ExcSheets = Array("Summary", "SYSTEM_TEMPLATE_LOOKUP", "DATA_HOLD", "PROJECT_EQUIPMENT_LIST", "PROJECT_SETTINGS", "INSTRUCTIONS", "Issuances", "Revision List", "_TEMP", "Equipment Report", "DWG Report", "Cutsheet Report", "Equipment Cost")

End Sub
Public Sub sheetList()
    Dim ws As Worksheet
    Dim X As Integer
    Dim i As Long
    Dim r As Long
    Dim text As Range
    Dim rng As Range
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    
    X = 1
    
    dataHold.Range("A:B").Clear
    
    For Each ws In Worksheets
        If Sheets(ws.Name).Visible = True Then
            dataHold.Cells(X, 1) = ws.Name
            X = X + 1
        End If
    Next ws
    
    LastRow = dataHold.Range("A" & dataHold.Rows.Count).End(xlUp).Row
    
    For i = 1 To LastRow
        Set txt = dataHold.Cells(i, "A")
        For r = 0 To UBound(ExcSheets)
            txt.Formula = Replace(txt.Formula, ExcSheets(r), "")
        Next r
    Next i
    
    Set rng = dataHold.Range("A1:A" & LastRow + 1).SpecialCells(xlCellTypeBlanks)
    rng.Rows.Delete Shift:=xlShiftUp
End Sub
Public Sub RemoveAllMacros(objDocument As Object)
' deletes all VBProject components from objDocument
' removes the code from built-in components that can't be deleted
' use like this: RemoveAllMacros ActiveWorkbook ' in Excel
' or like this: RemoveAllMacros ActiveWorkbookDocument ' in Word
' requires a reference to the
' Microsoft Visual Basic for Applications Extensibility library
Dim i As Long, l As Long
    If objDocument Is Nothing Then Exit Sub
    i = 0
    On Error Resume Next
    i = objDocument.VBProject.VBComponents.Count
    On Error GoTo 0
    If i < 1 Then ' no VBComponents or protected VBProject
        MsgBox "The VBProject in " & objDocument.Name & _
            " is protected or has no components!", _
            vbInformation, "Remove All Macros"
        Exit Sub
    End If
    With objDocument.VBProject
        For i = .VBComponents.Count To 1 Step -1
            On Error Resume Next
            .VBComponents.Remove .VBComponents(i)
            ' delete the component
            On Error GoTo 0
        Next i
    End With
    With objDocument.VBProject
        For i = .VBComponents.Count To 1 Step -1
            l = 1
            On Error Resume Next
            l = .VBComponents(i).CodeModule.CountOfLines
            .VBComponents(i).CodeModule.DeleteLines 1, l
            ' clear lines
            On Error GoTo 0
        Next i
    End With
End Sub
Public Sub SetWsVisibility(Optional ByVal vis As Boolean = True, _
                           Optional ByVal visibleWs As Long = 0)

    Static vSet As Boolean, hSet As Boolean, wsCount As Long, lastV As Long, i As Long
    
    'Setup Excluded Sheet Array
    Call excSetup

    With ActiveWorkbook

        wsCount = .Worksheets.Count - 1

        'if visibleWs is 0 last ws is visible, or use any other valid sheet index
        visibleWs = IIf(visibleWs < 1 Or visibleWs > wsCount, wsCount + 1, visibleWs)

        If wsCount <> .Worksheets.Count - 1 Or visibleWs <> lastV Then
            vSet = False
            hSet = False
        Else
            If vSet And vis Then .CustomViews("ShowAllWs").Show:        Exit Sub
            If hSet And Not vis Then .CustomViews("HideAllWs").Show:    Exit Sub
        End If

        If vis Then
            For i = 1 To wsCount + 1
            excArray = IsInArray(Worksheets(i).Name, ExcSheets)
            If excArray = False Then
                With .Worksheets(i)
                    If Not .Visible Then .Visible = vis
                End With
            End If
            Next
            .Worksheets(1).Activate
            .CustomViews.Add ViewName:="ShowAllWs"  'Save View (one-time operation)
            vSet = True
        Else
            If visibleWs <> lastV Then
                For i = 1 To wsCount + 1
                    excArray = IsInArray(ActiveSheet.Name, ExcSheets)
                    If excArray = False Then
                        With .Worksheets(i)
                            If Not .Visible Then .Visible = 1
                        End With
                    End If
                Next
            End If

            Dim arr() As Variant, j As Long
            ReDim arr(1 To wsCount)
            j = 1
            For i = 1 To wsCount + 1
                If i <> visibleWs Then arr(j) = i Else j = j - 1
                j = j + 1
            Next
            .Worksheets(arr).Visible = vis
            .CustomViews.Add ViewName:="HideAllWs"  'Save View (one-time operation)
            hSet = True
            lastV = visibleWs
        End If
    End With
End Sub
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
Public Function GETNETWORKPATH(ByVal DriveName As String) As String
     
     Dim objNtWork  As Object
     Dim objDrives  As Object
     Dim lngLoop    As Long
     
     
     Set objNtWork = CreateObject("WScript.Network")
     Set objDrives = objNtWork.enumnetworkdrives
     
     For lngLoop = 0 To objDrives.Count - 1 Step 2
         If UCase(objDrives.Item(lngLoop)) = UCase(DriveName) Then
             GETNETWORKPATH = objDrives.Item(lngLoop + 1)
             Exit For
         End If
     Next

 End Function
Public Function WorksheetExists(ByVal WorksheetName As String) As Boolean
    Dim sht As Worksheet

    For Each sht In ActiveWorkbook.Worksheets
        If Application.Proper(sht.Name) = Application.Proper(WorksheetName) Then
            WorksheetExists = True
            Exit Function
        End If
    Next sht
    WorksheetExists = False
End Function
Public Function GetLocalFile(wb As Workbook) As String
    Dim OneDriveYes As Integer
    ' Set default return
    GetLocalFile = wb.FullName
    OneDriveYes = InStr(GetLocalFile, "http")
    If OneDriveYes > 0 Then
        MsgBox "This file is save in OneDrive and is not compatable with this Macro.  Please save locally or in the project folder."
        GetLocalFile = "ODyes"
        Exit Function
    End If
'    MsgBox GetLocalFile

'    Const HKEY_CURRENT_USER = &H80000001
'
'    Dim strValue As String
'
'    Dim objReg As Object: Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
'    Dim strRegPath As String: strRegPath = "Software\SyncEngines\Providers\OneDrive\"
'    Dim arrSubKeys() As Variant
'    objReg.EnumKey HKEY_CURRENT_USER, strRegPath, arrSubKeys
'
'    Dim varKey As Variant
'    For Each varKey In arrSubKeys
'        ' check if this key has a value named "UrlNamespace", and save the value to strValue
'        objReg.getStringValue HKEY_CURRENT_USER, strRegPath & varKey, "UrlNamespace", strValue
'
'        ' If the namespace is in FullName, then we know we have a URL and need to get the path on disk
'        If InStr(wb.FullName, strValue) > 0 Then
'            Dim strTemp As String
'            Dim strCID As String
'            Dim strMountpoint As String
'
'            ' Get the mount point for OneDrive
'            objReg.getStringValue HKEY_CURRENT_USER, strRegPath & varKey, "MountPoint", strMountpoint
'
'            ' Get the CID
'            objReg.getStringValue HKEY_CURRENT_USER, strRegPath & varKey, "CID", strCID
'
'            ' strip off the namespace and CID
'            If Len(strCID) > 0 Then strValue = strValue & "/" & strCID     '#####
'            strTemp = Right(wb.FullName, Len(wb.FullName) - Len(strValue)) '#####
'
'            ' replace all forward slashes with backslashes
'            GetLocalFile = strMountpoint & "\" & Replace(strTemp, "/", "\")
'            Exit Function
'        End If
'    Next
End Function
Sub autofoldercheck()

    'Base folder set
    aFile = GetLocalPath(ActiveWorkbook.FullName)
    If aFile = "ODyes" Then
        Exit Sub
    End If
    BaseFolder = Left(aFile, InStrRev(aFile, "\") - 1)
    
    'Check for Archive Folder and create if not there
    strFolderExists = Dir(BaseFolder & "\Archive", vbDirectory)
    If strFolderExists = "" Then
        MkDir BaseFolder & "\Archive"
    End If
    
    'Check for Issued Folder and create if not there
    strFolderExists = Dir(BaseFolder & "\Issued", vbDirectory)
    If strFolderExists = "" Then
        MkDir BaseFolder & "\Issued"
    End If
    
    'Save Current book
    If revAsk.ComboBox1.Value <> "" Then
        If revAsk.ComboBox1.Value = "Add Issuance" Then
            iName = revAsk.TextBox1.Value
        Else
            iName = revAsk.ComboBox1.Value
        End If
    Else
        iName = archName.TextBox1.Value
    End If
    shortName = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 12)
    
    textlen = Len(BaseFolder & "\Archive\" & shortName & iName & ".xlsx")
    If textlen > 255 Then
       MsgBox "Your system name must not exceed 31 character and not include any of the following characters: : \ / ?.  Please enter the name again."
    End If
    
    ActiveWorkbook.SaveCopyAs fileName:=BaseFolder & "\Archive\" & shortName & iName & ".xlsx"
'    ActiveWorkbook.SaveAs BaseFolder & "\Archive\" & shortName & iName & ".xlsx", FileFormat:=51

    If revAsk.ComboBox1.Value <> "" Then
        EPC = True
    End If
    
    OneDriveYes = InStr(ActiveWorkbook.FullName, "http")
    If OneDriveYes > 0 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Workbooks.Open BaseFolder & "\Archive\" & shortName & iName & ".xlsx", UpdateLinks:=False
        ActiveWorkbook.AutoSaveOn = False
        Workbooks(shortName & iName & ".xlsx").Close
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If

End Sub
Sub autofoldersave(BaseFolder, exType)

    Dim fileSaver As FileDialog
    Set fileSaver = Application.FileDialog(msoFileDialogSaveAs)
    If revAsk.ComboBox1.Value = "Add Issuance" Then
        iName = revAsk.TextBox1.Value
    Else
        iName = revAsk.ComboBox1.Value
    End If


    textlen = Len(BaseFolder & "\Issued\" & exType & iName & ".xlsx")
    If textlen > 255 Then
       MsgBox "Your system name must not exceed 31 character and not include any of the following characters: : \ / ?.  Please enter the name again."
    End If

    WBName = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", InitialFileName:=BaseFolder & "\Issued\" & exType & iName & ".xlsx")
    If WBName <> False Then
        ActiveWorkbook.SaveAs fileName:=WBName, FileFormat:=51, ConflictResolution:=xlLocalSessionChanges
    Else
        WBName = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", InitialFileName:=BaseFolder & "\Issued\" & exType & iName & ".xlsx")
        If WBName <> False Then
            ActiveWorkbook.SaveAs fileName:=WBName, FileFormat:=51, ConflictResolution:=xlLocalSessionChanges
        End If
    End If
    
End Sub
Private Sub contActive_Click()
    textlen = Len(TextBox1.text)
    If textlen > 255 Then
       MsgBox "Your system name must not exceed 31 character and not include any of the following characters: : \ / ?.  Please enter the name again."
    End If
End Sub
Public Function ValidFileName(fileName As String) As Boolean
'PURPOSE: Determine If A Given Excel File Name Is Valid

Const sBadChar As String = "\/:*?<>|[]"""
Dim i As Long

'Assume valid unless it isn't
  ValidFileName = True

'Loop through each "Bad Character" and test for an instance
  For i = 1 To Len(sBadChar)
    If InStr(fileName, Mid$(sBadChar, i, 1)) > 0 Then
      ValidFileName = False 'Invalid
      Exit For
    End If
  Next
  
End Function
Public Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
