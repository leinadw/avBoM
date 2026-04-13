Attribute VB_Name = "Standard_Copy"
Public Sub CopySheetFromClosedWorkbook()
    Dim awName As Workbook: Set awName = ActiveWorkbook
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim sourceBook As Workbook
    
    fileName = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")
    
    'Get System List of Source Workbook
    GetSheetsNames (fileName)
    
    copySheetAsk.Show
    
    awNameFull = awName.Name
    
    
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
    
    'Copy Selected System
    Set sourceBook = Workbooks.Open(fileName)
    Workbooks.Open fileName:=fileName
    
    Windows(fileName).Activate
    ws.Select
    Cells.Replace What:="=", Replacement:="#", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
  
    sourceBook.Sheets(ws).Copy After:=awName.Sheets(awName.Sheets.Count)
    
    Cells.Replace What:="#", Replacement:="=", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    Unload copySheetAsk
    
    'Change links to activeworkbook
    ActiveWorkbook.ChangeLink Name:=fileName, NewName:=awNameFull, Type:=xlExcelLinks
    
    'Remove Room Numbers
    AciveWorkbook.Range("D2").Value = ""
    
    'Add missing equpment to ActiveWorkbook
    
    
    fileName.Close
    
    'Clean DATA_HOLD
    dataHold.Range("L:L").Clear
    
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
  
End Sub
Function GetSheetsNames(file)
Set sh = GetObject(file).Worksheets
'Setup Excluded Sheet Array
    Call excSetup
    
    For Each c In sh
        excArray = IsInArray(c.Name, ExcSheets)
        If excArray = False Then
            i = i + 1
            ActiveWorkbook.Worksheets("DATA_HOLD").Range("L" & i).Value = c.Name
        End If
    Next
End Function
