Attribute VB_Name = "SMW_Tool_Checkers"
Sub toolcheck()
    Dim sheethere As Boolean
    Dim wb As Workbook: Set wb = ActiveWorkbook
    
    sheethere = WorksheetExists("DATA_HOLD")

    If sheethere = True Then
        If wb.Worksheets("DATA_HOLD").Range("AA1").Value = "TEMPLATE" Then
            AddIns("SMW-AV_EQL Tool").Installed = True
        Else
            MsgBox "This add-in is for use with the SMW AV Equipment List.  Please open the current SMW template to use this tool."
            End
        End If
    Else
        MsgBox "This add-in is for use with the SMW AV Equipment List"
        End
    End If
End Sub
Sub templateCheck()
    templateVersion = ActiveWorkbook.Worksheets("DATA_HOLD").Range("AB1").Value
    addinTemplateVersion = Workbooks("SMW-AV_EQL Tool.xlam").Worksheets("DATA").Range("B1").Value
    
    
    If templateVersion <> addinTemplateVersion Then
        MsgBox "This Equipment List template is not compatible with the currently installed SMW AV Tools.  Please updated your SMW AV Tools Add-in to version 3.0.0 from Teams."
        End
    End If
End Sub
Sub vCheck()
    Dim TheOS As String
    
    TheOS = Application.OperatingSystem
    
    If Val(Application.version) = 8 Then
        excelVer = "Excel 97"
    ElseIf Val(Application.version) = 9 Then
        excelVer = "Excel 2000"
    ElseIf Val(Application.version) = 10 Then
        excelVer = "Excel 2002"
    ElseIf Val(Application.version) = 11 Then
        excelVer = "Excel 2003"
    ElseIf Val(Application.version) = 12 Then
        excelVer = "Excel 2007"
    ElseIf Val(Application.version) = 14 Then
        excelVer = "Excel 2010"
    ElseIf Val(Application.version) = 15 Then
        excelVer = "Excel 2013"
    ElseIf Val(Application.version) = 16 Then
        excelVer = "Excel 2016"
    End If
    
    
    
    workbookVer = ActiveWorkbook.Worksheets("DATA_HOLD").Range("AB1").Value
       
    addinVer = ThisWorkbook.Worksheets("DATA").Range("B2").Value
    
    MsgBox "Operating System: " & TheOS & vbCrLf & "Excel Version: " & excelVer & vbCrLf & "Workbook Version: " & workbookVer & vbCrLf & "Add-In Version: " & addinVer
    
End Sub
