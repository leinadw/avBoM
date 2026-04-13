Attribute VB_Name = "Issuance_Revision_Updater"
Sub revUp()
    Dim ws As Worksheet
    Dim revList As Worksheet: Set revList = ActiveWorkbook.Worksheets("Revision List")
    Dim iSheet As Worksheet: Set iSheet = ActiveWorkbook.Worksheets("Issuances")
    
    revAsk.Show
    
    iName = revAsk.ComboBox1.Value
    issueC = False
    
    'Add if new
    If iName = "Add Issuance" Then
        iLastRow = iSheet.Range("A" & Rows.Count).End(xlUp).Row
        If iLastRow < 5 Then
            iLastRow = 5
        Else
            iLastRow = iLastRow + 1
        End If
'        revName.Show
        iName = revAsk.TextBox1.Value
        iSheet.Range("A" & iLastRow).Value = iName
        issueC = True
    End If
    
    'Find Issuance Line
    Do Until issueC = True
        iSheet.Activate
        iRow = iSheet.Cells.Find(What:=iName, After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
        
        'See if already issued
        If iSheet.Range("C" & iRow).Value <> "" Then
            issueCheck.Show
        Else
            issueC = True
        End If
    Loop
    
    For Each ws In Worksheets
        excArray = IsInArray(ws.Name, ExcSheets)
        If excArray = False Then
            pubArray = IsInArray(ws.Name, pubSheets)
            If pubArray = True Then
                ws.Activate
                Call NameRevUp(iName)
            End If
        End If
    Next
    
'    Unload revAsk
    
    LastRow = revList.Range("A" & Rows.Count).End(xlUp).Row
    revList.PageSetup.PrintArea = "$A$1:$H$" & LastRow
    
    
End Sub
Sub NameRevUp(iName)
    Dim i As Long
    Dim r As Long
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim revList As Worksheet: Set revList = ActiveWorkbook.Worksheets("Revision List")
    Dim iSheet As Worksheet: Set iSheet = ActiveWorkbook.Worksheets("Issuances")
    
    'Find Issuance Line
    iSheet.Activate
    iRow = iSheet.Cells.Find(What:=iName, After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
    
    'Find EndCell
    asName.Activate
    p = 5
    Do While endCellfound = False
        asName.Range("A" & p).Select
        cellColor = ActiveCell.Interior.Color
        If cellColor = 14277081 Or cellColor = 13288897 Then
            endCellfound = True
            endCellrow = p
        End If
        p = p + 1
    Loop
    
    'Update Last Issued and Date
    asName.Range("A3").Value = iName
    asName.Range("I1").Value = iSheet.Range("B" & iRow).Value
    
    
    r = 6
    Do Until r = endCellrow
        If asName.Range("AE" & r).Value = iName Then
            'Find Revision Last Row
            revList.Activate
            dblastRow = revList.Range("A" & Rows.Count).End(xlUp).Row
            If dblastRow = 4 Then
                dblastRow = dblastRow + 1
            End If
            
            'Add name and info to Revision
            revList.Range("A" & dblastRow + 1).Value = asName.Range("A2").Value
            revList.Range("B" & dblastRow + 1).Value = asName.Range("AE" & r).Value
            revList.Range("C" & dblastRow + 1).Value = asName.Range("AF" & r).Value
            revList.Range("D" & dblastRow + 1).Value = asName.Range("AG" & r).Value
            revList.Range("E" & dblastRow + 1).Value = asName.Range("AH" & r).Value
            revList.Range("F" & dblastRow + 1).Value = asName.Range("AI" & r).Value
            revList.Range("G" & dblastRow + 1).Value = asName.Range("AJ" & r).Value
            revList.Range("H" & dblastRow + 1).Value = asName.Range("AK" & r).Value
            
        End If
        r = r + 1
    Loop
    
    'Update System List on Issuance Page
    If iSheet.Range("C" & iRow).Value <> "" Then
        iSheet.Range("C" & iRow).Value = iSheet.Range("C" & iRow).Value & ", " & asName.Range("A2").Value
    Else
        iSheet.Activate
        With iSheet.Range("B" & iRow)
            .Value = revAsk.DateText.Value
        End With
        iSheet.Range("C" & iRow).Value = asName.Range("A2").Value
    End If
    
End Sub
Sub bbRev()
    Dim ws As Worksheet
    Dim bbSheets() As Variant
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim revList As Worksheet: Set revList = ActiveWorkbook.Worksheets("Revision List")
    Dim iSheet As Worksheet: Set iSheet = ActiveWorkbook.Worksheets("Issuances")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Setup Excluded Sheet Array
    Call excSetup
    
    bbAsk.Show
    
    'Set pubSheet Array
    sysCount = dataHold.Range("B" & dataHold.Rows.Count).End(xlUp).Row
   
    ReDim bbSheets(sysCount)
    
    ac = 0
    For i = 1 To sysCount
        bbSheets(ac) = dataHold.Range("B" & i).Value
        ac = ac + 1
    Next i
    
    'Get Issuance Name
    iName = bbAsk.ComboBox1.Value
    iSheet.Activate
    iRow = iSheet.Cells.Find(What:=iName, After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row

    For Each ws In Worksheets
        excArray = IsInArray(ws.Name, ExcSheets)
        If excArray = False Then
            bbArray = IsInArray(ws.Name, bbSheets)
            If bbArray = True Then
                ws.Activate
                Call bbRevUp(iName)
            End If
        End If
    Next
    
    'Hide Budget Issuance
    If bbAsk.OptionButton1.Value = True Then
        iSheet.Activate
        iSheet.Range("A5:C" & iRow - 1).Select
        Selection.EntireRow.Hidden = True
    End If
    
    Unload bbAsk
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
Sub bbRevUp(iName)
    Dim i As Long
    Dim r As Long
    Dim rowNum As String
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    Dim revList As Worksheet: Set revList = ActiveWorkbook.Worksheets("Revision List")
    Dim iSheet As Worksheet: Set iSheet = ActiveWorkbook.Worksheets("Issuances")
    
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
            endCellrow = i
        End If
        i = i + 1
    Loop
    
    r = 6
    rowNum = r
    Do Until r = endCellrow
        sectionBreak = IsInArray(rowNum, sec)
        If sectionBreak = True Then
            r = r + 1
            rowNum = r
        Else
            asName.Range("AE" & r).Value = iName
            asName.Range("AH" & r).ClearContents
            If IsEmpty(asName.Range("F" & r).Value) = False And asName.Range("F" & r).Value = 0 Then
                Rows(r).Select
                Selection.EntireRow.ClearContents
                asName.Range("A" & r).Value = "DELETE"
            End If
            r = r + 1
            rowNum = r
        End If
    Loop
    
    r = 6
    Do Until r = endCellrow
        If asName.Range("A" & r).Value = "DELETE" Then
            Rows(r).Select
            Selection.EntireRow.Delete
        End If
        r = r + 1
    Loop
End Sub
