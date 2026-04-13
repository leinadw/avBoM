Attribute VB_Name = "Summary_Sheet"
Sub sumSheetSet()
    Dim ws As Worksheet
    Dim sumSheet As Worksheet: Set sumSheet = ActiveWorkbook.Worksheets("Summary")
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")
    
    Application.ScreenUpdating = False
    
    sumSheet.Activate
    On Error Resume Next
    Set LastRow = sumSheet.Cells.Find(What:="TOTAL EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If LastRow = "" Then
        Set LastRow = sumSheet.Cells.Find(What:="TOTAL EQUIPMENT & NON-EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    End If
    On Error GoTo 0

    sumRow = 5
          
    
    'Setup Summary Per Project
    rOset = 0
    cOset = 0
    eneSub = psSheet.Range("P6").Value
    licSub = psSheet.Range("P3").Value
    supSub = psSheet.Range("P9").Value
    
    'Combine Equipment and Non Equipment
    If eneSub = True Then
        'Add SubTotal Column
        On Error Resume Next
        Set enesubCol = sumSheet.Cells.Find(What:="EQUIPMENT & NON-EQUIPMENT SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
        If enesubCol = "" Then
            sumSheet.Activate
            Columns(7 + cOset).Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.Insert Shift:=xlRight
            sumSheet.Cells(4, 8 + cOset).Value = "EQUIPMENT & NON-EQUIPMENT SUBTOTAL"
            enesubColnum = 8 + cOset
            'Addd Subtotal formula and Fix System Subtotal Formula
            enesubLetter = Col_Letter(8 + cOset)
            sumSheet.Range(enesubLetter & "5").Formula = "=SUM($F5:$G5)"
            sumSheet.Range(enesubLetter & "5").Select
            Application.CutCopyMode = False
            Selection.Copy
            sumSheet.Range(enesubLetter & "6:" & enesubLetter & "7").Select
            ActiveSheet.Paste
            cOset = cOset + 1
            enetotalLetter = Col_Letter(10 + cOset)
            conSubLetter = Col_Letter(9 + cOset)
            sumSheet.Range(enetotalLetter & "5").Formula = "=SUM(" & enesubLetter & "5," & conSubLetter & "5)"
            sumSheet.Range(enetotalLetter & "5").Select
            Application.CutCopyMode = False
            Selection.Copy
            sumSheet.Range(enetotalLetter & "6:" & enetotalLetter & "7").Select
            ActiveSheet.Paste
        Else
            licCol = 8 + cOset
            cOset = cOset + 1
        End If
        On Error GoTo 0
        
        'Remove Non-Equipment Total Row and update formula
        On Error Resume Next
        Set nequiprow = sumSheet.Cells.Find(What:="TOTAL NON-EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
        If Not nequiprow Is Nothing Then
            sumSheet.Cells(LastRow.Row, 11 + cOset).Value = "TOTAL EQUIPMENT & NON-EQUIPMENT COST SUBTOTAL"
            sumSheet.Range(LastRow.Row + 2 + rOset & ":" & LastRow.Row + 3 + rOset).Select
            Selection.Delete
            rOset = rOset - 2
            qtyLetter = Col_Letter(11 + cOset)
            sumSheet.Cells(LastRow.Row, 12 + cOset).Formula = "=SUMPRODUCT(" & enesubLetter & "5:" & enesubLetter & "7," & qtyLetter & "5:" & qtyLetter & "7)"
            lastRowColLetter = Col_Letter(12 + cOset)
            sumSheet.Cells(LastRow.Row + 6 + rOset, 12 + cOset).Formula = "=SUM(" & lastRowColLetter & LastRow.Row & ":" & lastRowColLetter & LastRow.Row + 5 + rOset & ")"
 
        Else
            rOset = rOset - 2
        End If
    End If
    
    'License as seperate sub total
    If licSub = True Then
        'Add License Column
        On Error Resume Next
        Set licCol = sumSheet.Cells.Find(What:="LICENSE", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
        If licCol = "" Then
            sumSheet.Activate
            Columns(7 + cOset).Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.Insert Shift:=xlRight
            sumSheet.Cells(4, 8 + cOset).Value = "LICENSE"
            cOset = cOset + 1
            If eneSub = True Then
                enetotalLetter = Col_Letter(10 + cOset)
                conSubLetter = Col_Letter(9 + cOset)
                licsubLetter = Col_Letter(7 + cOset)
                qtyLetter = Col_Letter(11 + cOset)
                sumSheet.Range(enetotalLetter & "5").Formula = "=SUM(" & enesubLetter & "5," & licsubLetter & "5," & conSubLetter & "5)"
                sumSheet.Range(enetotalLetter & "5").Select
                Application.CutCopyMode = False
                Selection.Copy
                sumSheet.Range(enetotalLetter & "6:" & enetotalLetter & "7").Select
                ActiveSheet.Paste
                sumSheet.Cells(LastRow.Row, 12 + cOset).Formula = "=SUMPRODUCT(" & enesubLetter & "5:" & enesubLetter & "7," & qtyLetter & "5:" & qtyLetter & "7)"
            Else
                enetotalLetter = Col_Letter(10 + cOset)
                licsubLetter = Col_Letter(7 + cOset)
                qtyLetter = Col_Letter(11 + cOset)
                sumSheet.Range(enetotalLetter & "5").Formula = "=SUM(F5,G5,H5,J5)"
                sumSheet.Range(enetotalLetter & "5").Select
                Application.CutCopyMode = False
                Selection.Copy
                sumSheet.Range(enetotalLetter & "6:" & enetotalLetter & "7").Select
                ActiveSheet.Paste
                sumSheet.Cells(LastRow.Row + 2, 12 + cOset).Formula = "=SUMPRODUCT(G5:G7,L5:L7)"
            End If
        Else
            cOset = cOset + 1
            licCol = 8
        End If
        On Error GoTo 0

        'Add License Total Row
        On Error Resume Next
        Set licrow = sumSheet.Cells.Find(What:="TOTAL LICENSE COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
        If licrow = "" Then
            sumSheet.Activate
            Rows(LastRow.Row & ":" & LastRow.Row + 1).Select
            Application.CutCopyMode = False
            Selection.Copy
            Rows(LastRow.Row + 4 + rOset & ":" & LastRow.Row + 4 + rOset).Select
            Selection.Insert Shift:=xlDown
            qtyLetter = Col_Letter(11 + cOset)
            sumSheet.Cells(LastRow.Row + 4 + rOset, 11 + cOset).Value = "TOTAL LICENSE COST SUBTOTAL"
            sumSheet.Cells(LastRow.Row + 4 + rOset, 12 + cOset).Formula = "=SUMPRODUCT(" & licsubLetter & "5:" & licsubLetter & "7," & qtyLetter & "5:" & qtyLetter & "7)"
            rOset = rOset + 2
        Else
            rOset = rOset + 2
        End If
        On Error GoTo 0
    End If

    'Support Add in Totals
    If supSub = True Then
        On Error Resume Next
        Set suprow = sumSheet.Cells.Find(What:="TOTAL WARRANTY COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
        If suprow = "" Then
            sumSheet.Activate
            Rows(LastRow.Row + 4 + rOset & ":" & LastRow.Row + 5 + rOset).Select
            Application.CutCopyMode = False
            Selection.Copy
            Rows(LastRow.Row + 4 + rOset & ":" & LastRow.Row + 4 + rOset).Select
            Selection.Insert Shift:=xlUp
            sumSheet.Cells(LastRow.Row + 4 + rOset, 11 + cOset).Value = "TOTAL WARRANTY COST SUBTOTAL"
            sumSheet.Cells(LastRow.Row + 4 + rOset, 12 + cOset).Value = 0
            rOset = rOset + 2
        Else
            rOset = rOset + 1
        End If
        On Error GoTo 0
    End If
    
    Application.CutCopyMode = False
    
    'Show Rows and clean
    Rows("5:" & LastRow.Row - 2).Select
    sumSheet.Range("B5:G" & LastRow.Row - 2).ClearContents
    skipColLetter1 = Col_Letter(7 + cOset)
    skipColLetter2 = Col_Letter(9 + cOset)
    
    sumSheet.Range(skipColLetter1 & "5:" & skipColLetter2 & LastRow.Row - 2).ClearContents
    skipColLetter3 = Col_Letter(11 + cOset)
    sumSheet.Range(skipColLetter3 & "5:" & skipColLetter3 & LastRow.Row - 2).ClearContents
    If LastRow.Row > 9 Then
        Rows("7:" & LastRow.Row - 2).Select
        Selection.Delete
    End If
    
    'Show hidden systems
    If psSheet.Range("N3").Value <> True Then
        SetWsVisibility 1, 5
        Application.ScreenUpdating = False
    End If

    If EPC = True Then
        'Remove estimate columns
        sumSheet.Activate
        sumSheet.Range("E:F").Select
        Selection.Delete
        If eneSub = True Then
            sumSheet.Range("F5").Formula = "=SUM($D5:$E5)"
            Application.CutCopyMode = False
            sumSheet.Range("F5").Select
            Selection.Copy
            sumSheet.Range("F6").Select
            ActiveSheet.Paste
        End If
        conPSubLetter = Col_Letter(6 + cOset)
        conSubLetter = Col_Letter(7 + cOset)
        sumSheet.Range(conPSubLetter & ":" & conSubLetter).Select
        Selection.Delete

        'Update Formulas
        sysSubColLetter = Col_Letter(6 + cOset)
        If eneSub = True Then
            subSumColLetter1 = "F"
        Else
            subSumColLetter1 = "D"
        End If
        If licSub = True Then
            subSumColLetter2 = Col_Letter(5 + cOset)
        Else
            subSumColLetter2 = Col_Letter(5 + cOset)
        End If
        subSumColLetter3 = Col_Letter(7 + cOset)
        sumSheet.Range(sysSubColLetter & "5:" & sysSubColLetter & LastRow.Row - 2).Select
        Selection.Formula = "=SUM(" & subSumColLetter1 & "5:" & subSumColLetter2 & "5)"
        sumSheet.Range(LastRow.Row + 4 + rOset & ":" & LastRow.Row + 5 + rOset).Select
        Selection.EntireRow.Delete
        totalColLetter = Col_Letter(8 + cOset)
        qtyLetter = Col_Letter(7 + cOset)
        If eneSub = True Then
            sumSheet.Range(totalColLetter & LastRow.Row).Formula = "=SUMPRODUCT(" & subSumColLetter1 & "5:" & subSumColLetter1 & LastRow.Row - 2 & "," & qtyLetter & "5:" & qtyLetter & LastRow.Row - 2 & ")"
        Else
            sumSheet.Range(totalColLetter & LastRow.Row).Formula = "=SUMPRODUCT(" & subSumColLetter1 & "5:" & subSumColLetter1 & LastRow.Row - 2 & "," & subSumColLetter3 & "5:" & subSumColLetter3 & LastRow.Row - 2 & ")"
            sumSheet.Range(totalColLetter & LastRow.Row + 4).Formula = "=SUM(" & totalColLetter & LastRow.Row & ":" & totalColLetter & LastRow.Row + 3 & ")"
        End If
        sumSheet.Range("B5").Select
        Calculate
    End If
        
    
    'Build Summary Sheet
    For Each ws In ActiveWorkbook.Worksheets
        excArray = IsInArray(ws.Name, ExcSheets)
        If excArray = False And Sheets(ws.Name).Visible = True Then
            If sumRow > 5 Then
                sumSheet.Activate
                Rows(sumRow & ":" & sumRow).Select
                Selection.Copy
                Selection.Insert Shift:=xlDown
                Application.CutCopyMode = False
            End If
            ws.Select
            If EPC = False Then
                Call pullNumEST(sumRow, eneSub, licSub, supSub, cOset)
                sumRow = sumRow + 1
            ElseIf EPC = True Then
                Call pullNum(sumRow, eneSub, licSub, supSub, cOset)
                sumRow = sumRow + 1
            End If
        End If
    Next ws
    
    'Showing Summary Sheet
    ActiveWorkbook.Worksheets("Summary").Visible = xlSheetVisible
    ActiveWorkbook.Worksheets("Summary").Activate
    
    'AutoFit Rows
    Rows("5:" & sumRow).EntireRow.AutoFit
    
    'Clean Blanks
    If eneSub = True Then
        Set LastRow = sumSheet.Cells.Find(What:="TOTAL EQUIPMENT & NON-EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Else
        Set LastRow = sumSheet.Cells.Find(What:="TOTAL EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    End If
    If sumRow > 5 And LastRow.Row - 2 > sumRow Then
        Rows(sumRow + 1 & ":" & LastRow.Row - 2).Select
        Selection.Delete
    End If
    
    Application.ScreenUpdating = True
    
End Sub
Sub sumThisSheetSet()
    Dim ws As Worksheet
    Dim sumSheet As Worksheet: Set sumSheet = ActiveWorkbook.Worksheets("Summary")
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Application.ScreenUpdating = False
    
    sumSheet.Activate
    sumRow = sumSheet.Range("B" & Rows.Count).End(xlUp).Row + 1
    
    asName.Activate
    
    Call pullNum(sumRow)
    
    
    sumSheet.Activate
    If sumRow < 104 Then
        Rows(sumRow + 1 & ":104").Select
        Selection.EntireRow.Hidden = True
    End If
    
    Application.ScreenUpdating = True
        
End Sub
Sub pullNum(sumRow, eneSub, licSub, supSub, cOset)
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Dim sumSheet As Worksheet: Set sumSheet = ActiveWorkbook.Worksheets("Summary")
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")
    Dim equipT As Range: Set equipT = asName.Cells.Find(What:="TOTAL EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Dim nequipT As Range: Set nequipT = asName.Cells.Find(What:="TOTAL NON-EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    
    
    activeN = ActiveSheet.Name
    
    'Check endCell found
    If equipT Is Nothing Then
        Set equipT = asName.Cells.Find(What:="TOTAL EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    End If
    If Not equipT Is Nothing Then
        'Formulas to summary
        sumSheet.Range("B" & sumRow).Formula = "='" & activeN & "'!A2"
        sumSheet.Range("C" & sumRow).Formula = "='" & activeN & "'!D2"
        
        'Find totals column
        asName.Range(equipT.Address).Select
        sumCol = Split(ActiveCell.Offset(0, 1).Address(1, 0), "$")(0)
        sumSheet.Range("D" & sumRow).Formula = "='" & activeN & "'!" & sumCol & equipT.Row
        sumSheet.Range("E" & sumRow).Formula = "='" & activeN & "'!" & sumCol & nequipT.Row
        If licSub = True Then
            Dim licT As Range: Set licT = asName.Cells.Find(What:="LICENSE COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
            licColLetter = Col_Letter(5 + cOset)
            sumSheet.Range(licColLetter & sumRow).Formula = "='" & activeN & "'!" & sumCol & licT.Row
        End If
        
        'Room Count
        If asName.Range("C2").Value = "Room Numbers" Then
            rCount = CommaCount(asName.Range("D2"))
            sumSheet.Cells(sumRow, 7 + cOset).Value = rCount + 1
        ElseIf asName.Range("C2").Value = "System Count" Then
            sumSheet.Cells(sumRow, 7 + cOset).Value = asName.Range("D2").Value
        End If
        
        
        'Link to sheet
        sumSheet.Activate
        sumSheet.Range("B" & sumRow).Select
        sheetName = Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1)
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=sheetName

    End If
    
End Sub
Sub pullNumEST(sumRow, eneSub, licSub, supSub, cOset)
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    Dim sumSheet As Worksheet: Set sumSheet = ActiveWorkbook.Worksheets("Summary")
    Dim psSheet As Worksheet: Set psSheet = ActiveWorkbook.Worksheets("PROJECT_SETTINGS")
    Dim equipT As Range: Set equipT = asName.Cells.Find(What:="TOTAL EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Dim dPerequipT As Range: Set dPerequipT = asName.Cells.Find(What:="DISCOUNT FROM MSRP", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Dim DequipT As Range: Set DequipT = asName.Cells.Find(What:="DISCOUNTED EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Dim nequipT As Range: Set nequipT = asName.Cells.Find(What:="TOTAL NON-EQUIPMENT COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Dim conPerT As Range: Set conPerT = asName.Cells.Find(What:="CONTINGENCY PERCENTAGE", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    Dim conT As Range: Set conT = asName.Cells.Find(What:="CONTINGENCY", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    
    
    activeN = ActiveSheet.Name
    
    'Check endCell found
    If equipT Is Nothing Then
        Set equipT = asName.Cells.Find(What:="//", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    End If
    If Not equipT Is Nothing Then
        'Formulas to summary
        sumSheet.Range("B" & sumRow).Formula = "='" & activeN & "'!A2"
        sumSheet.Range("C" & sumRow).Formula = "='" & activeN & "'!D2"
        sumSheet.Range("C" & sumRow).NumberFormat = "General"
        
        'Find totals column
        asName.Range(equipT.Address).Select
        sumCol = Split(ActiveCell.Offset(0, 1).Address(1, 0), "$")(0)
        sumSheet.Range("D" & sumRow).Formula = "='" & activeN & "'!" & sumCol & equipT.Row
        sumSheet.Range("E" & sumRow).Formula = "='" & activeN & "'!" & sumCol & dPerequipT.Row
        sumSheet.Range("F" & sumRow).Formula = "='" & activeN & "'!" & sumCol & DequipT.Row
        sumSheet.Range("G" & sumRow).Formula = "='" & activeN & "'!" & sumCol & nequipT.Row
        If licSub = True Then
            Dim licT As Range: Set licT = asName.Cells.Find(What:="LICENSE COST SUBTOTAL", After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
            licColLetter = Col_Letter(7 + cOset)
            conPerTColLetter = Col_Letter(8 + cOset)
            conTColLetter = Col_Letter(9 + cOset)
            sumSheet.Range(licColLetter & sumRow).Formula = "='" & activeN & "'!" & sumCol & licT.Row
            sumSheet.Range(conPerTColLetter & sumRow).Formula = "='" & activeN & "'!" & sumCol & conPerT.Row
            sumSheet.Range(conTColLetter & sumRow).Formula = "='" & activeN & "'!" & sumCol & conT.Row
        Else
            sumSheet.Range("H" & sumRow).Formula = "='" & activeN & "'!" & sumCol & conPerT.Row
            sumSheet.Range("I" & sumRow).Formula = "='" & activeN & "'!" & sumCol & conT.Row
        End If
        
        'Room Count
        If asName.Range("C2").Value = "Room Numbers" Then
            rCount = CommaCount(asName.Range("D2"))
            sumSheet.Cells(sumRow, 11 + cOset).Value = rCount + 1
        ElseIf asName.Range("C2").Value = "System Count" Then
            sumSheet.Cells(sumRow, 11 + cOset).Value = asName.Range("D2").Value
        End If
        
        'Link to sheet
        sumSheet.Activate
        sumSheet.Range("B" & sumRow).Select
        sheetName = Right(ActiveCell.Formula, Len(ActiveCell.Formula) - 1)
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=sheetName

    End If
    
End Sub
Function CommaCount(rng As Range) As Integer
    Dim strTest As String, i As Integer
    strTest = rng.text
    i = Len(strTest)
    strTest = Application.WorksheetFunction.Substitute(strTest, ",", "")
    CommaCount = i - Len(strTest)
End Function

