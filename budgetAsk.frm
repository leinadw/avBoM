VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} budgetAsk 
   ClientHeight    =   7548
   ClientLeft      =   72
   ClientTop       =   300
   ClientWidth     =   8976.001
   OleObjectBlob   =   "budgetAsk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "budgetAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
         For i = 0 To ListBox1.ListCount - 1
             ListBox1.Selected(i) = True
         Next i
    End If
    
    If CheckBox1.Value = False Then
         For i = 0 To ListBox1.ListCount - 1
             ListBox1.Selected(i) = False
         Next i
    End If
End Sub
Private Sub CheckBox2_Click()
    If CheckBox2.Value = True Then
         For i = 0 To ListBox2.ListCount - 1
             ListBox2.Selected(i) = True
         Next i
    End If
    
    If CheckBox2.Value = False Then
         For i = 0 To ListBox2.ListCount - 1
             ListBox2.Selected(i) = False
         Next i
    End If
End Sub

Private Sub contActive_Click()
    Dim dataHold As Worksheet: Set dataHold = ActiveWorkbook.Worksheets("DATA_HOLD")
    dataHold.Range("B:B").Clear
    For i = 1 To Me.ListBox2.ListCount
      dataHold.Range("B" & i).Value = Me.ListBox2.List(i - 1)
    Next i
    
    If notePubY = True Then
        noteStat = True
    ElseIf notePubN = True Then
        noteStat = False
    End If
    
    If costPubY = True Then
        costStat = True
    ElseIf costPubN = True Then
        costStat = False
    End If
    
    If equipPubY = True Then
        equipStat = True
    ElseIf equipPubN = True Then
        equipStat = False
    End If
    
    If laborPubY = True Then
        laborStat = True
    ElseIf laborPubN = True Then
        laborStat = False
    End If
    
    budgetAsk.Hide
End Sub
Private Sub addActive_Click()
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
        Z = 0
        For X = 0 To ListBox2.ListCount - 1
            If ListBox2.List(X) = ListBox1.List(i) Then
                Z = Z + 1
            End If
        Next X
        If Z = 0 Then
            ListBox2.AddItem ListBox1.List(i)
            ListBox1.Selected(i) = False
        Else
            ListBox1.Selected(i) = False
        End If
    End If
    Next i
    
    CheckBox1.Value = False
    CheckBox2.Value = False
End Sub

Private Sub canActive_Click()
    Unload Me
    End
End Sub


Private Sub Label7_Click()

End Sub

Private Sub notePubN_Click()

End Sub

Private Sub removeActive_Click()
Dim counter As Integer
    counter = 0

    For i = 0 To ListBox2.ListCount - 1
         If ListBox2.Selected(i - counter) Then
             ListBox2.RemoveItem (i - counter)
             counter = counter + 1
         End If
    Next i
    
    CheckBox1.Value = False
    CheckBox2.Value = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Unload Me
        End
    End If
End Sub
Private Sub UserForm_Initialize()
    Dim n As Long
    Dim i As Long
    
    
    With ActiveWorkbook.Worksheets("DATA_HOLD")
        n = .Cells(Rows.Count, 1).End(xlUp).Row
    End With
    With ListBox1
        .Clear
        For i = 1 To n
            .AddItem ActiveWorkbook.Worksheets("DATA_HOLD").Cells(i, 1).Value
        Next i
    End With
            
    With Me
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

contInactive.Visible = True
canInactive.Visible = True
addInactive.Visible = True
removeInactive.Visible = True


End Sub
Sub contInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

contInactive.Visible = False
canInactive.Visible = True
addInactive.Visible = True
removeInactive.Visible = True

End Sub
Sub canInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

contInactive.Visible = True
canInactive.Visible = False
addInactive.Visible = True
removeInactive.Visible = True

End Sub
Sub addInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

contInactive.Visible = True
canInactive.Visible = True
addInactive.Visible = False
removeInactive.Visible = True

End Sub
Sub removeInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

contInactive.Visible = True
canInactive.Visible = True
addInactive.Visible = True
removeInactive.Visible = False

End Sub
