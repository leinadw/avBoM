VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SysAsk 
   Caption         =   "Add a New System Tab"
   ClientHeight    =   2640
   ClientLeft      =   84
   ClientTop       =   300
   ClientWidth     =   3672
   OleObjectBlob   =   "SysAsk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SysAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub contActive_Click()
    Me.Hide
End Sub

Private Sub OptionButton3_Click()
    Me.ComboBox1.Enabled = True
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    excArray = IsInArray(asName.Name, ExcSheets)
    If excArray = False Then
        asNameCount = ActiveWorkbook.Worksheets("DATA_HOLD").Cells.Find(What:=asName.Name, After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
        Me.ComboBox1.ListIndex = asNameCount - 1
    End If
End Sub
Private Sub UserForm_Initialize()
    Dim n As Long
    Dim i As Long
    
    With ActiveWorkbook.Worksheets("DATA_HOLD")
        n = .Cells(Rows.Count, 1).End(xlUp).Row
    End With
    With ComboBox1
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
Private Sub canActive_Click()
    Unload Me
    End
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Unload Me
        End
    End If
End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

contInactive.Visible = True
canInactive.Visible = True

End Sub
Sub contInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

contInactive.Visible = False
canInactive.Visible = True

End Sub
Sub canInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

contInactive.Visible = True
canInactive.Visible = False

End Sub

