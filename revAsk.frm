VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} revAsk 
   Caption         =   "Update Revision List"
   ClientHeight    =   3672
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4944
   OleObjectBlob   =   "revAsk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "revAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub canActive_Click()
    Unload Me
    End
End Sub

Private Sub ComboBox1_Change()
    If Me.ComboBox1.Value = "Add Issuance" Then
        Me.Label8.Visible = True
        Me.TextBox1.Visible = True
    Else
        Me.Label8.Visible = False
        Me.TextBox1.Visible = False
    End If
End Sub
Private Sub contActive_Click()
    Me.Hide
    If ValidFileName(TextBox1.Value) = False Then
        MsgBox "Your name must not include any of the following characters: : \ / ? * [ ].  Please enter the name again."
        TextBox1.Value = ""
        Me.Show
    End If
End Sub

Private Sub Label7_Click()

End Sub
Private Sub UserForm_Initialize()
    Dim n As Long
    Dim i As Long
    
    
    With ActiveWorkbook.Worksheets("Issuances")
        n = .Cells(Rows.Count, 1).End(xlUp).Row
    End With
    With ComboBox1
        .Clear
        .AddItem "Add Issuance"
        For i = 5 To n
            .AddItem ActiveWorkbook.Worksheets("Issuances").Cells(i, 1).Value
        Next i
    End With
'    ComboBox1.Value = "Add Issuance"
    With Me
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
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

