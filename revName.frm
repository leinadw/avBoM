VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} revName 
   Caption         =   "New Issuance or Archive Name"
   ClientHeight    =   1230
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4200
   OleObjectBlob   =   "revName.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "revName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub contActive_Click()
    Me.Hide
    If ValidFileName(TextBox1.Value) = False Then
        MsgBox "Your name must not include any of the following characters: : \ / ? * [ ].  Please enter the name again."
        TextBox1.Value = ""
        Me.Show
    End If
End Sub
Private Sub UserForm_Initialize()
    subRemoveCloseButton Me
    With Me
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

contInactive.Visible = True

End Sub
Sub contInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

contInactive.Visible = False

End Sub

