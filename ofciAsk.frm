VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ofciAsk 
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3768
   OleObjectBlob   =   "ofciAsk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ofciAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
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

Private Sub canActive_Click()
    Unload Me
    End
End Sub

Private Sub contActive_Click()
    Me.Hide
End Sub
