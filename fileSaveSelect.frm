VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fileSaveSelect 
   ClientHeight    =   1320
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3060
   OleObjectBlob   =   "fileSaveSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fileSaveSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub contActive_Click()
    Me.Hide
End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

contInactive.Visible = True

End Sub
Sub contInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

contInactive.Visible = False

End Sub
