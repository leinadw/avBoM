VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} issueCheck 
   Caption         =   "Overwrite Existing Issuance"
   ClientHeight    =   1664
   ClientLeft      =   108
   ClientTop       =   444
   ClientWidth     =   4320
   OleObjectBlob   =   "issueCheck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "issueCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub yesActive_Click()
    Dim iSheet As Worksheet: Set iSheet = ActiveWorkbook.Worksheets("Issuances")
    
    'Find Issuance Line
    iName = revAsk.ComboBox1.Value
    iSheet.Activate
    iRow = iSheet.Cells.Find(What:=iName, After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
    
    'Clean Issuance
    iSheet.Range("B" & iRow & ":C" & iRow).ClearContents
    
    
    Unload Me
End Sub
Private Sub noActive_Click()
    revAsk.Show
    Unload Me
End Sub
Private Sub UserForm_Initialize()
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

yesInactive.Visible = True
noInactive.Visible = True

End Sub
Sub yesInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

yesInactive.Visible = False
noInactive.Visible = True

End Sub
Sub noInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

noInactive.Visible = False
yesInactive.Visible = True

End Sub
