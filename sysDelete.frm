VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sysDelete 
   Caption         =   "Select System To Delete"
   ClientHeight    =   1845
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "sysDelete.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sysDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub contInactive_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()
    Dim n As Long
    Dim i As Long
    Dim asName As Worksheet: Set asName = ActiveWorkbook.ActiveSheet
    
    With ActiveWorkbook.Worksheets("DATA_HOLD")
        n = .Cells(Rows.Count, 1).End(xlUp).Row
    End With
    With ComboBox1
        .Clear
        .AddItem ""
        For i = 1 To n
            .AddItem ActiveWorkbook.Worksheets("DATA_HOLD").Cells(i, 1).Value
        Next i
    End With
    
    On Error Resume Next
    asNameCount = ActiveWorkbook.Worksheets("DATA_HOLD").Cells.Find(What:=asName.Name, After:=ActiveCell, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
    Me.ComboBox1.ListIndex = asNameCount
    On Error GoTo 0
    
    If ActiveWorkbook.Worksheets("DATA_HOLD").Range("A1").Value = "" Then
        MsgBox "There are no System Tabs to delete."
        Unload sysDelete
        End
    End If
    
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
