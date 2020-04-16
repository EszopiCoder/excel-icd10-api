VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formLog 
   Caption         =   "Log"
   ClientHeight    =   5892
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8772
   OleObjectBlob   =   "formLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Set up textbox
    txtLog.ScrollBars = fmScrollBarsVertical
    txtLog.MultiLine = True
    
    ' Add tooltip text
    btnExport.ControlTipText = "Export log to text document"
    btnClose.ControlTipText = "Close log"
End Sub

Private Sub txtLog_Change()
    txtLog.SetFocus
    txtLog.SelStart = 0
    txtLog.SelLength = 0
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub btnExport_Click()
    If Len(txtLog.Text) = 0 Then
        MsgBox "Log is empty", vbInformation
        Exit Sub
    End If
    Call modSaveDialog.SaveAsTextDoc(txtLog.Text)
End Sub
