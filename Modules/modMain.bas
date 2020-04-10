Attribute VB_Name = "modMain"
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" ()

Public Sub Main()
    On Error Resume Next
    InitCommonControls
    If App.PrevInstance Then End
    'Show Form
    If frmMain.chkStartMode.Value = 1 Then
        frmMain.cmdControl_Click 5
    Else
        frmMain.Show
    End If
End Sub
