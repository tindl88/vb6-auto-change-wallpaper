VERSION 5.00
Begin VB.Form frmView 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   216
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   293
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   3105
      Left            =   0
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   0
      Top             =   0
      Width           =   4185
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If frmMain.chkOnTop.Value = 1 Then Call OnTop(Me, True)
End Sub

Private Sub Form_Resize()
    Picture1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    FitPicture frmMain.PicVisible, frmView.Picture1
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Unload Me
End Sub
