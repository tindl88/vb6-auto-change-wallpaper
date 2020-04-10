VERSION 5.00
Begin VB.Form frmSlideOption 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Slide Show"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   525
      Left            =   1620
      TabIndex        =   0
      Top             =   1620
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1575
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   3255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   150
         ScaleHeight     =   975
         ScaleWidth      =   3015
         TabIndex        =   3
         Top             =   270
         Width           =   3015
         Begin VB.OptionButton Option1 
            Caption         =   "Normal"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Shuffle"
            Height          =   255
            Left            =   450
            TabIndex        =   7
            Top             =   330
            Width           =   885
         End
         Begin VB.Frame Frame2 
            Caption         =   "Delay"
            Height          =   975
            Left            =   1530
            TabIndex        =   4
            Top             =   0
            Width           =   1335
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "frmSlideOption.frx":0000
               Left            =   150
               List            =   "frmSlideOption.frx":0022
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   390
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "s"
               Height          =   195
               Left            =   1140
               TabIndex        =   6
               Top             =   420
               Width           =   75
            End
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Random"
            Height          =   345
            Left            =   0
            TabIndex        =   8
            Top             =   630
            Width           =   1665
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   1290
         Width           =   3045
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   525
      Left            =   60
      TabIndex        =   1
      Top             =   1620
      Width           =   1575
   End
End
Attribute VB_Name = "frmSlideOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Load frmSlideShow
    frmSlideShow.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call MoveForm(Me, FrmCenter)
    If frmMain.chkOnTop.Value = 1 Then Call OnTop(Me, True)
    
    Option1.Value = GetSetting("AC Wallpaper", "SlideShow", "Normal")
    Option2.Value = GetSetting("AC Wallpaper", "SlideShow", "Random")
    Check1.Value = GetSetting("AC Wallpaper", "SlideShow", "Shuffle")
    Combo1.ListIndex = GetSetting("AC Wallpaper", "SlideShow", "Delay", 2)

    Label2.Caption = "Space: Pause/Resume SS" & Space$(9) & "Esc: End SS"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "AC Wallpaper", "SlideShow", "Normal", Option1.Value
    SaveSetting "AC Wallpaper", "SlideShow", "Random", Option2.Value
    SaveSetting "AC Wallpaper", "SlideShow", "Shuffle", Check1.Value
    SaveSetting "AC Wallpaper", "SlideShow", "Delay", Combo1.ListIndex
End Sub

Private Sub Option1_Click()
    Check1.Enabled = True
End Sub

Private Sub Option2_Click()
    Check1.Enabled = False
End Sub

