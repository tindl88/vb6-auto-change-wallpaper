VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00F4F5EB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AC Wallpaper"
   ClientHeight    =   6510
   ClientLeft      =   150
   ClientTop       =   555
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOptions 
      BackColor       =   &H00FEFFEB&
      Caption         =   "Options"
      Height          =   300
      Left            =   2850
      TabIndex        =   48
      Top             =   4230
      Width           =   1050
   End
   Begin VB.PictureBox PicOptions 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1125
      Left            =   500
      ScaleHeight     =   1125
      ScaleWidth      =   2895
      TabIndex        =   43
      Top             =   2730
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CheckBox chkBeep 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sound"
         Height          =   255
         Left            =   90
         TabIndex        =   47
         Top             =   810
         Width           =   765
      End
      Begin VB.CheckBox chkOnTop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "On Top"
         Height          =   255
         Left            =   90
         TabIndex        =   46
         Top             =   570
         Width           =   825
      End
      Begin VB.CheckBox chkStartMode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start Minimized in System Tray"
         Height          =   255
         Left            =   90
         TabIndex        =   45
         Top             =   90
         Width           =   2505
      End
      Begin VB.CheckBox chkStartWin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Start AC Wallpaper with Windows"
         Height          =   255
         Left            =   90
         TabIndex        =   44
         Top             =   330
         Width           =   2715
      End
      Begin VB.Shape ShapeOptions 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   1125
         Left            =   0
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdControl 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Hide"
      Height          =   435
      Index           =   5
      Left            =   7380
      TabIndex        =   31
      Top             =   5730
      Width           =   825
   End
   Begin VB.CommandButton cmdControl 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Set Wallpaper"
      Height          =   435
      Index           =   4
      Left            =   5385
      TabIndex        =   22
      Top             =   5730
      Width           =   2010
   End
   Begin VB.CommandButton cmdControl 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Slide Show"
      Height          =   435
      Index           =   6
      Left            =   3890
      TabIndex        =   36
      Top             =   5730
      Width           =   1515
   End
   Begin VB.PictureBox PicTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      Picture         =   "frmMain.frx":57E2
      ScaleHeight     =   945
      ScaleWidth      =   8265
      TabIndex        =   34
      Top             =   0
      Width           =   8265
      Begin VB.Image Image1 
         Height          =   735
         Left            =   150
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.x"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   435
         Left            =   5520
         TabIndex        =   37
         Top             =   410
         Width           =   645
      End
   End
   Begin VB.PictureBox PicStatus 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F4F5EB&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8250
      TabIndex        =   32
      Top             =   6195
      Width           =   8250
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2475
         Top             =   0
      End
      Begin VB.Label lbPathName 
         AutoSize        =   -1  'True
         BackColor       =   &H00F4F5EB&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   60
         TabIndex        =   33
         Top             =   75
         Width           =   90
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   8280
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   0
         X2              =   8280
         Y1              =   15
         Y2              =   0
      End
   End
   Begin VB.CommandButton cmdControl 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Remove All"
      Height          =   435
      Index           =   3
      Left            =   2925
      TabIndex        =   4
      Top             =   5730
      Width           =   975
   End
   Begin VB.CommandButton cmdControl 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Remove"
      Height          =   435
      Index           =   2
      Left            =   1965
      TabIndex        =   27
      Top             =   5730
      Width           =   975
   End
   Begin VB.CommandButton cmdControl 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Add File(s)"
      Height          =   435
      Index           =   1
      Left            =   885
      TabIndex        =   28
      Top             =   5730
      Width           =   1095
   End
   Begin VB.CommandButton cmdControl 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Add Dir"
      Height          =   435
      Index           =   0
      Left            =   45
      TabIndex        =   29
      Top             =   5730
      Width           =   855
   End
   Begin VB.CommandButton cmdControl 
      BackColor       =   &H00FEFFEB&
      Caption         =   "Set"
      Height          =   300
      Index           =   7
      Left            =   2190
      TabIndex        =   26
      Top             =   4230
      Width           =   670
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1590
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   25
      Top             =   4275
      Width           =   570
   End
   Begin VB.PictureBox PicBorder 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3210
      Left            =   3945
      ScaleHeight     =   214
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   23
      Top             =   990
      Width           =   4255
      Begin VB.PictureBox lbArt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   15
         ScaleHeight     =   3180
         ScaleWidth      =   4230
         TabIndex        =   40
         Top             =   15
         Width           =   4230
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "tindl88"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   240
            Left            =   1830
            TabIndex        =   42
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "www.caulacbovb.com"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1200
            TabIndex        =   41
            Top             =   1380
            Width           =   1875
         End
      End
      Begin VB.PictureBox PicShow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00E0E0E0&
         Height          =   3180
         Left            =   15
         ScaleHeight     =   212
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   282
         TabIndex        =   24
         Top             =   15
         Width           =   4230
      End
   End
   Begin VB.Frame frmPicInfomation 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Picture Infomation"
      Height          =   1160
      Left            =   45
      TabIndex        =   0
      Top             =   4500
      Width           =   3840
      Begin VB.PictureBox PicIcon 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F4F5EB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   180
         ScaleHeight     =   525
         ScaleWidth      =   555
         TabIndex        =   38
         Top             =   390
         Width           =   555
      End
      Begin VB.Label lbDimensions 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensions :"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   855
         TabIndex        =   3
         Top             =   270
         Width           =   900
      End
      Begin VB.Label lbFileLen 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "File Size      :"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   855
         TabIndex        =   2
         Top             =   810
         Width           =   900
      End
      Begin VB.Label lbExt 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Extension   :"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   855
         TabIndex        =   1
         Top             =   540
         Width           =   900
      End
   End
   Begin VB.ListBox List1 
      Height          =   3210
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":FD63
      Left            =   30
      List            =   "frmMain.frx":FD65
      TabIndex        =   35
      Top             =   990
      Width           =   3855
   End
   Begin VB.Frame FramePosition 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Position"
      Height          =   1455
      Left            =   3945
      TabIndex        =   17
      Top             =   4230
      Width           =   1395
      Begin VB.PictureBox PicWallTypes 
         BackColor       =   &H00F4F5EB&
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
         Height          =   1155
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   1155
         TabIndex        =   18
         Top             =   240
         Width           =   1155
         Begin VB.OptionButton OptStyle 
            BackColor       =   &H00F4F5EB&
            Caption         =   "Tile"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   915
         End
         Begin VB.OptionButton OptStyle 
            BackColor       =   &H00F4F5EB&
            Caption         =   "Center"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   420
            Width           =   915
         End
         Begin VB.OptionButton OptStyle 
            BackColor       =   &H00F4F5EB&
            Caption         =   "Stretch"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   0
            Value           =   -1  'True
            Width           =   915
         End
      End
   End
   Begin VB.Frame FrameTimer 
      BackColor       =   &H00F4F5EB&
      Caption         =   "Time"
      Height          =   1455
      Left            =   5385
      TabIndex        =   5
      Top             =   4230
      Width           =   2820
      Begin VB.PictureBox PicTimes 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F5EB&
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
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   2595
         TabIndex        =   6
         Top             =   240
         Width           =   2595
         Begin VB.CommandButton cmdPause 
            BackColor       =   &H00F4F5EB&
            Caption         =   "Pause"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1260
            TabIndex        =   7
            Tag             =   "0"
            Top             =   720
            Width           =   1275
         End
         Begin VB.ComboBox cbTime 
            BackColor       =   &H00F4F5EB&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            ItemData        =   "frmMain.frx":FD67
            Left            =   1740
            List            =   "frmMain.frx":FE1F
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   0
            Width           =   615
         End
         Begin VB.ComboBox cbTime 
            BackColor       =   &H00F4F5EB&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            ItemData        =   "frmMain.frx":FF13
            Left            =   900
            List            =   "frmMain.frx":FFCB
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   0
            Width           =   615
         End
         Begin VB.ComboBox cbTime 
            BackColor       =   &H00F4F5EB&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            ItemData        =   "frmMain.frx":100BF
            Left            =   60
            List            =   "frmMain.frx":1010B
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   0
            Width           =   615
         End
         Begin VB.Timer TmrCTime 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   2250
            Top             =   270
         End
         Begin VB.CommandButton cmdStart 
            BackColor       =   &H00F4F5EB&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Tag             =   "Stop"
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton cmdStop 
            BackColor       =   &H00F4F5EB&
            Caption         =   "Stop"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Tag             =   "0"
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label lbHours 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   720
            TabIndex        =   16
            Top             =   60
            Width           =   135
         End
         Begin VB.Label lbMyTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00:00:00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   600
            TabIndex        =   15
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label lbMin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "M"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1560
            TabIndex        =   14
            Top             =   60
            Width           =   150
         End
         Begin VB.Label lbSec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2400
            TabIndex        =   13
            Top             =   60
            Width           =   120
         End
      End
   End
   Begin VB.PictureBox PicVisible 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   5400
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   39
      Top             =   1050
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lbDeskBckGrd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop Background"
      Height          =   195
      Left            =   60
      TabIndex        =   30
      Top             =   4275
      Width           =   1470
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Visible         =   0   'False
      Begin VB.Menu mnuFolder 
         Caption         =   "Open Folder"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete File"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelDead 
         Caption         =   "Remove Dead Entries"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemDup 
         Caption         =   "Remove Duplicate"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuWall 
      Caption         =   "Wall"
      Visible         =   0   'False
      Begin VB.Menu mnuSetwallpaper 
         Caption         =   "Set Wallpaper"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Minutes As Byte
Dim Seconds As Byte
Dim MyTime As Date
Dim Hours As Byte

Private Sub ConvertTime()
    Hours = Val(cbTime(0).Text)
    Minutes = Val(cbTime(1).Text)
    Seconds = Val(cbTime(2).Text)
    MyTime = TimeSerial(Hours, Minutes, Seconds)
    lbMyTime.Caption = Format$(MyTime, "HH") & ":" & Format$(MyTime, "nn") & ":" & Format$(MyTime, "ss")
End Sub

Private Sub cbTime_Change(Index As Integer)
    ConvertTime
End Sub

Private Sub cbTime_Click(Index As Integer)
    ConvertTime
End Sub

Private Sub chkOnTop_Click()
    Call OnTop(Me, chkOnTop)
End Sub

Private Sub chkStartWin_Click()
    If chkStartWin.Value = 1 Then
        RegCreateKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AC Wallpaper", App.Path & "\" & App.EXEName & ".exe"
    Else
        RegDeleteKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AC Wallpaper"
    End If
End Sub

Private Sub cmdOptions_Click()
    If cmdOptions.tag = 0 Then
        cmdOptions.tag = 1
        PicOptions.Visible = True
        PicOptions.Move cmdOptions.Width - 50, cmdOptions.Top - PicOptions.Height
    Else
        cmdOptions.tag = 0
        PicOptions.Visible = False
    End If
End Sub

Private Sub cmdStop_Click()
    cmdStop.Visible = False
    cmdStart.Visible = True
    cbTime(0).Enabled = True
    cbTime(1).Enabled = True
    cbTime(2).Enabled = True
    Me.Caption = "AC Wallpaper"
    cmdPause.Enabled = False
    cmdPause.tag = 0
    cmdPause.Caption = "Pause"
    TmrCTime.Enabled = False
    ConvertTime
    cmdStart.tag = "Stop"
End Sub

Public Sub cmdStart_Click()
    If lbMyTime.Caption = "00:00:00" Then
        MsgBox "Please! Set Times.", vbInformation, "AC Wall"
        Exit Sub
    ElseIf List1.ListCount = 0 Then
        MsgBox "Please! Add Pictures To The List.", vbInformation, "AC Wall"
        Exit Sub
    Else
        cmdStart.Visible = False
        cmdStop.Visible = True
        cbTime(0).Enabled = False
        cbTime(1).Enabled = False
        cbTime(2).Enabled = False
        cmdPause.Enabled = True
        TmrCTime.Enabled = True
        cmdStart.tag = "Play"
    End If
End Sub

Private Sub Image1_Click()
    If MsgBox(".::" & Me.Caption & "::." & " Version " & App.Major & "." & App.Minor & vbCrLf & vbCrLf & "Supports: " & UCase$("jpg,jpeg,gif,bmp,wmf,dib,pcx,tga,png,tif,ico,cur") & vbCrLf & vbCrLf & "Email     : tindl88@yahoo.com" & vbCrLf & "Website: www.caulacbovb.com" & vbCrLf & vbCrLf & "Contact to tindl88?", vbInformation + vbYesNo, "About") = vbYes Then
        Shell "EXPLORER.EXE " & "ymsgr:sendIM?tindl88"
    Else
        Exit Sub
    End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Shell "EXPLORER.EXE " & "http://www.caulacbovb.com", vbMaximizedFocus
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Shell "EXPLORER.EXE " & "ymsgr:sendIM?tindl88"
End Sub

Private Sub lbPathName_DblClick()
    mnuFolder_Click
End Sub

Private Sub List1_Click()
On Error Resume Next
    If List1.SelCount = 1 Then
        If ObjExist(List1.List(List1.ListIndex)) = True Then
            lbArt.Visible = False
            'Load and Fitting Picture
            If Right(List1.List(List1.ListIndex), 4) = ".pcx" Then
                PCX.LoadPCX List1.List(List1.ListIndex)
                PCX.DrawPCX PicVisible
            ElseIf Right(List1.List(List1.ListIndex), 4) = ".tga" Then
                TGA.LoadTGA List1.List(List1.ListIndex)
                TGA.DrawTGA PicVisible
            ElseIf Right(List1.List(List1.ListIndex), 4) = ".png" Then
                PNG.DrawPNG = PicVisible
                PNG.LoadPNG List1.List(List1.ListIndex)
            ElseIf Right(List1.List(List1.ListIndex), 4) = ".tif" Then
                TIF.LoadTIFF List1.List(List1.ListIndex)
            Else
                PicVisible.Picture = LoadPicture(List1.List(List1.ListIndex))
            End If
            FitPicture PicVisible, PicShow
            'Get Dimensions
            lbDimensions.Caption = "Dimensions : " & PicVisible.ScaleWidth & "x" & PicVisible.ScaleHeight
            'Get Icon
            GetIcon PicIcon, List1.List(List1.ListIndex), 1
            'Get Extension
            lbExt.Caption = "Extension   : " & StrConv(GetExtOfFileName(List1.List(List1.ListIndex)), vbUpperCase)
            'Get File Size
            lbFileLen.Caption = "File Size      : " & ConvertBytes(FileLen(List1.List(List1.ListIndex)))
        Else
            RemoveDeadEntries List1
            ListBoxHScroll List1
        End If
    End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ITem As Integer
If Button = vbRightButton Then
    ITem = ((y / Screen.TwipsPerPixelY) \ (TextHeight("A") / Screen.TwipsPerPixelY)) + List1.TopIndex
    If ITem > List1.ListCount - 1 Then ITem = -1
    List1.ListIndex = ITem
    If List1.SelCount = 1 Then PopupMenu mnuFile
End If
End Sub

Private Sub mnuDelDead_Click()
    RemoveDeadEntries List1
End Sub

Private Sub mnuDelete_Click()
    If MsgBox("Are you sure you want to delete selected?", vbInformation + vbYesNo, "Delete") = vbYes Then
        Kill List1.List(List1.ListIndex)
        cmdControl_Click 2
    End If
End Sub

Private Sub mnuFolder_Click()
    If List1.ListCount > 0 Then Shell "Explorer.exe " & GetFolderPathFromPath(List1.List(List1.ListIndex)), vbMaximizedFocus
End Sub

Private Sub mnuProperties_Click()
    ShowPropeties List1.List(List1.ListIndex), Me.hWnd
End Sub

Private Sub mnuRemDup_Click()
    RemoveDuplicates List1
End Sub

Private Sub mnusetwallpaper_Click()
    cmdControl_Click 4
End Sub

Private Sub PicShow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Load frmView: frmView.Show
    If Button = 2 Then PopupMenu mnuWall
End Sub

Private Sub Picture2_Click()
    ShowColor Me, Picture2
End Sub

Private Sub cmdPause_Click()
    If cmdPause.tag = 0 Then
        cmdPause.tag = 1
        Me.Caption = "AC Wallpaper - Paused..."
        cmdPause.Caption = "Resume"
        TmrCTime.Enabled = False
    Else
        cmdPause.tag = 0
        cbTime(1).Locked = True
        cbTime(2).Locked = True
        cbTime(0).Locked = True
        Me.Caption = "AC Wallpaper"
        cmdPause.Caption = "Pause"
        TmrCTime.Enabled = True
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lMsg As Single
    lMsg = x / Screen.TwipsPerPixelX
    Select Case lMsg
        Case WM_LBUTTONDOWN: Me.Show
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingsINI
    RemoveTrayIcon
    
    Set PCX = Nothing
    Set TGA = Nothing
    Set PNG = Nothing
    Set TIF = Nothing
    
    Set frmMain = Nothing
    End
End Sub

Public Sub cmdControl_Click(Index As Integer)
Select Case Index
    Case 0  'Add Folder
            Dim folder As String
            folder = BrowseForFolder(Me, "Select Folder", GetSetting("AC Wallpaper", "Settings", "BrowseFolder"))
            If Len(folder) = 0 Then Exit Sub  'User Selected Cancel
            AddFolder folder
            ListBoxHScroll List1
    Case 1  'Add Files
            AddFiles GetSetting("AC Wallpaper", "Settings", "LastDir"): ListBoxHScroll List1
    Case 2  'Remove Selected
            Dim LastIndex As Long
                If List1.ListCount = 1 Then
                    cmdControl_Click 3
                Else
                    LastIndex = List1.ListIndex
                    List1.RemoveItem LastIndex
                    If LastIndex = List1.ListCount Then LastIndex = List1.ListCount - 1
                    If LastIndex = 0 Then List1.ListIndex = LastIndex
                    List1.Selected(LastIndex) = True
                End If
            ListBoxHScroll List1
    Case 3  'Remove All
            List1.Clear
            PicVisible.Picture = LoadPicture("")
            PicShow.Picture = LoadPicture("")
            lbArt.Visible = True
            lbDimensions.Caption = "Dimensions :"
            lbExt.Caption = "Extension   :"
            lbFileLen.Caption = "File Size      :"
            PicIcon.Cls
            ListBoxHScroll List1
    Case 4 'Set Wallpaper
            If OptStyle(0).Value = True Then            'Stretch
                SetWallpaper PicVisible, WallStretch
            ElseIf OptStyle(1).Value = True Then        'Center
                SetWallpaper PicVisible, WallCenter
            Else
                SetWallpaper PicVisible, WallTile       'Tile
            End If
            'Ding
            If chkBeep.Value = 1 Then PlaySnd "ding.wav"
    Case 5: Me.Hide: AddTrayIcon Me, Me.Caption & Space$(1) & App.Major & "." & App.Minor
    Case 6
            If List1.ListCount = 0 Then
                MsgBox "Please! Add Pictures To The List.", vbInformation, "AC Wall"
                Exit Sub
            Else
                frmSlideOption.Show
            End If
    Case 7  'Set Desktop Background
            SetSysColors 1, COLOR_BACKGROUND, Picture2.BackColor '1 là Desktop
            FitPicture PicVisible, PicShow
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call MoveForm(Me, FrmBottomRight)
    PicTitle.PaintPicture PicTitle.Picture, 6360, 0, PicTitle.Width + 6360, PicTitle.Height, 6000, 0, 440, PicTitle.Height
    lbArt.Visible = True
    Me.Caption = "AC Wallpaper"
    cmdOptions.tag = 0
    Label1.Caption = App.Major & "." & App.Minor
    Picture2.BackColor = GetSysColor(COLOR_BACKGROUND)
    AddTrayIcon Me, Me.Caption & Space$(1) & App.Major & "." & App.Minor
    LoadSettingsINI
End Sub

Private Sub TmrCTime_Timer()
    If (Format$(MyTime, "HH") & ":" & Format$(MyTime, "nn") & ":" & Format$(MyTime, "ss")) <> "00:00:00" Then
        MyTime = DateAdd("s", -1, MyTime)
        lbMyTime.Caption = Format$(MyTime, "HH") & ":" & Format$(MyTime, "nn") & ":" & Format$(MyTime, "ss")
    Else
        ConvertTime
        If List1.ListIndex = List1.ListCount - 1 Then
            cmdControl_Click 4
            List1.ListIndex = 0
        Else
            List1.ListIndex = List1.ListIndex + 1
        End If
        cmdControl_Click 4
    End If
End Sub

Private Sub Timer1_Timer()
    If List1.ListCount = 0 Then
        cmdControl(2).Enabled = False
        cmdControl(3).Enabled = False
        cmdControl(4).Enabled = False
        lbPathName.Caption = "Total: " & List1.ListCount
        cmdStop_Click
    Else
        If List1.ListIndex = -1 Then
            cmdControl(2).Enabled = False
            cmdControl(4).Enabled = False
            lbPathName.Caption = "Total: " & List1.ListCount
        Else
            cmdControl(2).Enabled = True
            cmdControl(4).Enabled = True
            lbPathName.Caption = "Selected " & List1.ListIndex + 1 & "/" & List1.ListCount & "<|>" & " Pictures" & List1.List(List1.ListIndex)
        End If
        cmdControl(3).Enabled = True
    End If
End Sub

