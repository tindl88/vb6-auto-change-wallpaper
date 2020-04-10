VERSION 5.00
Begin VB.Form frmSlideShow 
   BorderStyle     =   0  'None
   Caption         =   "Slide Show"
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   8715
      Left            =   90
      ScaleHeight     =   581
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   705
      TabIndex        =   0
      Top             =   90
      Width           =   10575
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   75
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   90
      Top             =   1290
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   90
      Top             =   840
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   90
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2040
      Left            =   570
      ScaleHeight     =   136
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   90
      Top             =   1710
   End
End
Attribute VB_Name = "frmSlideShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Picture1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    FitPicture Picture2, Picture1
End Sub

Private Sub List1_Click()
    If Right(List1.List(List1.ListIndex), 4) = ".pcx" Then
        PCX.LoadPCX List1.List(List1.ListIndex)
        PCX.DrawPCX Picture2
    ElseIf Right(List1.List(List1.ListIndex), 4) = ".tga" Then
        TGA.LoadTGA List1.List(List1.ListIndex)
        TGA.DrawTGA Picture2
    ElseIf Right(List1.List(List1.ListIndex), 4) = ".png" Then
        PNG.DrawPNG = Picture2
        PNG.LoadPNG List1.List(List1.ListIndex)
    ElseIf Right(List1.List(List1.ListIndex), 4) = ".tif" Then
        TIF.LoadTIFF List1.List(List1.ListIndex)
    Else
        Picture2.Picture = LoadPicture(List1.List(List1.ListIndex))
    End If
    FitPicture Picture2, Picture1
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 0 To frmMain.List1.ListCount - 1
        List1.AddItem frmMain.List1.List(i)
    Next i
    
    If frmMain.chkOnTop.Value = 1 Then Call OnTop(Me, True)

    If frmSlideOption.Option1.Value = True Then
        List1.ListIndex = 0
        Timer1.Interval = Val(frmSlideOption.Combo1.Text) * 1000
        Timer1.Enabled = True
    Else
        Randomize
        List1.ListIndex = Rnd * (List1.ListCount - 1)

        Timer2.Interval = Val(frmSlideOption.Combo1.Text) * 1000
        Timer2.Enabled = True
    End If
End Sub

Private Sub Picture1_DblClick()
    Unload Me
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    
    If frmSlideOption.Option1.Value = True Then
        If KeyCode = vbKeySpace Then
            Timer1.Enabled = Not Timer1.Enabled
            Label1.Caption = IIf(Timer1.Enabled, "", "Paused...")
        End If
    Else
        If KeyCode = vbKeySpace Then
            Timer2.Enabled = Not Timer2.Enabled
            Label1.Caption = IIf(Timer2.Enabled, "", "Paused...")
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    If frmSlideOption.Check1.Value = 1 Then
        If List1.ListIndex = List1.ListCount - 1 Then
            List1.ListIndex = 0
        Else
            List1.ListIndex = List1.ListIndex + 1
        End If
    Else
        If List1.ListIndex = List1.ListCount - 1 Then
            Timer1.Enabled = False
            Unload frmSlideShow
        Else
            List1.ListIndex = List1.ListIndex + 1
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    Randomize
    List1.ListIndex = Rnd * (List1.ListCount - 1)
End Sub

Private Sub Timer3_Timer()
    Picture1.SetFocus
End Sub
