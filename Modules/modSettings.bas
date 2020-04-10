Attribute VB_Name = "modSettings"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Dim sFile As String
Dim I As Integer

Private Function ReadINI(Filename As String, Section As String, Key As String, Default As Variant) As String
Dim strBuf As String * 255
Dim L As Long
    L = GetPrivateProfileString(Section, Key, Default, strBuf, 255, Filename)
    ReadINI = Left$(strBuf, L)
End Function

Private Sub WriteINI(Filename As String, Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, Filename
End Sub


Public Sub LoadSettingsINI()
    sFile = Environ$("WinDir") & "\TinSettings.ini"

    Dim Tmp As String
        Do
        Tmp = ReadINI(sFile, "LastList", CLng(I), "")
            If Tmp <> "" Then
                I = I + 1
                frmMain.List1.AddItem Tmp
            End If
        Loop Until Tmp = ""
    
    ListBoxHScroll frmMain.List1
    frmMain.cbTime(0).ListIndex = ReadINI(sFile, "Settings", "Hour", 0)
    frmMain.cbTime(1).ListIndex = ReadINI(sFile, "Settings", "Min", 15)
    frmMain.cbTime(2).ListIndex = ReadINI(sFile, "Settings", "Sec", 0)
    frmMain.OptStyle(0).Value = ReadINI(sFile, "Settings", "Stretch", 1)
    frmMain.OptStyle(1).Value = ReadINI(sFile, "Settings", "Center", 0)
    frmMain.OptStyle(2).Value = ReadINI(sFile, "Settings", "Tile", 0)
    frmMain.chkOnTop.Value = ReadINI(sFile, "Settings", "OnTop", 0)
    frmMain.chkStartMode.Value = ReadINI(sFile, "Settings", "StartMode", 0)
    If ReadINI(sFile, "Settings", "AutoStart", 0) = "Play" Then frmMain.cmdStart_Click
    frmMain.chkStartWin.Value = ReadINI(sFile, "Settings", "Run", 0)
    frmMain.List1.ListIndex = ReadINI(sFile, "Settings", "PicIndex", "")
    frmMain.chkBeep = ReadINI(sFile, "Settings", "Beep", 0)
End Sub

Public Sub SaveSettingsINI()
On Error Resume Next
    sFile = Environ$("WinDir") & "\TinSettings.ini"
    Kill sFile
    
    For I = 0 To frmMain.List1.ListCount - 1
        WriteINI sFile, "LastList", CLng(I), frmMain.List1.List(I)
    Next
    
    WriteINI sFile, "Settings", "Hour", frmMain.cbTime(0).ListIndex
    WriteINI sFile, "Settings", "Min", frmMain.cbTime(1).ListIndex
    WriteINI sFile, "Settings", "Sec", frmMain.cbTime(2).ListIndex
    WriteINI sFile, "Settings", "Stretch", frmMain.OptStyle(0).Value
    WriteINI sFile, "Settings", "Center", frmMain.OptStyle(1).Value
    WriteINI sFile, "Settings", "Tile", frmMain.OptStyle(2).Value
    WriteINI sFile, "Settings", "OnTop", frmMain.chkOnTop.Value
    WriteINI sFile, "Settings", "StartMode", frmMain.chkStartMode.Value
    WriteINI sFile, "Settings", "AutoStart", frmMain.cmdStart.Tag
    WriteINI sFile, "Settings", "Run", frmMain.chkStartWin.Value
    WriteINI sFile, "Settings", "PicIndex", frmMain.List1.ListIndex
    WriteINI sFile, "Settings", "Beep", frmMain.chkBeep.Value
End Sub

