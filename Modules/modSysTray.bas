Attribute VB_Name = "modSysTray"
Option Explicit

Private Const MAX_TOOLTIP As Integer = 64
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201

Private Type NOTIFYICONDATA
    cbSize           As Long
    hWnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
End Type
Private nfIconData As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Sub AddTrayIcon(Frm As Form, ByVal sText As String)
'Dua Icon xuông' Systray
'VD: AddTrayIcon frmMain, "Text"
    With nfIconData
        .hWnd = Frm.hWnd
        .uID = Frm.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Frm.Icon.Handle
        .szTip = sText & vbNullChar
        .cbSize = Len(nfIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
End Sub

Public Sub RemoveTrayIcon()
'Xóa Icon o? Systray
'VD: RemoveTrayIcon
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub

