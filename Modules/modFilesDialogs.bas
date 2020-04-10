Attribute VB_Name = "modFilesDialogs"
Option Explicit
'////////////////////////////////////////BROWSE FOLDER//////////////////////////////

Private Const BIF_DONTGOBELOWDOMAIN     As Long = &H2
Private Const BIF_RETURNONLYFSDIRS      As Long = &H1
Private Const BIF_STATUSTEXT            As Long = &H4
Private Const BIF_USENEWUI              As Long = &H40
Private Const MAX_PATH                  As Long = 260

Private Const WM_USER                   As Long = &H400
Private Const BFFM_INITIALIZED          As Long = 1
Private Const BFFM_SELCHANGED           As Long = 2
Private Const BFFM_SETSTATUSTEXT        As Long = (WM_USER + 100)
Private Const BFFM_SETSELECTION         As Long = (WM_USER + 102)
Private Type BrowseInfo
  hWndFrm                               As Long
  pIDLRoot                              As Long
  pszDisplayName                        As Long
  lpszTitle                             As Long
  ulFlags                               As Long
  lpfnCallback                          As Long
  lParam                                As Long
  iImage                                As Long
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private CurrentDir                      As String
'//////////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////COMMOMDIALOG//////////////////////////////
Private Type OPENFILENAME
    lStructSize                         As Long
    hWndFrm                             As Long
    hInstance                           As Long
    lpstrFilter                         As String
    lpstrCustomFilter                   As String
    nMaxCustFilter                      As Long
    nFilterIndex                        As Long
    lpstrFile                           As String
    nMaxFile                            As Long
    lpstrFileTitle                      As String
    nMaxFileTitle                       As Long
    lpstrInitialDir                     As String
    lpstrTitle                          As String
    flags                               As Long
    nFileOffset                         As Integer
    nFileExtension                      As Integer
    lpstrDefExt                         As String
    lCustData                           As Long
    lpfnHook                            As Long
    lpTemplateName                      As String
End Type

Private Const OFN_ALLOWMULTISELECT      As Long = &H200
Private Const OFN_EXPLORER              As Long = &H80000
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'////////////////////////////////////////PROPERTIES FILE//////////////////////////////
Type SHELLEXECUTEINFO
   cbSize As Long
   fMask As Long
   hWnd As Long
   lpVerb As String
   lpFile As String
   lpParameters As String
   lpDirectory As String
   nShow As Long
   hInstApp As Long
   lpIDList As Long
   lpClass As String
   hkeyClass As Long
   dwHotKey As Long
   hIcon As Long
   hProcess As Long
End Type
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Public Sub ShowPropeties(Filename As String, OwnerhWnd As Long)
   Dim SEI As SHELLEXECUTEINFO
   Dim R As Long
   With SEI
      .cbSize = Len(SEI)
      .fMask = SEE_MASK_NOCLOSEPROCESS Or _
      SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
      .hWnd = OwnerhWnd
      .lpVerb = "properties"
      .lpFile = Filename
      .lpParameters = vbNullChar
      .lpDirectory = vbNullChar
      .nShow = 0
      .hInstApp = 0
      .lpIDList = 0
   End With
   R = ShellExecuteEX(SEI)
End Sub
'////////////////////////////////////////PROPERTIES FILE//////////////////////////////

'////////////////////////////////////////COMMOMDIALOG//////////////////////////////
Public Function AddFiles(sDir As String) As String
Dim OFN As OPENFILENAME
    OFN.lStructSize = Len(OFN)
    OFN.hWndFrm = frmMain.hWnd
    OFN.hInstance = App.hInstance
    OFN.lpstrFilter = "Pictures(*.jpg;*.jpeg;*.gif;*.bmp;*.wmf;*.dib;*.pcx;*.tga;*.png;*.tif;*.ico;*.cur)" & Chr$(0) + "*.jpg;*.jpeg;*.gif;*.bmp;*.wmf;*.dib;*.pcx;*.tga;*.png;*.tif;*.ico;*.cur" & Chr$(0) + "All Files (*.*)" & Chr$(0) + "*.*" & Chr$(0)
    OFN.lpstrFile = Space$(99999)
    OFN.nMaxFile = 100000
    OFN.lpstrFileTitle = Space$(254)
    OFN.nMaxFileTitle = 255
    OFN.lpstrInitialDir = sDir
    OFN.lpstrTitle = "Open Pictures"
    OFN.flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER
    
    If GetOpenFileName(OFN) Then
        Dim nFiles As Variant
        Dim p As Integer
            nFiles = Split(OFN.lpstrFile, vbNullChar)
            Select Case UBound(nFiles)
                Case Is = 2
                    frmMain.List1.AddItem nFiles(0)
                Case Is > 2
                    If Right$(nFiles(0), 1) <> "\" Then nFiles(0) = nFiles(0) & "\"
                    For p = 1 To UBound(nFiles) - 2
                        frmMain.List1.AddItem nFiles(0) & nFiles(p)
                    Next
            End Select
            SaveSetting "AC Wallpaper", "Settings", "LastDir", GetFolderPathFromPath(nFiles(0))
    Else
        Exit Function
    End If
End Function
'////////////////////////////////////////COMMOMDIALOG//////////////////////////////

'////////////////////////////////////////BROWSE FOLDER//////////////////////////////
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    Dim lpIDList As Long
    Dim ret As Long
    Dim sPath As String
    On Error Resume Next
        Select Case uMsg
            Case BFFM_INITIALIZED
                Call SendMessage(hWnd, BFFM_SETSELECTION, 1, CurrentDir)
            Case BFFM_SELCHANGED
                sPath = Space(MAX_PATH)
                ret = SHGetPathFromIDList(lp, sPath)
            If ret = 1 Then
                Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sPath)
            End If
        End Select
  BrowseCallbackProc = 0
End Function

Private Function GetAddressofFunction(Add As Long) As Long
    GetAddressofFunction = Add
End Function

Public Function BrowseForFolder(Frm As Form, Title As String, sDir As String) As String
    Dim lpIDList As Long
    Dim sPath As String
    Dim tBrowseInfo As BrowseInfo
    CurrentDir = sDir & vbNullChar
    
    With tBrowseInfo
        .hWndFrm = Frm.hWnd
        .lpszTitle = lstrcat(Title, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT Or BIF_USENEWUI
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sPath = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
        
        BrowseForFolder = sPath
        SaveSetting "AC Wallpaper", "Settings", "BrowseFolder", sPath
    Else
        BrowseForFolder = ""
    End If
End Function

Public Function AddFolder(sPath) As String
On Error Resume Next
Dim i As Byte
Dim Tmp As String
Dim EXT(8) As String
    EXT(0) = UCase$("jpg")
    EXT(1) = UCase$("jpeg")
    EXT(2) = UCase$("gif")
    EXT(3) = UCase$("bmp")
    EXT(4) = UCase$("wmf")
    EXT(5) = UCase$("pcx")
    EXT(6) = UCase$("tga")
    EXT(7) = UCase$("png")
    EXT(8) = UCase$("tif")
    EXT(9) = UCase$("tif")
    EXT(10) = UCase$("dib")
    EXT(11) = UCase$("ico")
    EXT(12) = UCase$("cur")
    Tmp = Dir$(sPath & "\", vbHidden Or vbSystem Or vbArchive Or vbReadOnly)
        Do
            For i = 0 To UBound(EXT)
                If EXT(i) = UCase$(Right$(Tmp, 3)) Or EXT(i) = UCase$(Right$(Tmp, 4)) Then
                    frmMain.List1.AddItem Replace$(sPath & "\" & Tmp, "\\", "\")
            Exit For
                End If
            Next
    Tmp = Dir
        Loop Until Tmp = ""
End Function
'//////////////////////////////////////////////////////////////////////////////////////////

