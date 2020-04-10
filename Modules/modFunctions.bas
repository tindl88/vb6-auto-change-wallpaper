Attribute VB_Name = "modFunctions"
Option Explicit
'Move Form and Set Wallpaper
Public Enum WallType
    WallStretch
    WallCenter
    WallTile
End Enum

Public Enum FormPosition
    FrmTopLeft
    FrmTopRight
    FrmCenter
    FrmBottomLeft
    FrmBottomRight
End Enum

Public Type RECT
    Left                                        As Long
    Top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SETDESKWALLPAPER              As Long = 20
Private Const SPIF_SENDWININICHANGE             As Long = &H2
Private Const SPIF_UPDATEINIFILE                As Long = &H1
Private Const SPI_GETWORKAREA                   As Long = 48
'Set On Top
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST                      As Long = -1
Private Const HWND_NOTOPMOST                    As Long = -2
Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_SHOWWINDOW                    As Long = &H40
'Form Drag
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN                  As Long = &HA1
Private Const HTCAPTION                         As Long = 2
'Add HScrollListbox
Private Const LB_SETHORIZONTALEXTENT            As Long = &H194
'Play Sound
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ASYNC                         As Long = &H1
Private Const SND_NODEFAULT                     As Long = &H2
'Color Dialog
Private Type ChooseColor
    lStructSize                                 As Long
    hwndOwner                                   As Long
    hInstance                                   As Long
    rgbResult                                   As Long
    lpCustColors                                As String
    flags                                       As Long
    lCustData                                   As Long
    lpfnHook                                    As Long
    lpTemplateName                              As String
End Type
Private Const CC_FULLOPEN                       As Long = &H2
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
'Desktop Color
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Const COLOR_BACKGROUND               As Long = 1
'Icon
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Public Function ShowColor(Frm As Form, sPic As PictureBox)
'Hiên? thi. hôp. chon. màu
'VD: Call ShowColor(FrmMain, Picture1)
    Dim CC As ChooseColor
        CC.lStructSize = Len(CC)
        CC.hwndOwner = Frm.hWnd
        CC.flags = CC_FULLOPEN
        CC.lCustData = 0
        CC.lpCustColors = String$(256, vbNullChar)
        If ChooseColor(CC) Then
             sPic.BackColor = CC.rgbResult
        Else
            Exit Function
        End If
End Function

Public Function PlaySnd(ByVal sFile As String) As String
'Phát môt. file nhac.
'VD: Call PlaySnd("C:\Tin.wav")
    sndPlaySound sFile, SND_ASYNC Or SND_NODEFAULT
End Function

Public Function OnTop(Frm As Form, Value As Boolean) As Long
'Dat. form luôn lên trên các form khác
'VD: Call OnTop(frmMain, Check1.Value)
    If Value = True Then
        SetWindowPos Frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos Frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Function

Public Function DragForm(Frm As Form) As Long
'Kéo form
'VD: Call FormDrag(frmMain)
    ReleaseCapture
    Call SendMessage(Frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Function

Public Function MoveForm(Frm As Form, sType As FormPosition) As Long
'Di chuyên? Form
'VD: Call MoveForm(frmMain, FrmBottomRight)
Dim Area As RECT
    If SystemParametersInfo(SPI_GETWORKAREA, 0, Area, 0) <> 0 Then
        Select Case sType
            Case FrmTopLeft: Frm.Move 0, 0
            Case FrmTopRight: Frm.Move Frm.ScaleX(Area.Right, vbPixels, vbTwips) - Frm.Width, 0
            Case FrmCenter: Frm.Move (Frm.ScaleX(Area.Right, vbPixels, vbTwips) - Frm.Width) \ 2, (Frm.ScaleY(Area.Bottom, vbPixels, vbTwips) - Frm.Height) \ 2
            Case FrmBottomLeft: Frm.Move 0, Frm.ScaleY(Area.Bottom, vbPixels, vbTwips) - Frm.Height
            Case FrmBottomRight: Frm.Move Frm.ScaleX(Area.Right, vbPixels, vbTwips) - Frm.Width, Frm.ScaleY(Area.Bottom, vbPixels, vbTwips) - Frm.Height
        End Select
    End If
End Function

Public Function GetFileNameFromPath(ByVal sPath As String) As String
'VD: Call GetFileNameFromPath("D:\Tin\Hinh.jpg") return "Hinh.jpg"
    GetFileNameFromPath = Mid$(sPath, InStrRev(sPath, "\") + 1)
End Function

Public Function GetExtOfFileName(ByVal sFileName As String) As String
'VD: Call GetExtOfFileName("D:\Tin\Hinh.jpg") return "jpg"
    GetExtOfFileName = Mid$(sFileName, InStrRev(sFileName, ".") + 1)
End Function

Public Function GetFolderPathFromPath(ByVal sPath As String) As String
'VD: Call GetFolderPathFromPath("D:\Tin\Hinh.jpg") return "D:\Tin"
    GetFolderPathFromPath = Left$(sPath, InStrRev(sPath, "\") - 1)
End Function

Public Function GetNameOfFileName(ByVal sFileName As String) As String
'VD: Call GetNameOfFileName("Hinh.jpg") return "Hinh"
    GetNameOfFileName = Left$(sFileName, InStrRev(sFileName, ".") - 1)
End Function

Public Function ConvertBytes(ByVal Bytes As Double, Optional sFormat As String = "0.0") As String
On Error GoTo Hell
    If Bytes >= 1024 ^ 3 Then
        ConvertBytes = Format$(Bytes / 1024 ^ 3, sFormat) & " GB"
    ElseIf Bytes >= 1024 ^ 2 Then
        ConvertBytes = Format$(Bytes / 1024 ^ 2, sFormat) & " MB"
    ElseIf Bytes >= 1024 Then
        ConvertBytes = Format$(Bytes / 1024, sFormat) & " KB"
    ElseIf Bytes < 1024 Then
        ConvertBytes = Bytes & " Bytes"
    End If
    Exit Function
Hell:
    ConvertBytes = "0 Bytes"
End Function

Public Function RegCreateKey(ByVal FullPathAndKey As String, ByVal Value As String) As String
'Tao. key trong registry
'VD: RegCreateKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AC Wallpaper", App.Path & "\" & App.EXEName & ".exe"
On Error Resume Next
Dim Obj As Object
Set Obj = CreateObject("Wscript.Shell")
    Obj.RegWrite FullPathAndKey, Value
End Function

Public Function RegDeleteKey(ByVal FullPathAndKey As String) As Long
'Xóa key trong registry
'VD: RegDeleteKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AC Wallpaper"
On Error Resume Next
Dim Obj As Object
Set Obj = CreateObject("Wscript.Shell")
    Obj.RegDelete FullPathAndKey
End Function

Public Function RegReadKey(ByVal FullPathAndKey As String) As Long
'Doc. key trong registry
'VD: Label1.Caption = RegReadKey("HKCU\Software\Microsoft\Windows\CurrentVersion\Run\AC Wallpaper")
On Error Resume Next
Dim Obj As Object
Set Obj = CreateObject("Wscript.Shell")
    RegReadKey = Obj.RegRead(FullPathAndKey)
End Function

Public Function SetWallpaper(sPic As PictureBox, Optional sWallTypes As WallType) As String
'Dat. hình ra Desktop
'VD: Call SetWallpaper(Picture1, WallStretch)
    SavePicture sPic.Picture, Environ$("WinDir") & "\Tin Wallpaper.bmp"
    'Các kiêu? hiên? thi. cua? hình
    Select Case sWallTypes
        Case WallStretch
            RegCreateKey "HKCU\Control Panel\Desktop\TileWallpaper", "0"
            RegCreateKey "HKCU\Control Panel\Desktop\WallpaperStyle", "2"
        Case WallCenter
            RegCreateKey "HKCU\Control Panel\Desktop\TileWallpaper", "0"
            RegCreateKey "HKCU\Control Panel\Desktop\WallpaperStyle", "0"
        Case WallTile
            RegCreateKey "HKCU\Control Panel\Desktop\TileWallpaper", "1"
            RegCreateKey "HKCU\Control Panel\Desktop\WallpaperStyle", "0"
    End Select
    'Dat. Wallpaper ra Desktop
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, ByVal Environ$("WinDir") & "\Tin Wallpaper.bmp", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
End Function

Public Function GetIcon(sPic As PictureBox, ByVal sFile As String, ByRef IconIndex As Long) As Long
'Lây' Icon cua? file
'VD: GetIcon Picture1, "C:\Tin.jpg", 1
Dim sIcon As Long
    sPic.Cls
    sIcon = ExtractAssociatedIcon(App.hInstance, sFile, IconIndex)
    DrawIcon sPic.hdc, 0, 0, sIcon
    DestroyIcon sIcon
End Function

Public Function ObjExist(ByVal sObject As String) As Boolean
'Kiêm? tra su. tôn` tai. cua? file hoac. thu muc.
'File: ObjExist "C:\Tin.txt"
'Folder: ObjExist "C:\Tin"
    ObjExist = Not (Dir$(sObject, vbHidden Or vbSystem) = "")
End Function


Public Sub ListBoxHScroll(ctlListBox As ListBox)
'Thêm thanh cuôn. ngang cho listbox
'VD: Call ListBoxHScroll(List1)
    Dim i As Long
    Dim intGreatestLen As Long
    Dim lngGreatestWidth As Long
    
    For i = 0 To ctlListBox.ListCount - 1
        If Len(ctlListBox.List(i)) > Len(ctlListBox.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
        
    lngGreatestWidth = ctlListBox.Parent.TextWidth(ctlListBox.List(intGreatestLen) & Space(1))
    lngGreatestWidth = ctlListBox.Parent.ScaleX(lngGreatestWidth, ctlListBox.Parent.ScaleMode, vbPixels)
    SendMessage ctlListBox.hWnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
End Sub

Public Function RemoveDuplicates(ctrListBox As ListBox)
'Nêu' trong listbox có nhung~ dòng trùng nhau thì se~ xóa và chi? chu`a 1 dòng
    Dim i As Long: Dim j As Long
    Dim ITem As String
    For i = 0 To ctrListBox.ListCount
        ITem = ctrListBox.List(i)
        For j = i + 1 To ctrListBox.ListCount
            If ITem = ctrListBox.List(j) Then
                ctrListBox.RemoveItem j
                j = j - 1
            End If
        Next
    Next
End Function

Public Function RemoveDeadEntries(ctrListBox As ListBox)
'Nêu' trong listbox có path cua? file mà file do' không tôn` tai. thì xóa dòng do'
    Dim i As Long
    Dim j As Long: j = ctrListBox.ListCount
    For i = 0 To j - 1
        If ObjExist(ctrListBox.List(i)) = False Then
            ctrListBox.RemoveItem i
            i = i - 1
            j = j - 1
        End If
    Next
End Function
