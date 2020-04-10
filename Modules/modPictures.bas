Attribute VB_Name = "modPictures"
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Const STRETCH_HALFTONE As Long = &H4&

Public Function FitPicture(PicIn As PictureBox, PicOut As PictureBox)
    Dim PicSrcWidth As Single
    Dim PicSrcHeight As Single
    Dim PicDesWidth As Single
    Dim PicDesHeight As Single
    PicIn.AutoRedraw = True
    PicIn.ScaleMode = vbPixels

    PicOut.AutoRedraw = True
    PicOut.ScaleMode = vbPixels
    PicOut.BackColor = GetSysColor(COLOR_BACKGROUND)
    PicOut.Cls

    PicSrcWidth = PicOut.ScaleX(PicIn.Width, vbTwips, PicOut.ScaleMode)
    PicSrcHeight = PicOut.ScaleY(PicIn.Height, vbTwips, PicOut.ScaleMode)

    If PicSrcWidth > PicOut.ScaleWidth Then
        PicDesWidth = PicOut.ScaleWidth
        PicSrcHeight = PicSrcHeight * (PicDesWidth / PicSrcWidth)
    Else
        PicDesHeight = PicOut.ScaleHeight
        PicDesWidth = PicSrcWidth
    End If

    If PicSrcHeight > PicOut.ScaleHeight Then
        PicDesHeight = PicOut.ScaleHeight
        PicDesWidth = PicDesWidth * (PicDesHeight / PicSrcHeight)
    Else
        PicDesHeight = PicSrcHeight
    End If

    SetStretchBltMode PicOut.hdc, STRETCH_HALFTONE
    StretchBlt PicOut.hdc, (PicOut.ScaleWidth - PicDesWidth) / 2, (PicOut.ScaleHeight - PicDesHeight) / 2, PicDesWidth, PicDesHeight, PicIn.hdc, 0, 0, PicIn.ScaleWidth, PicIn.ScaleHeight, vbSrcCopy
End Function

