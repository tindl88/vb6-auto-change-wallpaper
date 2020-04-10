VERSION 5.00
Begin VB.Form frmCalendar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   3285
   ClientLeft      =   9525
   ClientTop       =   1830
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4320
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call MoveForm(Me, FrmTopRight)
End Sub
