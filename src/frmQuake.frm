VERSION 5.00
Begin VB.Form frmQuake 
   AutoRedraw      =   0   'False
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QuakeVB"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9600
   LinkTopic       =   "frmQuake"
   MaxButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmQuake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmQuake - the game window. Input is handled by QuakeWndProc (vid_win.bas).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Sub Form_Paint()
    VID_RepaintWindow
End Sub

Private Sub Form_Resize()
    VID_RepaintWindow
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        VID_RequestQuit
        Cancel = 1
    End If
End Sub
