Attribute VB_Name = "crc"
' crc.bas - 16 bit, non-reflected CRC using the polynomial 0x1021 and the
' initial and final xor values shown below; in other words, the CCITT
' standard CRC used by XMODEM (crc.c). The table is computed on first use
' (it equals crc.c's crctable entry for entry).
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const CRC_INIT_VALUE As Long = &HFFFF&
Private Const CRC_XOR_VALUE As Long = &H0

Private crctable(0 To 255) As Long
Private tablemade As Boolean

Private Sub MakeTable()
    Dim i As Long, j As Long, c As Long
    For i = 0 To 255
        c = i * 256
        For j = 1 To 8
            If c And &H8000& Then
                c = ((c * 2) Xor &H1021&) And &HFFFF&
            Else
                c = (c * 2) And &HFFFF&
            End If
        Next
        crctable(i) = c
    Next
    tablemade = True
End Sub

Public Sub CRC_Init(crcvalue As Long)
    If Not tablemade Then MakeTable
    crcvalue = CRC_INIT_VALUE
End Sub

Public Sub CRC_ProcessByte(crcvalue As Long, ByVal data As Long)
    crcvalue = ((crcvalue * 256) And &HFFFF&) Xor crctable(((crcvalue \ 256) Xor data) And 255)
End Sub

Public Function CRC_Value(ByVal crcvalue As Long) As Long
    CRC_Value = crcvalue Xor CRC_XOR_VALUE
End Function
