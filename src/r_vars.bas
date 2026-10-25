Attribute VB_Name = "r_vars"
' r_vars.bas - global refresh variables (r_vars.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' all global and static refresh variables are collected in a contiguous block
' to avoid cache conflicts.

'-------------------------------------------------------
' global refresh variables
'-------------------------------------------------------

' FIXME: make into one big structure, like cl or sv
' FIXME: do separately for refresh engine and driver

Public r_bmodelactive As Long
