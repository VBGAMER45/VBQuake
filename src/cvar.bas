Attribute VB_Name = "cvar"
' cvar.bas - dynamic variable tracking (cvar.c).
'
' The variables live in cvars(1 .. num_cvars) (quakedef.bas); C links each
' newly registered cvar_t in at the head of cvar_vars, so walking that list
' is walking the array from the end. A C cvar global is the index its
' Cvar_RegisterVariable returned; index 0 is the unregistered dummy.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

' index of the variable, 0 if there is none (C: NULL)
Public Function Cvar_FindVar(ByVal var_name As String) As Long
    Dim i As Long
    For i = num_cvars To 1 Step -1
        If cvars(i).name = var_name Then
            Cvar_FindVar = i
            Exit Function
        End If
    Next
    Cvar_FindVar = 0
End Function

Public Function Cvar_VariableValue(ByVal var_name As String) As Single
    Dim v As Long
    v = Cvar_FindVar(var_name)
    If v = 0 Then
        Cvar_VariableValue = 0
        Exit Function
    End If
    Cvar_VariableValue = Q_atof(cvars(v).strval)
End Function

Public Function Cvar_VariableString(ByVal var_name As String) As String
    Dim v As Long
    v = Cvar_FindVar(var_name)
    If v = 0 Then
        Cvar_VariableString = ""
        Exit Function
    End If
    Cvar_VariableString = cvars(v).strval
End Function

' "" if nothing matches (C: NULL)
Public Function Cvar_CompleteVariable(ByVal partial As String) As String
    Dim i As Long, l As Long

    l = Len(partial)
    If l = 0 Then Exit Function

    ' check functions
    For i = num_cvars To 1 Step -1
        If Q_strncmp(partial, cvars(i).name, l) = 0 Then
            Cvar_CompleteVariable = cvars(i).name
            Exit Function
        End If
    Next
End Function

Public Sub Cvar_Set(ByVal var_name As String, ByVal value As String)
    Dim v As Long, changed As Boolean

    v = Cvar_FindVar(var_name)
    If v = 0 Then
        ' there is an error in C code if this happens
        Con_Printf "Cvar_Set: variable %s not found\n", var_name
        Exit Sub
    End If

    changed = (cvars(v).strval <> value)

    cvars(v).strval = value
    cvars(v).value = Q_atof(value)

    If cvars(v).server And changed Then
        If sv.active Then SV_BroadcastPrintf """%s"" changed to ""%s""\n", cvars(v).name, cvars(v).strval
    End If
End Sub

Public Sub Cvar_SetValue(ByVal var_name As String, ByVal value As Single)
    Cvar_Set var_name, va("%f", value)
End Sub

' Adds a freestanding variable to the variable list and returns its index
' (C: Cvar_RegisterVariable (&variable) with variable = {name, string, archive, server}).
Public Function Cvar_RegisterVariable(ByVal name As String, ByVal strval As String, _
        Optional ByVal archive As Boolean = False, Optional ByVal server As Boolean = False) As Long

    ' first check to see if it has allready been defined
    If Cvar_FindVar(name) Then
        Con_Printf "Can't register variable %s, allready defined\n", name
        Exit Function
    End If

    ' check for overlap with a command
    If Cmd_Exists(name) Then
        Con_Printf "Cvar_RegisterVariable: %s is a command\n", name
        Exit Function
    End If

    ' link the variable in
    num_cvars = num_cvars + 1
    ReDim Preserve cvars(0 To num_cvars)
    cvars(num_cvars).name = name
    cvars(num_cvars).strval = strval
    cvars(num_cvars).archive = archive
    cvars(num_cvars).server = server
    cvars(num_cvars).value = Q_atof(strval)
    Cvar_RegisterVariable = num_cvars
End Function

' Handles variable inspection and changing from the console
Public Function Cvar_Command() As Boolean
    Dim v As Long

    ' check variables
    v = Cvar_FindVar(Cmd_Argv(0))
    If v = 0 Then
        Cvar_Command = False
        Exit Function
    End If

    ' perform a variable print or set
    If Cmd_Argc() = 1 Then
        Con_Printf """%s"" is ""%s""\n", cvars(v).name, cvars(v).strval
        Cvar_Command = True
        Exit Function
    End If

    Cvar_Set cvars(v).name, Cmd_Argv(1)
    Cvar_Command = True
End Function

' Writes lines containing "set variable value" for all variables
' with the archive flag set to true (f: an open VB file number).
Public Sub Cvar_WriteVariables(ByVal f As Integer)
    Dim i As Long
    For i = num_cvars To 1 Step -1
        If cvars(i).archive Then Print #f, va("%s ""%s""", cvars(i).name, cvars(i).strval)
    Next
End Sub
