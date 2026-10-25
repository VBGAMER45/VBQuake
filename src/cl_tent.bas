Attribute VB_Name = "cl_tent"
' cl_tent.bas - client side temporary entities (cl_tent.c)
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private num_temp_entities As Long
' (cl_temp_entities[] are cl_entities(CL_TEMP_BASE ..), cl_beams() is in quakedef.bas)

' sfx handles (cluster F), 0 = NULL
Private cl_sfx_wizhit As Long
Private cl_sfx_knighthit As Long
Private cl_sfx_tink1 As Long
Private cl_sfx_ric1 As Long
Private cl_sfx_ric2 As Long
Private cl_sfx_ric3 As Long
Private cl_sfx_r_exp3 As Long

Public Sub CL_InitTEnts()
    cl_sfx_wizhit = S_PrecacheSound("wizard/hit.wav")
    cl_sfx_knighthit = S_PrecacheSound("hknight/hit.wav")
    cl_sfx_tink1 = S_PrecacheSound("weapons/tink1.wav")
    cl_sfx_ric1 = S_PrecacheSound("weapons/ric1.wav")
    cl_sfx_ric2 = S_PrecacheSound("weapons/ric2.wav")
    cl_sfx_ric3 = S_PrecacheSound("weapons/ric3.wav")
    cl_sfx_r_exp3 = S_PrecacheSound("weapons/r_exp3.wav")
End Sub

Private Sub CL_ParseBeam(ByVal m As Long)
    Dim ent As Long, start As vec3_t, end_ As vec3_t, b As Long

    ent = MSG_ReadShort()

    start.v(0) = MSG_ReadCoord()
    start.v(1) = MSG_ReadCoord()
    start.v(2) = MSG_ReadCoord()

    end_.v(0) = MSG_ReadCoord()
    end_.v(1) = MSG_ReadCoord()
    end_.v(2) = MSG_ReadCoord()

    ' override any beam with the same entity
    For b = 0 To MAX_BEAMS - 1
        If cl_beams(b).entity = ent Then
            cl_beams(b).entity = ent
            cl_beams(b).model = m
            cl_beams(b).endtime = cl.time + 0.2
            cl_beams(b).start = start
            cl_beams(b).end_ = end_
            Exit Sub
        End If
    Next

    ' find a free beam
    For b = 0 To MAX_BEAMS - 1
        If cl_beams(b).model = 0 Or cl_beams(b).endtime < cl.time Then
            cl_beams(b).entity = ent
            cl_beams(b).model = m
            cl_beams(b).endtime = cl.time + 0.2
            cl_beams(b).start = start
            cl_beams(b).end_ = end_
            Exit Sub
        End If
    Next
    Con_Printf "beam list overflow!\n"
End Sub

' (QuakeVB) the ricochet sound of TE_SPIKE and TE_SUPERSPIKE
Private Sub SpikeSound(pos As vec3_t)
    Dim rnd As Long

    If rand() Mod 5 Then
        S_StartSound -1, 0, cl_sfx_tink1, pos, 1, 1
    Else
        rnd = rand() And 3
        If rnd = 1 Then
            S_StartSound -1, 0, cl_sfx_ric1, pos, 1, 1
        ElseIf rnd = 2 Then
            S_StartSound -1, 0, cl_sfx_ric2, pos, 1, 1
        Else
            S_StartSound -1, 0, cl_sfx_ric3, pos, 1, 1
        End If
    End If
End Sub

Public Sub CL_ParseTEnt()
    Dim typ As Long, pos As vec3_t, dl As Long, colorStart As Long, colorLength As Long

    typ = MSG_ReadByte()
    Select Case typ
    Case TE_WIZSPIKE                     ' spike hitting wall
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        R_RunParticleEffect pos, vec3_origin, 20, 30
        S_StartSound -1, 0, cl_sfx_wizhit, pos, 1, 1

    Case TE_KNIGHTSPIKE                  ' spike hitting wall
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        R_RunParticleEffect pos, vec3_origin, 226, 20
        S_StartSound -1, 0, cl_sfx_knighthit, pos, 1, 1

    Case TE_SPIKE                        ' spike hitting wall
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        R_RunParticleEffect pos, vec3_origin, 0, 10
        SpikeSound pos

    Case TE_SUPERSPIKE                   ' super spike hitting wall
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        R_RunParticleEffect pos, vec3_origin, 0, 20
        SpikeSound pos

    Case TE_GUNSHOT                      ' bullet hitting wall
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        R_RunParticleEffect pos, vec3_origin, 0, 20

    Case TE_EXPLOSION                    ' rocket explosion
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        R_ParticleExplosion pos
        dl = CL_AllocDlight(0)
        cl_dlights(dl).origin = pos
        cl_dlights(dl).radius = 350
        cl_dlights(dl).die = cl.time + 0.5
        cl_dlights(dl).decay = 300
        S_StartSound -1, 0, cl_sfx_r_exp3, pos, 1, 1

    Case TE_TAREXPLOSION                 ' tarbaby explosion
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        R_BlobExplosion pos

        S_StartSound -1, 0, cl_sfx_r_exp3, pos, 1, 1

    Case TE_LIGHTNING1                   ' lightning bolts
        CL_ParseBeam Mod_ForName("progs/bolt.mdl", True)

    Case TE_LIGHTNING2                   ' lightning bolts
        CL_ParseBeam Mod_ForName("progs/bolt2.mdl", True)

    Case TE_LIGHTNING3                   ' lightning bolts
        CL_ParseBeam Mod_ForName("progs/bolt3.mdl", True)

    ' PGM 01/21/97
    Case TE_BEAM                         ' grappling hook beam
        CL_ParseBeam Mod_ForName("progs/beam.mdl", True)
    ' PGM 01/21/97

    Case TE_LAVASPLASH
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        R_LavaSplash pos

    Case TE_TELEPORT
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        R_TeleportSplash pos

    Case TE_EXPLOSION2                   ' color mapped explosion
        pos.v(0) = MSG_ReadCoord()
        pos.v(1) = MSG_ReadCoord()
        pos.v(2) = MSG_ReadCoord()
        colorStart = MSG_ReadByte()
        colorLength = MSG_ReadByte()
        R_ParticleExplosion2 pos, colorStart, colorLength
        dl = CL_AllocDlight(0)
        cl_dlights(dl).origin = pos
        cl_dlights(dl).radius = 350
        cl_dlights(dl).die = cl.time + 0.5
        cl_dlights(dl).decay = 300
        S_StartSound -1, 0, cl_sfx_r_exp3, pos, 1, 1

    Case Else
        Sys_Error "CL_ParseTEnt: bad type"
    End Select
End Sub

' returns a cl_entities() index, -1 = NULL
Private Function CL_NewTempEntity() As Long
    Dim ent As Long, blank As entity_t

    CL_NewTempEntity = -1
    If cl_numvisedicts = MAX_VISEDICTS Then Exit Function
    If num_temp_entities = MAX_TEMP_ENTITIES Then Exit Function
    ent = CL_TEMP_BASE + num_temp_entities
    cl_entities(ent) = blank
    num_temp_entities = num_temp_entities + 1
    cl_visedicts(cl_numvisedicts) = ent
    cl_numvisedicts = cl_numvisedicts + 1

    cl_entities(ent).colormap = 0        ' vid.colormap
    CL_NewTempEntity = ent
End Function

Public Sub CL_UpdateTEnts()
    Dim i As Long, b As Long, dist As vec3_t, org As vec3_t, d As Single, ent As Long
    Dim yaw_ As Single, pitch_ As Single, forward As Single

    num_temp_entities = 0

    ' update lightning
    ' (QuakeVB: i and b are separate counters, as in C's
    ' for (i=0, b=cl_beams ; i< MAX_BEAMS ; i++, b++), whose inner loop below
    ' reuses i: after a beam that made entities, i is 3, so the loop ends
    ' before it gets to the last beams)
    i = 0
    b = 0
    Do While i < MAX_BEAMS
        If cl_beams(b).model = 0 Or cl_beams(b).endtime < cl.time Then GoTo nextbeam

        ' if coming from the player, update the start position
        If cl_beams(b).entity = cl.viewentity Then cl_beams(b).start = cl_entities(cl.viewentity).origin

        ' calculate pitch and yaw
        VectorSubtract cl_beams(b).end_, cl_beams(b).start, dist

        If dist.v(1) = 0 And dist.v(0) = 0 Then
            yaw_ = 0
            If dist.v(2) > 0 Then
                pitch_ = 90
            Else
                pitch_ = 270
            End If
        Else
            yaw_ = ToInt(atan2(dist.v(1), dist.v(0)) * 180 / M_PI)
            If yaw_ < 0 Then yaw_ = yaw_ + 360

            forward = Sqr(dist.v(0) * dist.v(0) + dist.v(1) * dist.v(1))
            pitch_ = ToInt(atan2(dist.v(2), forward) * 180 / M_PI)
            If pitch_ < 0 Then pitch_ = pitch_ + 360
        End If

        ' add new entities for the lightning
        org = cl_beams(b).start
        d = VectorNormalize(dist)
        Do While d > 0
            ent = CL_NewTempEntity()
            If ent = -1 Then Exit Sub
            cl_entities(ent).origin = org
            cl_entities(ent).model = cl_beams(b).model
            cl_entities(ent).angles.v(0) = pitch_
            cl_entities(ent).angles.v(1) = yaw_
            cl_entities(ent).angles.v(2) = rand() Mod 360

            For i = 0 To 2
                org.v(i) = org.v(i) + dist.v(i) * 30
            Next
            d = d - 30
        Loop
nextbeam:
        i = i + 1
        b = b + 1
    Loop
End Sub
