Attribute VB_Name = "r_part"
' r_part.bas - particle effects (r_part.c)
'
' The particles are particles(); the particle_t pointers (active_particles,
' free_particles, next) are indices there, -1 = NULL. R_AllocParticle is
' the free-list pop that C repeats in every effect. Not ported: the QUAKE2
' and GLQUAKE parts, and the unused globals avelocity, partstep, timescale.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const MAX_PARTICLES As Long = 2048           ' default max # of particles at one
                                                     '  time
Private Const ABSOLUTE_MIN_PARTICLES As Long = 512   ' no fewer than this no matter what's
                                                     '  on the command line

Private ramp1(0 To 7) As Long
Private ramp2(0 To 7) As Long
Private ramp3(0 To 7) As Long

Public active_particles As Long, free_particles As Long

Public particles() As particle_t
Public r_numparticles As Long

Public r_pright As vec3_t, r_pup As vec3_t, r_ppn As vec3_t

Private avelocities(0 To NUMVERTEXNORMALS - 1) As vec3_t
Private Const beamlength As Single = 16

Private tracercount As Long              ' C: static of R_RocketTrail

Private Sub SetRamp(ramp() As Long, ByVal v As Variant)
    Dim i As Long

    For i = 0 To UBound(v)
        ramp(i) = v(i)
    Next
End Sub

Public Sub R_InitParticles()
    Dim i As Long

    ' (C: static initializers)
    SetRamp ramp1, Array(&H6F, &H6D, &H6B, &H69, &H67, &H65, &H63, &H61)
    SetRamp ramp2, Array(&H6F, &H6E, &H6D, &H6C, &H6B, &H6A, &H68, &H66)
    SetRamp ramp3, Array(&H6D, &H6B, 6, 5, 4, 3)

    i = COM_CheckParm("-particles")

    If i Then
        r_numparticles = Q_atoi(com_argv(i + 1))
        If r_numparticles < ABSOLUTE_MIN_PARTICLES Then r_numparticles = ABSOLUTE_MIN_PARTICLES
    Else
        r_numparticles = MAX_PARTICLES
    End If

    ReDim particles(0 To r_numparticles - 1)
End Sub

' (QuakeVB) takes a particle off the free list and puts it at the head of
' the active list; -1 if there is none
Private Function R_AllocParticle() As Long
    Dim p As Long

    If free_particles = -1 Then
        R_AllocParticle = -1
        Exit Function
    End If
    p = free_particles
    free_particles = particles(p).nxt
    particles(p).nxt = active_particles
    active_particles = p
    R_AllocParticle = p
End Function

' ent: a cl_entities() index
Public Sub R_EntityParticles(ByVal ent As Long)
    Dim count As Long, i As Long, p As Long, angle As Single
    Dim sr As Single, sp As Single, sy As Single, cr As Single, cp As Single, cy As Single
    Dim forward As vec3_t, dist As Single

    dist = 64
    count = 50

    If avelocities(0).v(0) = 0 Then
        For i = 0 To NUMVERTEXNORMALS * 3 - 1
            avelocities(i \ 3).v(i Mod 3) = (rand() And 255) * 0.01
        Next
    End If

    For i = 0 To NUMVERTEXNORMALS - 1
        angle = cl.time * avelocities(i).v(0)
        sy = Sin(angle)
        cy = Cos(angle)
        angle = cl.time * avelocities(i).v(1)
        sp = Sin(angle)
        cp = Cos(angle)
        angle = cl.time * avelocities(i).v(2)
        sr = Sin(angle)
        cr = Cos(angle)

        forward.v(0) = cp * cy
        forward.v(1) = cp * sy
        forward.v(2) = -sp

        p = R_AllocParticle()
        If p = -1 Then Exit Sub

        particles(p).die = cl.time + 0.01
        particles(p).color = &H6F
        particles(p).typ = pt_explode

        particles(p).org.v(0) = cl_entities(ent).origin.v(0) + r_avertexnormals(i, 0) * dist + forward.v(0) * beamlength
        particles(p).org.v(1) = cl_entities(ent).origin.v(1) + r_avertexnormals(i, 1) * dist + forward.v(1) * beamlength
        particles(p).org.v(2) = cl_entities(ent).origin.v(2) + r_avertexnormals(i, 2) * dist + forward.v(2) * beamlength
    Next
End Sub

Public Sub R_ClearParticles()
    Dim i As Long

    free_particles = 0
    active_particles = -1

    For i = 0 To r_numparticles - 1
        particles(i).nxt = i + 1
    Next
    particles(r_numparticles - 1).nxt = -1
End Sub

' (QuakeVB) C reads the file with fscanf ("%f %f %f\n"); here it is loaded
' whole and read a number at a time
Public Sub R_ReadPointFile_f()
    Dim buf() As Byte, org As vec3_t, c As Long, p As Long, pos As Long, j As Long, fname As String
    Dim tok As String

    fname = "maps/" & sv.name & ".pts"

    If Not COM_LoadFile(fname, buf) Then
        Con_Printf "couldn't open %s\n", fname
        Exit Sub
    End If

    Con_Printf "Reading %s...\n", fname
    c = 0
    pos = 0
    Do
        For j = 0 To 2
            tok = NextNumber(buf, pos)
            If tok = "" Then GoTo done
            org.v(j) = C_atof(tok)
        Next
        c = c + 1

        p = R_AllocParticle()
        If p = -1 Then
            Con_Printf "Not enough free particles\n"
            Exit Do
        End If

        particles(p).die = 99999
        particles(p).color = (-c) And 15
        particles(p).typ = pt_static
        particles(p).vel = vec3_origin
        particles(p).org = org
    Loop
done:

    Con_Printf "%i points read\n", c
End Sub

' (QuakeVB) the next whitespace-separated word of a zero-terminated buffer
Private Function NextNumber(buf() As Byte, pos As Long) As String
    Dim start As Long

    Do While buf(pos) <> 0 And buf(pos) <= 32
        pos = pos + 1
    Loop
    start = pos
    Do While buf(pos) > 32
        pos = pos + 1
    Loop
    If pos > start Then NextNumber = BytesToStr(buf, start, pos - start)
End Function

' Parse an effect out of the server message
Public Sub R_ParseParticleEffect()
    Dim org As vec3_t, dir_ As vec3_t, i As Long, count As Long, msgcount As Long, color As Long

    For i = 0 To 2
        org.v(i) = MSG_ReadCoord()
    Next
    For i = 0 To 2
        dir_.v(i) = MSG_ReadChar() * (1# / 16)
    Next
    msgcount = MSG_ReadByte()
    color = MSG_ReadByte()

    If msgcount = 255 Then
        count = 1024
    Else
        count = msgcount
    End If

    R_RunParticleEffect org, dir_, color, count
End Sub

Public Sub R_ParticleExplosion(org As vec3_t)
    Dim i As Long, j As Long, p As Long

    For i = 0 To 1023
        p = R_AllocParticle()
        If p = -1 Then Exit Sub

        particles(p).die = cl.time + 5
        particles(p).color = ramp1(0)
        particles(p).ramp = rand() And 3
        If i And 1 Then
            particles(p).typ = pt_explode
            For j = 0 To 2
                particles(p).org.v(j) = org.v(j) + ((rand() Mod 32) - 16)
                particles(p).vel.v(j) = (rand() Mod 512) - 256
            Next
        Else
            particles(p).typ = pt_explode2
            For j = 0 To 2
                particles(p).org.v(j) = org.v(j) + ((rand() Mod 32) - 16)
                particles(p).vel.v(j) = (rand() Mod 512) - 256
            Next
        End If
    Next
End Sub

Public Sub R_ParticleExplosion2(org As vec3_t, ByVal colorStart As Long, ByVal colorLength As Long)
    Dim i As Long, j As Long, p As Long, colorMod As Long

    colorMod = 0

    For i = 0 To 511
        p = R_AllocParticle()
        If p = -1 Then Exit Sub

        particles(p).die = cl.time + 0.3
        particles(p).color = colorStart + (colorMod Mod colorLength)
        colorMod = colorMod + 1

        particles(p).typ = pt_blob
        For j = 0 To 2
            particles(p).org.v(j) = org.v(j) + ((rand() Mod 32) - 16)
            particles(p).vel.v(j) = (rand() Mod 512) - 256
        Next
    Next
End Sub

Public Sub R_BlobExplosion(org As vec3_t)
    Dim i As Long, j As Long, p As Long

    For i = 0 To 1023
        p = R_AllocParticle()
        If p = -1 Then Exit Sub

        particles(p).die = cl.time + 1 + (rand() And 8) * 0.05

        If i And 1 Then
            particles(p).typ = pt_blob
            particles(p).color = 66 + rand() Mod 6
            For j = 0 To 2
                particles(p).org.v(j) = org.v(j) + ((rand() Mod 32) - 16)
                particles(p).vel.v(j) = (rand() Mod 512) - 256
            Next
        Else
            particles(p).typ = pt_blob2
            particles(p).color = 150 + rand() Mod 6
            For j = 0 To 2
                particles(p).org.v(j) = org.v(j) + ((rand() Mod 32) - 16)
                particles(p).vel.v(j) = (rand() Mod 512) - 256
            Next
        End If
    Next
End Sub

Public Sub R_RunParticleEffect(org As vec3_t, dir_ As vec3_t, ByVal color As Long, ByVal count As Long)
    Dim i As Long, j As Long, p As Long

    For i = 0 To count - 1
        p = R_AllocParticle()
        If p = -1 Then Exit Sub

        If count = 1024 Then
            ' rocket explosion
            particles(p).die = cl.time + 5
            particles(p).color = ramp1(0)
            particles(p).ramp = rand() And 3
            If i And 1 Then
                particles(p).typ = pt_explode
                For j = 0 To 2
                    particles(p).org.v(j) = org.v(j) + ((rand() Mod 32) - 16)
                    particles(p).vel.v(j) = (rand() Mod 512) - 256
                Next
            Else
                particles(p).typ = pt_explode2
                For j = 0 To 2
                    particles(p).org.v(j) = org.v(j) + ((rand() Mod 32) - 16)
                    particles(p).vel.v(j) = (rand() Mod 512) - 256
                Next
            End If
        Else
            particles(p).die = cl.time + 0.1 * (rand() Mod 5)
            particles(p).color = (color And Not 7) + (rand() And 7)
            particles(p).typ = pt_slowgrav
            For j = 0 To 2
                particles(p).org.v(j) = org.v(j) + ((rand() And 15) - 8)
                particles(p).vel.v(j) = dir_.v(j) * 15   ' + (rand()%300)-150;
            Next
        End If
    Next
End Sub

Public Sub R_LavaSplash(org As vec3_t)
    Dim i As Long, j As Long, k As Long, p As Long, vel As Single, dir_ As vec3_t

    For i = -16 To 15
        For j = -16 To 15
            For k = 0 To 0
                p = R_AllocParticle()
                If p = -1 Then Exit Sub

                particles(p).die = cl.time + 2 + (rand() And 31) * 0.02
                particles(p).color = 224 + (rand() And 7)
                particles(p).typ = pt_slowgrav

                dir_.v(0) = j * 8 + (rand() And 7)
                dir_.v(1) = i * 8 + (rand() And 7)
                dir_.v(2) = 256

                particles(p).org.v(0) = org.v(0) + dir_.v(0)
                particles(p).org.v(1) = org.v(1) + dir_.v(1)
                particles(p).org.v(2) = org.v(2) + (rand() And 63)

                VectorNormalize dir_
                vel = 50 + (rand() And 63)
                VectorScale dir_, vel, particles(p).vel
            Next
        Next
    Next
End Sub

Public Sub R_TeleportSplash(org As vec3_t)
    Dim i As Long, j As Long, k As Long, p As Long, vel As Single, dir_ As vec3_t

    For i = -16 To 15 Step 4
        For j = -16 To 15 Step 4
            For k = -24 To 31 Step 4
                p = R_AllocParticle()
                If p = -1 Then Exit Sub

                particles(p).die = cl.time + 0.2 + (rand() And 7) * 0.02
                particles(p).color = 7 + (rand() And 7)
                particles(p).typ = pt_slowgrav

                dir_.v(0) = j * 8
                dir_.v(1) = i * 8
                dir_.v(2) = k * 8

                particles(p).org.v(0) = org.v(0) + i + (rand() And 3)
                particles(p).org.v(1) = org.v(1) + j + (rand() And 3)
                particles(p).org.v(2) = org.v(2) + k + (rand() And 3)

                VectorNormalize dir_
                vel = 50 + (rand() And 63)
                VectorScale dir_, vel, particles(p).vel
            Next
        Next
    Next
End Sub

Public Sub R_RocketTrail(start As vec3_t, end_ As vec3_t, ByVal typ As Long)
    Dim vec As vec3_t, len_ As Single, j As Long, p As Long, dec As Long

    VectorSubtract end_, start, vec
    len_ = VectorNormalize(vec)
    If typ < 128 Then
        dec = 3
    Else
        dec = 1
        typ = typ - 128
    End If

    Do While len_ > 0
        len_ = len_ - dec

        p = R_AllocParticle()
        If p = -1 Then Exit Sub

        particles(p).vel = vec3_origin
        particles(p).die = cl.time + 2

        Select Case typ
        Case 0                           ' rocket trail
            particles(p).ramp = (rand() And 3)
            particles(p).color = ramp3(ToInt(particles(p).ramp))
            particles(p).typ = pt_fire
            For j = 0 To 2
                particles(p).org.v(j) = start.v(j) + ((rand() Mod 6) - 3)
            Next

        Case 1                           ' smoke smoke
            particles(p).ramp = (rand() And 3) + 2
            particles(p).color = ramp3(ToInt(particles(p).ramp))
            particles(p).typ = pt_fire
            For j = 0 To 2
                particles(p).org.v(j) = start.v(j) + ((rand() Mod 6) - 3)
            Next

        Case 2                           ' blood
            particles(p).typ = pt_grav
            particles(p).color = 67 + (rand() And 3)
            For j = 0 To 2
                particles(p).org.v(j) = start.v(j) + ((rand() Mod 6) - 3)
            Next

        Case 3, 5                        ' tracer
            particles(p).die = cl.time + 0.5
            particles(p).typ = pt_static
            If typ = 3 Then
                particles(p).color = 52 + (tracercount And 4) * 2
            Else
                particles(p).color = 230 + (tracercount And 4) * 2
            End If

            tracercount = tracercount + 1

            particles(p).org = start
            If tracercount And 1 Then
                particles(p).vel.v(0) = 30 * vec.v(1)
                particles(p).vel.v(1) = 30 * -vec.v(0)
            Else
                particles(p).vel.v(0) = 30 * -vec.v(1)
                particles(p).vel.v(1) = 30 * vec.v(0)
            End If

        Case 4                           ' slight blood
            particles(p).typ = pt_grav
            particles(p).color = 67 + (rand() And 3)
            For j = 0 To 2
                particles(p).org.v(j) = start.v(j) + ((rand() Mod 6) - 3)
            Next
            len_ = len_ - 3

        Case 6                           ' voor trail
            particles(p).color = 9 * 16 + 8 + (rand() And 3)
            particles(p).typ = pt_static
            particles(p).die = cl.time + 0.3
            For j = 0 To 2
                particles(p).org.v(j) = start.v(j) + ((rand() And 15) - 8)
            Next
        End Select

        VectorAdd start, vec, start
    Loop
End Sub

Public Sub R_DrawParticles()
    Dim p As Long, kill As Long, grav As Single, i As Long
    Dim time2 As Single, time3 As Single, time1 As Single, dvel As Single, frametime As Single

    D_StartParticles

    VectorScale vright, xscaleshrink, r_pright
    VectorScale vup, yscaleshrink, r_pup
    r_ppn = vpn

    frametime = cl.time - cl.oldtime
    time3 = frametime * 15
    time2 = frametime * 10               ' 15;
    time1 = frametime * 5
    grav = frametime * cvars(cv_sv_gravity).value * 0.05
    dvel = 4 * frametime

    Do
        kill = active_particles
        If kill <> -1 Then
            If particles(kill).die < cl.time Then
                active_particles = particles(kill).nxt
                particles(kill).nxt = free_particles
                free_particles = kill
                GoTo continue1
            End If
        End If
        Exit Do
continue1:
    Loop

    p = active_particles
    Do While p <> -1
        Do
            kill = particles(p).nxt
            If kill <> -1 Then
                If particles(kill).die < cl.time Then
                    particles(p).nxt = particles(kill).nxt
                    particles(kill).nxt = free_particles
                    free_particles = kill
                    GoTo continue2
                End If
            End If
            Exit Do
continue2:
        Loop

        D_DrawParticle p

        particles(p).org.v(0) = particles(p).org.v(0) + particles(p).vel.v(0) * frametime
        particles(p).org.v(1) = particles(p).org.v(1) + particles(p).vel.v(1) * frametime
        particles(p).org.v(2) = particles(p).org.v(2) + particles(p).vel.v(2) * frametime

        Select Case particles(p).typ
        Case pt_static

        Case pt_fire
            particles(p).ramp = particles(p).ramp + time1
            If particles(p).ramp >= 6 Then
                particles(p).die = -1
            Else
                particles(p).color = ramp3(ToInt(particles(p).ramp))
            End If
            particles(p).vel.v(2) = particles(p).vel.v(2) + grav

        Case pt_explode
            particles(p).ramp = particles(p).ramp + time2
            If particles(p).ramp >= 8 Then
                particles(p).die = -1
            Else
                particles(p).color = ramp1(ToInt(particles(p).ramp))
            End If
            For i = 0 To 2
                particles(p).vel.v(i) = particles(p).vel.v(i) + particles(p).vel.v(i) * dvel
            Next
            particles(p).vel.v(2) = particles(p).vel.v(2) - grav

        Case pt_explode2
            particles(p).ramp = particles(p).ramp + time3
            If particles(p).ramp >= 8 Then
                particles(p).die = -1
            Else
                particles(p).color = ramp2(ToInt(particles(p).ramp))
            End If
            For i = 0 To 2
                particles(p).vel.v(i) = particles(p).vel.v(i) - particles(p).vel.v(i) * frametime
            Next
            particles(p).vel.v(2) = particles(p).vel.v(2) - grav

        Case pt_blob
            For i = 0 To 2
                particles(p).vel.v(i) = particles(p).vel.v(i) + particles(p).vel.v(i) * dvel
            Next
            particles(p).vel.v(2) = particles(p).vel.v(2) - grav

        Case pt_blob2
            For i = 0 To 1
                particles(p).vel.v(i) = particles(p).vel.v(i) - particles(p).vel.v(i) * dvel
            Next
            particles(p).vel.v(2) = particles(p).vel.v(2) - grav

        Case pt_grav, pt_slowgrav
            particles(p).vel.v(2) = particles(p).vel.v(2) - grav
        End Select

        p = particles(p).nxt
    Loop

    D_EndParticles
End Sub
