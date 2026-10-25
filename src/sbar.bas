Attribute VB_Name = "sbar"
' sbar.bas - status bar code (sbar.c)
'
' The qpic_t pointers are pics() indices, 0 = NULL. Not ported:
' Sbar_UpdateScoreboard and its scoreboardtext/top/bottom arrays, which only
' the #if 0 part of Sbar_DrawScoreboard used.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Public sb_updates As Long                ' if >= vid.numpages, no update needed

Public Const SBAR_HEIGHT As Long = 24

Private Const STAT_MINUS As Long = 10    ' num frame for '-' stats digit
Private sb_nums(0 To 1, 0 To 10) As Long
Private sb_colon As Long, sb_slash As Long
Private sb_ibar As Long
Private sb_sbar As Long
Private sb_scorebar As Long

Private sb_weapons(0 To 6, 0 To 7) As Long   ' 0 is active, 1 is owned, 2-5 are flashes
Private sb_ammo(0 To 3) As Long
Private sb_sigil(0 To 3) As Long
Private sb_armor(0 To 2) As Long
Private sb_items(0 To 31) As Long

Private sb_faces(0 To 6, 0 To 1) As Long ' 0 is gibbed, 1 is dead, 2-6 are alive
                                         ' 0 is static, 1 is temporary animation
Private sb_face_invis As Long
Private sb_face_quad As Long
Private sb_face_invuln As Long
Private sb_face_invis_invuln As Long

Public sb_showscores As Boolean

' (sb_lines, the scan lines to draw, is in quakedef.bas)

Private rsb_invbar(0 To 1) As Long
Private rsb_weapons(0 To 4) As Long
Private rsb_items(0 To 1) As Long
Private rsb_ammo(0 To 2) As Long
Private rsb_teambord As Long             ' PGM 01/19/97 - team color border

'MED 01/04/97 added two more weapons + 3 alternates for grenade launcher
Private hsb_weapons(0 To 6, 0 To 4) As Long   ' 0 is active, 1 is owned, 2-5 are flashes
'MED 01/04/97 added array to simplify weapon parsing
Private hipweapons(0 To 3) As Long       ' {HIT_LASER_CANNON_BIT, HIT_MJOLNIR_BIT, 4, HIT_PROXIMITY_GUN_BIT}
'MED 01/04/97 added hipnotic items array
Private hsb_items(0 To 1) As Long

Private fragsort(0 To MAX_SCOREBOARD - 1) As Long
Private scoreboardlines As Long

' Tab key down
Private Sub Sbar_ShowScores()
    If sb_showscores Then Exit Sub
    sb_showscores = True
    sb_updates = 0
End Sub

' Tab key up
Private Sub Sbar_DontShowScores()
    sb_showscores = False
    sb_updates = 0
End Sub

' (QuakeVB) the console commands of this module, for Cmd_Call
Public Sub Sbar_CallCommand(ByVal fn As Long)
    Select Case fn
    Case XC_Sbar_ShowScores: Sbar_ShowScores
    Case XC_Sbar_DontShowScores: Sbar_DontShowScores
    End Select
End Sub

Public Sub Sbar_Changed()
    sb_updates = 0                       ' update next frame
End Sub

Public Sub Sbar_Init()
    Dim i As Long

    For i = 0 To 9
        sb_nums(0, i) = Draw_PicFromWad(va("num_%i", i))
        sb_nums(1, i) = Draw_PicFromWad(va("anum_%i", i))
    Next

    sb_nums(0, 10) = Draw_PicFromWad("num_minus")
    sb_nums(1, 10) = Draw_PicFromWad("anum_minus")

    sb_colon = Draw_PicFromWad("num_colon")
    sb_slash = Draw_PicFromWad("num_slash")

    sb_weapons(0, 0) = Draw_PicFromWad("inv_shotgun")
    sb_weapons(0, 1) = Draw_PicFromWad("inv_sshotgun")
    sb_weapons(0, 2) = Draw_PicFromWad("inv_nailgun")
    sb_weapons(0, 3) = Draw_PicFromWad("inv_snailgun")
    sb_weapons(0, 4) = Draw_PicFromWad("inv_rlaunch")
    sb_weapons(0, 5) = Draw_PicFromWad("inv_srlaunch")
    sb_weapons(0, 6) = Draw_PicFromWad("inv_lightng")

    sb_weapons(1, 0) = Draw_PicFromWad("inv2_shotgun")
    sb_weapons(1, 1) = Draw_PicFromWad("inv2_sshotgun")
    sb_weapons(1, 2) = Draw_PicFromWad("inv2_nailgun")
    sb_weapons(1, 3) = Draw_PicFromWad("inv2_snailgun")
    sb_weapons(1, 4) = Draw_PicFromWad("inv2_rlaunch")
    sb_weapons(1, 5) = Draw_PicFromWad("inv2_srlaunch")
    sb_weapons(1, 6) = Draw_PicFromWad("inv2_lightng")

    For i = 0 To 4
        sb_weapons(2 + i, 0) = Draw_PicFromWad(va("inva%i_shotgun", i + 1))
        sb_weapons(2 + i, 1) = Draw_PicFromWad(va("inva%i_sshotgun", i + 1))
        sb_weapons(2 + i, 2) = Draw_PicFromWad(va("inva%i_nailgun", i + 1))
        sb_weapons(2 + i, 3) = Draw_PicFromWad(va("inva%i_snailgun", i + 1))
        sb_weapons(2 + i, 4) = Draw_PicFromWad(va("inva%i_rlaunch", i + 1))
        sb_weapons(2 + i, 5) = Draw_PicFromWad(va("inva%i_srlaunch", i + 1))
        sb_weapons(2 + i, 6) = Draw_PicFromWad(va("inva%i_lightng", i + 1))
    Next

    sb_ammo(0) = Draw_PicFromWad("sb_shells")
    sb_ammo(1) = Draw_PicFromWad("sb_nails")
    sb_ammo(2) = Draw_PicFromWad("sb_rocket")
    sb_ammo(3) = Draw_PicFromWad("sb_cells")

    sb_armor(0) = Draw_PicFromWad("sb_armor1")
    sb_armor(1) = Draw_PicFromWad("sb_armor2")
    sb_armor(2) = Draw_PicFromWad("sb_armor3")

    sb_items(0) = Draw_PicFromWad("sb_key1")
    sb_items(1) = Draw_PicFromWad("sb_key2")
    sb_items(2) = Draw_PicFromWad("sb_invis")
    sb_items(3) = Draw_PicFromWad("sb_invuln")
    sb_items(4) = Draw_PicFromWad("sb_suit")
    sb_items(5) = Draw_PicFromWad("sb_quad")

    sb_sigil(0) = Draw_PicFromWad("sb_sigil1")
    sb_sigil(1) = Draw_PicFromWad("sb_sigil2")
    sb_sigil(2) = Draw_PicFromWad("sb_sigil3")
    sb_sigil(3) = Draw_PicFromWad("sb_sigil4")

    sb_faces(4, 0) = Draw_PicFromWad("face1")
    sb_faces(4, 1) = Draw_PicFromWad("face_p1")
    sb_faces(3, 0) = Draw_PicFromWad("face2")
    sb_faces(3, 1) = Draw_PicFromWad("face_p2")
    sb_faces(2, 0) = Draw_PicFromWad("face3")
    sb_faces(2, 1) = Draw_PicFromWad("face_p3")
    sb_faces(1, 0) = Draw_PicFromWad("face4")
    sb_faces(1, 1) = Draw_PicFromWad("face_p4")
    sb_faces(0, 0) = Draw_PicFromWad("face5")
    sb_faces(0, 1) = Draw_PicFromWad("face_p5")

    sb_face_invis = Draw_PicFromWad("face_invis")
    sb_face_invuln = Draw_PicFromWad("face_invul2")
    sb_face_invis_invuln = Draw_PicFromWad("face_inv2")
    sb_face_quad = Draw_PicFromWad("face_quad")

    Cmd_AddCommand "+showscores", XC_Sbar_ShowScores
    Cmd_AddCommand "-showscores", XC_Sbar_DontShowScores

    sb_sbar = Draw_PicFromWad("sbar")
    sb_ibar = Draw_PicFromWad("ibar")
    sb_scorebar = Draw_PicFromWad("scorebar")

    hipweapons(0) = HIT_LASER_CANNON_BIT
    hipweapons(1) = HIT_MJOLNIR_BIT
    hipweapons(2) = 4
    hipweapons(3) = HIT_PROXIMITY_GUN_BIT

    'MED 01/04/97 added new hipnotic weapons
    If hipnotic Then
        hsb_weapons(0, 0) = Draw_PicFromWad("inv_laser")
        hsb_weapons(0, 1) = Draw_PicFromWad("inv_mjolnir")
        hsb_weapons(0, 2) = Draw_PicFromWad("inv_gren_prox")
        hsb_weapons(0, 3) = Draw_PicFromWad("inv_prox_gren")
        hsb_weapons(0, 4) = Draw_PicFromWad("inv_prox")

        hsb_weapons(1, 0) = Draw_PicFromWad("inv2_laser")
        hsb_weapons(1, 1) = Draw_PicFromWad("inv2_mjolnir")
        hsb_weapons(1, 2) = Draw_PicFromWad("inv2_gren_prox")
        hsb_weapons(1, 3) = Draw_PicFromWad("inv2_prox_gren")
        hsb_weapons(1, 4) = Draw_PicFromWad("inv2_prox")

        For i = 0 To 4
            hsb_weapons(2 + i, 0) = Draw_PicFromWad(va("inva%i_laser", i + 1))
            hsb_weapons(2 + i, 1) = Draw_PicFromWad(va("inva%i_mjolnir", i + 1))
            hsb_weapons(2 + i, 2) = Draw_PicFromWad(va("inva%i_gren_prox", i + 1))
            hsb_weapons(2 + i, 3) = Draw_PicFromWad(va("inva%i_prox_gren", i + 1))
            hsb_weapons(2 + i, 4) = Draw_PicFromWad(va("inva%i_prox", i + 1))
        Next

        hsb_items(0) = Draw_PicFromWad("sb_wsuit")
        hsb_items(1) = Draw_PicFromWad("sb_eshld")
    End If

    If rogue Then
        rsb_invbar(0) = Draw_PicFromWad("r_invbar1")
        rsb_invbar(1) = Draw_PicFromWad("r_invbar2")

        rsb_weapons(0) = Draw_PicFromWad("r_lava")
        rsb_weapons(1) = Draw_PicFromWad("r_superlava")
        rsb_weapons(2) = Draw_PicFromWad("r_gren")
        rsb_weapons(3) = Draw_PicFromWad("r_multirock")
        rsb_weapons(4) = Draw_PicFromWad("r_plasma")

        rsb_items(0) = Draw_PicFromWad("r_shield1")
        rsb_items(1) = Draw_PicFromWad("r_agrav1")

        ' PGM 01/19/97 - team color border
        rsb_teambord = Draw_PicFromWad("r_teambord")
        ' PGM 01/19/97 - team color border

        rsb_ammo(0) = Draw_PicFromWad("r_ammolava")
        rsb_ammo(1) = Draw_PicFromWad("r_ammomulti")
        rsb_ammo(2) = Draw_PicFromWad("r_ammoplasma")
    End If
End Sub

'=============================================================================

' drawing routines are relative to the status bar location

Private Sub Sbar_DrawPic(ByVal x As Long, ByVal y As Long, ByVal pic As Long)
    If cl.gametype = GAME_DEATHMATCH Then
        Draw_Pic x, y + (vid.height - SBAR_HEIGHT), pic
    Else
        Draw_Pic x + ((vid.width - 320) \ 2), y + (vid.height - SBAR_HEIGHT), pic
    End If
End Sub

Private Sub Sbar_DrawTransPic(ByVal x As Long, ByVal y As Long, ByVal pic As Long)
    If cl.gametype = GAME_DEATHMATCH Then
        Draw_TransPic x, y + (vid.height - SBAR_HEIGHT), pic
    Else
        Draw_TransPic x + ((vid.width - 320) \ 2), y + (vid.height - SBAR_HEIGHT), pic
    End If
End Sub

' Draws one solid graphics character
Private Sub Sbar_DrawCharacter(ByVal x As Long, ByVal y As Long, ByVal num As Long)
    If cl.gametype = GAME_DEATHMATCH Then
        Draw_Character x + 4, y + vid.height - SBAR_HEIGHT, num
    Else
        Draw_Character x + ((vid.width - 320) \ 2) + 4, y + vid.height - SBAR_HEIGHT, num
    End If
End Sub

Private Sub Sbar_DrawString(ByVal x As Long, ByVal y As Long, ByVal str_ As String)
    If cl.gametype = GAME_DEATHMATCH Then
        Draw_String x, y + vid.height - SBAR_HEIGHT, str_
    Else
        Draw_String x + ((vid.width - 320) \ 2), y + vid.height - SBAR_HEIGHT, str_
    End If
End Sub

' (C writes into buf and returns the length)
Private Function Sbar_itoa(ByVal num As Long) As String
    Dim s As String, pow10 As Long, dig As Long

    If num < 0 Then
        s = "-"
        num = -num
    End If

    pow10 = 10
    Do While num >= pow10
        pow10 = pow10 * 10
    Loop

    Do
        pow10 = pow10 \ 10
        dig = num \ pow10
        s = s & ChrW$(48 + dig)
        num = num - dig * pow10
    Loop While pow10 <> 1

    Sbar_itoa = s
End Function

' (QuakeVB) the character code at position i (0-based) of s, as C's s[i]
Private Function CharAt(ByVal s As String, ByVal i As Long) As Long
    If i < Len(s) Then CharAt = AscW(Mid$(s, i + 1, 1))
End Function

Private Sub Sbar_DrawNum(ByVal x As Long, ByVal y As Long, ByVal num As Long, ByVal digits As Long, ByVal color As Long)
    Dim str_ As String, ptr As Long, l As Long, frame As Long

    str_ = Sbar_itoa(num)
    l = Len(str_)
    ptr = 0
    If l > digits Then ptr = ptr + (l - digits)
    If l < digits Then x = x + (digits - l) * 24

    Do While ptr < l
        If CharAt(str_, ptr) = 45 Then   ' '-'
            frame = STAT_MINUS
        Else
            frame = CharAt(str_, ptr) - 48
        End If

        Sbar_DrawTransPic x, y, sb_nums(color, frame)
        x = x + 24
        ptr = ptr + 1
    Loop
End Sub

'=============================================================================

Private Sub Sbar_SortFrags()
    Dim i As Long, j As Long, k As Long

    ' sort by frags
    scoreboardlines = 0
    For i = 0 To cl.maxclients - 1
        If Len(cl.scores(i).name) Then
            fragsort(scoreboardlines) = i
            scoreboardlines = scoreboardlines + 1
        End If
    Next

    For i = 0 To scoreboardlines - 1
        For j = 0 To scoreboardlines - 2 - i
            If cl.scores(fragsort(j)).frags < cl.scores(fragsort(j + 1)).frags Then
                k = fragsort(j)
                fragsort(j) = fragsort(j + 1)
                fragsort(j + 1) = k
            End If
        Next
    Next
End Sub

Private Function Sbar_ColorForMap(ByVal m As Long) As Long
    Sbar_ColorForMap = m + 8             ' C: m < 128 ? m + 8 : m + 8
End Function

Private Sub Sbar_SoloScoreboard()
    Dim str_ As String, minutes As Long, seconds As Long, tens As Long, units As Long, l As Long

    str_ = va("Monsters:%3i /%3i", cl.stats(STAT_MONSTERS), cl.stats(STAT_TOTALMONSTERS))
    Sbar_DrawString 8, 4, str_

    str_ = va("Secrets :%3i /%3i", cl.stats(STAT_SECRETS), cl.stats(STAT_TOTALSECRETS))
    Sbar_DrawString 8, 12, str_

    ' time
    minutes = ToInt(cl.time / 60)
    seconds = ToInt(cl.time - 60 * minutes)
    tens = seconds \ 10
    units = seconds - 10 * tens
    str_ = va("Time :%3i:%i%i", minutes, tens, units)
    Sbar_DrawString 184, 4, str_

    ' draw level name
    l = Len(cl.levelname)
    Sbar_DrawString 232 - l * 4, 12, cl.levelname
End Sub

Private Sub Sbar_DrawScoreboard()
    Sbar_SoloScoreboard
    If cl.gametype = GAME_DEATHMATCH Then Sbar_DeathmatchOverlay
End Sub

'=============================================================================

Private Sub Sbar_DrawInventory()
    Dim i As Long, num As String, time_ As Single, flashon As Long, grenadeflashing As Long

    If rogue Then
        If cl.stats(STAT_ACTIVEWEAPON) >= RIT_LAVA_NAILGUN Then
            Sbar_DrawPic 0, -24, rsb_invbar(0)
        Else
            Sbar_DrawPic 0, -24, rsb_invbar(1)
        End If
    Else
        Sbar_DrawPic 0, -24, sb_ibar
    End If

    ' weapons
    For i = 0 To 6
        If cl.items And (IT_SHOTGUN * BIT_(i)) Then
            time_ = cl.item_gettime(i)
            flashon = ToInt((cl.time - time_) * 10)
            If flashon >= 10 Then
                If cl.stats(STAT_ACTIVEWEAPON) = (IT_SHOTGUN * BIT_(i)) Then
                    flashon = 1
                Else
                    flashon = 0
                End If
            Else
                flashon = (flashon Mod 5) + 2
            End If

            Sbar_DrawPic i * 24, -16, sb_weapons(flashon, i)

            If flashon > 1 Then sb_updates = 0   ' force update to remove flash
        End If
    Next

    ' MED 01/04/97
    ' hipnotic weapons
    If hipnotic Then
        grenadeflashing = 0
        For i = 0 To 3
            If cl.items And BIT_(hipweapons(i)) Then
                time_ = cl.item_gettime(hipweapons(i))
                flashon = ToInt((cl.time - time_) * 10)
                If flashon >= 10 Then
                    If cl.stats(STAT_ACTIVEWEAPON) = BIT_(hipweapons(i)) Then
                        flashon = 1
                    Else
                        flashon = 0
                    End If
                Else
                    flashon = (flashon Mod 5) + 2
                End If

                ' check grenade launcher
                If i = 2 Then
                    If cl.items And HIT_PROXIMITY_GUN Then
                        If flashon Then
                            grenadeflashing = 1
                            Sbar_DrawPic 96, -16, hsb_weapons(flashon, 2)
                        End If
                    End If
                ElseIf i = 3 Then
                    If cl.items And (IT_SHOTGUN * BIT_(4)) Then
                        If flashon <> 0 And grenadeflashing = 0 Then
                            Sbar_DrawPic 96, -16, hsb_weapons(flashon, 3)
                        ElseIf grenadeflashing = 0 Then
                            Sbar_DrawPic 96, -16, hsb_weapons(0, 3)
                        End If
                    Else
                        Sbar_DrawPic 96, -16, hsb_weapons(flashon, 4)
                    End If
                Else
                    Sbar_DrawPic 176 + (i * 24), -16, hsb_weapons(flashon, i)
                End If
                If flashon > 1 Then sb_updates = 0   ' force update to remove flash
            End If
        Next
    End If

    If rogue Then
        ' check for powered up weapon.
        If cl.stats(STAT_ACTIVEWEAPON) >= RIT_LAVA_NAILGUN Then
            For i = 0 To 4
                If cl.stats(STAT_ACTIVEWEAPON) = (RIT_LAVA_NAILGUN * BIT_(i)) Then
                    Sbar_DrawPic (i + 2) * 24, -16, rsb_weapons(i)
                End If
            Next
        End If
    End If

    ' ammo counts
    For i = 0 To 3
        num = va("%3i", cl.stats(STAT_SHELLS + i))
        If CharAt(num, 0) <> 32 Then Sbar_DrawCharacter (6 * i + 1) * 8 - 2, -24, 18 + CharAt(num, 0) - 48
        If CharAt(num, 1) <> 32 Then Sbar_DrawCharacter (6 * i + 2) * 8 - 2, -24, 18 + CharAt(num, 1) - 48
        If CharAt(num, 2) <> 32 Then Sbar_DrawCharacter (6 * i + 3) * 8 - 2, -24, 18 + CharAt(num, 2) - 48
    Next

    flashon = 0
    ' items
    For i = 0 To 5
        If cl.items And BIT_(17 + i) Then
            time_ = cl.item_gettime(17 + i)
            If time_ <> 0 And time_ > cl.time - 2 And flashon <> 0 Then
                ' flash frame
                sb_updates = 0
            Else
                'MED 01/04/97 changed keys
                If Not hipnotic Or (i > 1) Then
                    Sbar_DrawPic 192 + i * 16, -16, sb_items(i)
                End If
            End If
            If time_ <> 0 And time_ > cl.time - 2 Then sb_updates = 0
        End If
    Next
    'MED 01/04/97 added hipnotic items
    ' hipnotic items
    If hipnotic Then
        For i = 0 To 1
            If cl.items And BIT_(24 + i) Then
                time_ = cl.item_gettime(24 + i)
                If time_ <> 0 And time_ > cl.time - 2 And flashon <> 0 Then
                    ' flash frame
                    sb_updates = 0
                Else
                    Sbar_DrawPic 288 + i * 16, -16, hsb_items(i)
                End If
                If time_ <> 0 And time_ > cl.time - 2 Then sb_updates = 0
            End If
        Next
    End If

    If rogue Then
        ' new rogue items
        For i = 0 To 1
            If cl.items And BIT_(29 + i) Then
                time_ = cl.item_gettime(29 + i)

                If time_ <> 0 And time_ > cl.time - 2 And flashon <> 0 Then
                    ' flash frame
                    sb_updates = 0
                Else
                    Sbar_DrawPic 288 + i * 16, -16, rsb_items(i)
                End If

                If time_ <> 0 And time_ > cl.time - 2 Then sb_updates = 0
            End If
        Next
    Else
        ' sigils
        For i = 0 To 3
            If cl.items And BIT_(28 + i) Then
                time_ = cl.item_gettime(28 + i)
                If time_ <> 0 And time_ > cl.time - 2 And flashon <> 0 Then
                    ' flash frame
                    sb_updates = 0
                Else
                    Sbar_DrawPic 320 - 32 + i * 8, -16, sb_sigil(i)
                End If
                If time_ <> 0 And time_ > cl.time - 2 Then sb_updates = 0
            End If
        Next
    End If
End Sub

'=============================================================================

Private Sub Sbar_DrawFrags()
    Dim i As Long, k As Long, l As Long, top As Long, bottom As Long
    Dim x As Long, y As Long, f As Long, xofs As Long, num As String

    Sbar_SortFrags

    ' draw the text
    If scoreboardlines <= 4 Then l = scoreboardlines Else l = 4

    x = 23
    If cl.gametype = GAME_DEATHMATCH Then
        xofs = 0
    Else
        xofs = (vid.width - 320) \ 2
    End If
    y = vid.height - SBAR_HEIGHT - 23

    For i = 0 To l - 1
        k = fragsort(i)
        If Len(cl.scores(k).name) = 0 Then GoTo nexti

        ' draw background
        top = cl.scores(k).colors And &HF0
        bottom = (cl.scores(k).colors And 15) * 16
        top = Sbar_ColorForMap(top)
        bottom = Sbar_ColorForMap(bottom)

        Draw_Fill xofs + x * 8 + 10, y, 28, 4, top
        Draw_Fill xofs + x * 8 + 10, y + 4, 28, 3, bottom

        ' draw number
        f = cl.scores(k).frags
        num = va("%3i", f)

        Sbar_DrawCharacter (x + 1) * 8, -24, CharAt(num, 0)
        Sbar_DrawCharacter (x + 2) * 8, -24, CharAt(num, 1)
        Sbar_DrawCharacter (x + 3) * 8, -24, CharAt(num, 2)

        If k = cl.viewentity - 1 Then
            Sbar_DrawCharacter x * 8 + 2, -24, 16
            Sbar_DrawCharacter (x + 4) * 8 - 4, -24, 17
        End If
        x = x + 4
nexti:
    Next
End Sub

'=============================================================================

Private Sub Sbar_DrawFace()
    Dim f As Long, anim As Long, top As Long, bottom As Long, xofs As Long, num As String, s As Long

    ' PGM 01/19/97 - team color drawing
    ' PGM 03/02/97 - fixed so color swatch only appears in CTF modes
    If rogue And (cl.maxclients <> 1) And (cvars(cv_teamplay).value > 3) And (cvars(cv_teamplay).value < 7) Then
        s = cl.viewentity - 1
        ' draw background
        top = cl.scores(s).colors And &HF0
        bottom = (cl.scores(s).colors And 15) * 16
        top = Sbar_ColorForMap(top)
        bottom = Sbar_ColorForMap(bottom)

        If cl.gametype = GAME_DEATHMATCH Then
            xofs = 113
        Else
            xofs = ((vid.width - 320) \ 2) + 113
        End If

        Sbar_DrawPic 112, 0, rsb_teambord
        Draw_Fill xofs, vid.height - SBAR_HEIGHT + 3, 22, 9, top
        Draw_Fill xofs, vid.height - SBAR_HEIGHT + 12, 22, 9, bottom

        ' draw number
        f = cl.scores(s).frags
        num = va("%3i", f)

        If top = 8 Then
            If CharAt(num, 0) <> 32 Then Sbar_DrawCharacter 109, 3, 18 + CharAt(num, 0) - 48
            If CharAt(num, 1) <> 32 Then Sbar_DrawCharacter 116, 3, 18 + CharAt(num, 1) - 48
            If CharAt(num, 2) <> 32 Then Sbar_DrawCharacter 123, 3, 18 + CharAt(num, 2) - 48
        Else
            Sbar_DrawCharacter 109, 3, CharAt(num, 0)
            Sbar_DrawCharacter 116, 3, CharAt(num, 1)
            Sbar_DrawCharacter 123, 3, CharAt(num, 2)
        End If

        Exit Sub
    End If
    ' PGM 01/19/97 - team color drawing

    If (cl.items And (IT_INVISIBILITY Or IT_INVULNERABILITY)) = (IT_INVISIBILITY Or IT_INVULNERABILITY) Then
        Sbar_DrawPic 112, 0, sb_face_invis_invuln
        Exit Sub
    End If
    If cl.items And IT_QUAD Then
        Sbar_DrawPic 112, 0, sb_face_quad
        Exit Sub
    End If
    If cl.items And IT_INVISIBILITY Then
        Sbar_DrawPic 112, 0, sb_face_invis
        Exit Sub
    End If
    If cl.items And IT_INVULNERABILITY Then
        Sbar_DrawPic 112, 0, sb_face_invuln
        Exit Sub
    End If

    If cl.stats(STAT_HEALTH) >= 100 Then
        f = 4
    Else
        f = cl.stats(STAT_HEALTH) \ 20
    End If

    If cl.time <= cl.faceanimtime Then
        anim = 1
        sb_updates = 0                   ' make sure the anim gets drawn over
    Else
        anim = 0
    End If
    Sbar_DrawPic 112, 0, sb_faces(f, anim)
End Sub

Public Sub Sbar_Draw()
    If scr_con_current = vid.height Then Exit Sub   ' console is full screen

    If sb_updates >= vid.numpages Then Exit Sub

    scr_copyeverything = 1

    sb_updates = sb_updates + 1

    If sb_lines <> 0 And vid.width > 320 Then Draw_TileClear 0, vid.height - sb_lines, vid.width, sb_lines

    If sb_lines > 24 Then
        Sbar_DrawInventory
        If cl.maxclients <> 1 Then Sbar_DrawFrags
    End If

    If sb_showscores Or cl.stats(STAT_HEALTH) <= 0 Then
        Sbar_DrawPic 0, 0, sb_scorebar
        Sbar_DrawScoreboard
        sb_updates = 0
    ElseIf sb_lines Then
        Sbar_DrawPic 0, 0, sb_sbar

        ' keys (hipnotic only)
        'MED 01/04/97 moved keys here so they would not be overwritten
        If hipnotic Then
            If cl.items And IT_KEY1 Then Sbar_DrawPic 209, 3, sb_items(0)
            If cl.items And IT_KEY2 Then Sbar_DrawPic 209, 12, sb_items(1)
        End If
        ' armor
        If cl.items And IT_INVULNERABILITY Then
            Sbar_DrawNum 24, 0, 666, 3, 1
            Sbar_DrawPic 0, 0, draw_disc
        Else
            If rogue Then
                Sbar_DrawNum 24, 0, cl.stats(STAT_ARMOR), 3, -(cl.stats(STAT_ARMOR) <= 25)
                If cl.items And RIT_ARMOR3 Then
                    Sbar_DrawPic 0, 0, sb_armor(2)
                ElseIf cl.items And RIT_ARMOR2 Then
                    Sbar_DrawPic 0, 0, sb_armor(1)
                ElseIf cl.items And RIT_ARMOR1 Then
                    Sbar_DrawPic 0, 0, sb_armor(0)
                End If
            Else
                Sbar_DrawNum 24, 0, cl.stats(STAT_ARMOR), 3, -(cl.stats(STAT_ARMOR) <= 25)
                If cl.items And IT_ARMOR3 Then
                    Sbar_DrawPic 0, 0, sb_armor(2)
                ElseIf cl.items And IT_ARMOR2 Then
                    Sbar_DrawPic 0, 0, sb_armor(1)
                ElseIf cl.items And IT_ARMOR1 Then
                    Sbar_DrawPic 0, 0, sb_armor(0)
                End If
            End If
        End If

        ' face
        Sbar_DrawFace

        ' health
        Sbar_DrawNum 136, 0, cl.stats(STAT_HEALTH), 3, -(cl.stats(STAT_HEALTH) <= 25)

        ' ammo icon
        If rogue Then
            If cl.items And RIT_SHELLS Then
                Sbar_DrawPic 224, 0, sb_ammo(0)
            ElseIf cl.items And RIT_NAILS Then
                Sbar_DrawPic 224, 0, sb_ammo(1)
            ElseIf cl.items And RIT_ROCKETS Then
                Sbar_DrawPic 224, 0, sb_ammo(2)
            ElseIf cl.items And RIT_CELLS Then
                Sbar_DrawPic 224, 0, sb_ammo(3)
            ElseIf cl.items And RIT_LAVA_NAILS Then
                Sbar_DrawPic 224, 0, rsb_ammo(0)
            ElseIf cl.items And RIT_PLASMA_AMMO Then
                Sbar_DrawPic 224, 0, rsb_ammo(1)
            ElseIf cl.items And RIT_MULTI_ROCKETS Then
                Sbar_DrawPic 224, 0, rsb_ammo(2)
            End If
        Else
            If cl.items And IT_SHELLS Then
                Sbar_DrawPic 224, 0, sb_ammo(0)
            ElseIf cl.items And IT_NAILS Then
                Sbar_DrawPic 224, 0, sb_ammo(1)
            ElseIf cl.items And IT_ROCKETS Then
                Sbar_DrawPic 224, 0, sb_ammo(2)
            ElseIf cl.items And IT_CELLS Then
                Sbar_DrawPic 224, 0, sb_ammo(3)
            End If
        End If

        Sbar_DrawNum 248, 0, cl.stats(STAT_AMMO), 3, -(cl.stats(STAT_AMMO) <= 10)
    End If

    If vid.width > 320 Then
        If cl.gametype = GAME_DEATHMATCH Then Sbar_MiniDeathmatchOverlay
    End If
End Sub

'=============================================================================

Private Sub Sbar_IntermissionNumber(ByVal x As Long, ByVal y As Long, ByVal num As Long, ByVal digits As Long, ByVal color As Long)
    Dim str_ As String, ptr As Long, l As Long, frame As Long

    str_ = Sbar_itoa(num)
    l = Len(str_)
    ptr = 0
    If l > digits Then ptr = ptr + (l - digits)
    If l < digits Then x = x + (digits - l) * 24

    Do While ptr < l
        If CharAt(str_, ptr) = 45 Then   ' '-'
            frame = STAT_MINUS
        Else
            frame = CharAt(str_, ptr) - 48
        End If

        Draw_TransPic x, y, sb_nums(color, frame)
        x = x + 24
        ptr = ptr + 1
    Loop
End Sub

Private Sub Sbar_DeathmatchOverlay()
    Dim pic As Long, i As Long, k As Long, l As Long, top As Long, bottom As Long
    Dim x As Long, y As Long, f As Long, num As String

    scr_copyeverything = 1
    scr_fullupdate = 0

    pic = Draw_CachePic("gfx/ranking.lmp")
    M_DrawPic (320 - pics(pic).width) \ 2, 8, pic

    ' scores
    Sbar_SortFrags

    ' draw the text
    l = scoreboardlines

    x = 80 + ((vid.width - 320) \ 2)
    y = 40
    For i = 0 To l - 1
        k = fragsort(i)
        If Len(cl.scores(k).name) = 0 Then GoTo nexti

        ' draw background
        top = cl.scores(k).colors And &HF0
        bottom = (cl.scores(k).colors And 15) * 16
        top = Sbar_ColorForMap(top)
        bottom = Sbar_ColorForMap(bottom)

        Draw_Fill x, y, 40, 4, top
        Draw_Fill x, y + 4, 40, 4, bottom

        ' draw number
        f = cl.scores(k).frags
        num = va("%3i", f)

        Draw_Character x + 8, y, CharAt(num, 0)
        Draw_Character x + 16, y, CharAt(num, 1)
        Draw_Character x + 24, y, CharAt(num, 2)

        If k = cl.viewentity - 1 Then Draw_Character x - 8, y, 12

        ' draw name
        Draw_String x + 64, y, cl.scores(k).name

        y = y + 10
nexti:
    Next
End Sub

Private Sub Sbar_MiniDeathmatchOverlay()
    Dim i As Long, k As Long, l As Long, top As Long, bottom As Long
    Dim x As Long, y As Long, f As Long, num As String, numlines As Long

    If vid.width < 512 Or sb_lines = 0 Then Exit Sub

    scr_copyeverything = 1
    scr_fullupdate = 0

    ' scores
    Sbar_SortFrags

    ' draw the text
    l = scoreboardlines
    y = vid.height - sb_lines
    numlines = sb_lines \ 8
    If numlines < 3 Then Exit Sub

    'find us
    For i = 0 To scoreboardlines - 1
        If fragsort(i) = cl.viewentity - 1 Then Exit For
    Next

    If i = scoreboardlines Then          ' we're not there
        i = 0
    Else                                 ' figure out start
        i = i - numlines \ 2
    End If

    If i > scoreboardlines - numlines Then i = scoreboardlines - numlines
    If i < 0 Then i = 0

    x = 324
    Do While i < scoreboardlines And y < vid.height - 8
        k = fragsort(i)
        If Len(cl.scores(k).name) = 0 Then GoTo nexti

        ' draw background
        top = cl.scores(k).colors And &HF0
        bottom = (cl.scores(k).colors And 15) * 16
        top = Sbar_ColorForMap(top)
        bottom = Sbar_ColorForMap(bottom)

        Draw_Fill x, y + 1, 40, 3, top
        Draw_Fill x, y + 4, 40, 4, bottom

        ' draw number
        f = cl.scores(k).frags
        num = va("%3i", f)

        Draw_Character x + 8, y, CharAt(num, 0)
        Draw_Character x + 16, y, CharAt(num, 1)
        Draw_Character x + 24, y, CharAt(num, 2)

        If k = cl.viewentity - 1 Then
            Draw_Character x, y, 16
            Draw_Character x + 32, y, 17
        End If

        ' draw name
        Draw_String x + 48, y, cl.scores(k).name

        y = y + 8
nexti:
        i = i + 1
    Loop
End Sub

Public Sub Sbar_IntermissionOverlay()
    Dim pic As Long, dig As Long, num As Long

    scr_copyeverything = 1
    scr_fullupdate = 0

    If cl.gametype = GAME_DEATHMATCH Then
        Sbar_DeathmatchOverlay
        Exit Sub
    End If

    pic = Draw_CachePic("gfx/complete.lmp")
    Draw_Pic 64, 24, pic

    pic = Draw_CachePic("gfx/inter.lmp")
    Draw_TransPic 0, 56, pic

    ' time
    dig = cl.completed_time \ 60
    Sbar_IntermissionNumber 160, 64, dig, 3, 0
    num = cl.completed_time - dig * 60
    Draw_TransPic 234, 64, sb_colon
    Draw_TransPic 246, 64, sb_nums(0, num \ 10)
    Draw_TransPic 266, 64, sb_nums(0, num Mod 10)

    Sbar_IntermissionNumber 160, 104, cl.stats(STAT_SECRETS), 3, 0
    Draw_TransPic 232, 104, sb_slash
    Sbar_IntermissionNumber 240, 104, cl.stats(STAT_TOTALSECRETS), 3, 0

    Sbar_IntermissionNumber 160, 144, cl.stats(STAT_MONSTERS), 3, 0
    Draw_TransPic 232, 144, sb_slash
    Sbar_IntermissionNumber 240, 144, cl.stats(STAT_TOTALMONSTERS), 3, 0
End Sub

Public Sub Sbar_FinaleOverlay()
    Dim pic As Long

    scr_copyeverything = 1

    pic = Draw_CachePic("gfx/finale.lmp")
    Draw_TransPic (vid.width - pics(pic).width) \ 2, 16, pic
End Sub
