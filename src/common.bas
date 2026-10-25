Attribute VB_Name = "common"
' common.bas - misc functions used in client and server (common.c): command
' line, string and number parsing, message buffers, the file system.
'
' Text is held in VB Strings with one character per byte (character code =
' byte value, made with ChrW$ and read with AscW), so Quake's high-bit
' characters survive without any code page conversion. Buffers that C keeps
' as char arrays (files, messages, the console) stay Byte arrays.
'
' The printf family (va, Con_Printf, Sys_Error, Host_Error, ...) takes the C
' format string with its backslash escapes ("\n", octal "\35"; a quote is
' written "" as usual in VB) and C's conversions; vsprintf below interprets
' both.
'
' QuakeVB: a Visual Basic 6 port of the Quake source release.
' Copyright (C) 1996-1997 Id Software, Inc.  Licensed under the GNU GPL v2.
Option Explicit

Private Const NUM_SAFE_ARGVS As Long = 7

Public cv_cmdline As Long

Public com_modified As Boolean       ' set true if using non-id files
Public proghack As Boolean
Public static_registered As Long     ' only for startup check, then set
Public msg_suppress_1 As Boolean

' if a packfile directory differs from this, it is assumed to be hacked
Private Const PAK0_COUNT As Long = 339
Private Const PAK0_CRC As Long = 32981

Public com_token As String
Public com_argc As Long
Public com_argv(0 To MAX_NUM_ARGVS + NUM_SAFE_ARGVS) As String

Private Const CMDLINE_LENGTH As Long = 256
Public com_cmdline As String

Public standard_quake As Boolean
Public rogue As Boolean
Public hipnotic As Boolean

' message reading
Public msg_readcount As Long
Public msg_badread As Boolean        ' set if a read goes beyond end of message

' file system
Public com_filesize As Long
Public com_cachedir As String
Public com_gamedir As String

Private Type packfile_t
    name As String
    filepos As Long
    filelen As Long
End Type

Private Type pack_t
    filename As String
    handle As Long
    numfiles As Long
    files() As packfile_t
End Type

' C: a linked list with the newest entry at the head; here the newest entry
' is the last one, so searches run from num_searchpaths down to 1
Private Type searchpath_t
    filename As String
    pack As Long                     ' index into packs(); 0 = a directory
End Type

Private Const MAX_FILES_IN_PACK As Long = 2048

Private packs() As pack_t
Private num_packs As Long
Private searchpaths() As searchpath_t
Private num_searchpaths As Long

' this graphic needs to be in the pak file to use registered features
Private pop(0 To 127) As Long

'============================================================================
' link_t: links are indices into link_prev() / link_next(), -1 = NULL.
' Links 0 .. MAX_EDICTS - 1 are the edicts' area links (C: ent->area, so
' EDICT_FROM_AREA(l) is l); world.bas's areanodes own the ones after those.

Public link_prev() As Long
Public link_next() As Long

Public Sub InitLinks(ByVal count As Long)
    Dim i As Long
    ReDim link_prev(0 To count - 1)
    ReDim link_next(0 To count - 1)
    For i = 0 To count - 1
        link_prev(i) = -1
        link_next(i) = -1
    Next
End Sub

' ClearLink is used for new headnodes
Public Sub ClearLink(ByVal l As Long)
    link_prev(l) = l
    link_next(l) = l
End Sub

Public Sub RemoveLink(ByVal l As Long)
    link_prev(link_next(l)) = link_prev(l)
    link_next(link_prev(l)) = link_next(l)
End Sub

Public Sub InsertLinkBefore(ByVal l As Long, ByVal before As Long)
    link_next(l) = before
    link_prev(l) = link_prev(before)
    link_next(link_prev(l)) = l
    link_prev(link_next(l)) = l
End Sub

Public Sub InsertLinkAfter(ByVal l As Long, ByVal after As Long)
    link_next(l) = link_next(after)
    link_prev(l) = after
    link_next(link_prev(l)) = l
    link_prev(link_next(l)) = l
End Sub

'============================================================================
' Strings and bytes

' Bytes b(ofs .. ofs+n-1) as a String, one character per byte.
Public Function BytesToStr(b() As Byte, ByVal ofs As Long, ByVal n As Long) As String
    Dim w() As Integer, i As Long, s As String
    If n <= 0 Then
        BytesToStr = ""
        Exit Function
    End If
    ReDim w(0 To n - 1)
    For i = 0 To n - 1
        w(i) = b(ofs + i)
    Next
    s = Space$(n)
    CopyMemory ByVal StrPtr(s), w(0), n * 2
    BytesToStr = s
End Function

' The zero-terminated string at b(ofs) (C: (char *)b + ofs).
Public Function BufToStr(b() As Byte, ByVal ofs As Long) As String
    Dim e As Long, last As Long
    last = UBound(b)
    e = ofs
    Do While e <= last
        If b(e) = 0 Then Exit Do
        e = e + 1
    Loop
    BufToStr = BytesToStr(b, ofs, e - ofs)
End Function

' At most maxlen characters of the zero-terminated string at b(ofs) (C: char name[maxlen]).
Public Function BufToStrN(b() As Byte, ByVal ofs As Long, ByVal maxlen As Long) As String
    Dim e As Long
    e = ofs
    Do While e < ofs + maxlen
        If b(e) = 0 Then Exit Do
        e = e + 1
    Loop
    BufToStrN = BytesToStr(b, ofs, e - ofs)
End Function

' A String as bytes with a terminating 0: b(0 To Len(s)).
Public Sub StrToBytesZ(ByVal s As String, b() As Byte)
    Dim n As Long, i As Long, w() As Integer
    n = Len(s)
    ReDim b(0 To n)
    If n = 0 Then Exit Sub
    ReDim w(0 To n - 1)
    CopyMemory w(0), ByVal StrPtr(s), n * 2
    For i = 0 To n - 1
        b(i) = w(i) And &HFF
    Next
End Sub

' Copies the characters of s into b(ofs ..) without a terminator.
Public Sub StrToBuf(ByVal s As String, b() As Byte, ByVal ofs As Long)
    Dim n As Long, i As Long, w() As Integer
    n = Len(s)
    If n = 0 Then Exit Sub
    ReDim w(0 To n - 1)
    CopyMemory w(0), ByVal StrPtr(s), n * 2
    For i = 0 To n - 1
        b(ofs + i) = w(i) And &HFF
    Next
End Sub

' Little-endian values inside byte buffers (C: *(int *)&b[ofs], LittleLong)
Public Function GetLong(b() As Byte, ByVal ofs As Long) As Long
    Dim l As Long
    CopyMemory l, b(ofs), 4
    GetLong = l
End Function

Public Function GetShort(b() As Byte, ByVal ofs As Long) As Long
    Dim s As Integer
    CopyMemory s, b(ofs), 2
    GetShort = s
End Function

Public Function GetFloat(b() As Byte, ByVal ofs As Long) As Single
    Dim f As Single
    CopyMemory f, b(ofs), 4
    GetFloat = f
End Function

Public Sub PutLong(b() As Byte, ByVal ofs As Long, ByVal l As Long)
    CopyMemory b(ofs), l, 4
End Sub

Public Sub PutShort(b() As Byte, ByVal ofs As Long, ByVal s As Long)
    b(ofs) = s And &HFF
    b(ofs + 1) = (s And &HFF00&) \ &H100
End Sub

Public Function BigShort(ByVal l As Long) As Long
    Dim b1 As Long, b2 As Long
    b1 = l And 255
    b2 = (l And &HFF00&) \ &H100
    BigShort = b1 * 256 + b2
    If BigShort > 32767 Then BigShort = BigShort - 65536
End Function

Public Function BigLong(ByVal l As Long) As Long
    Dim b1 As Long, b2 As Long, b3 As Long, b4 As Long
    b1 = l And 255
    b2 = (l And &HFF00&) \ &H100
    b3 = (l And &HFF0000) \ &H10000
    b4 = (l And &HFF000000) \ &H1000000 And 255
    BigLong = ((b1 And &H7F) * &H1000000) Or (b2 * &H10000) Or (b3 * &H100) Or b4
    If b1 And &H80 Then BigLong = BigLong Or &H80000000
End Function

'============================================================================
' LIBRARY REPLACEMENT FUNCTIONS

' ASCII-only case folding, as Q_strncasecmp does (LCase$ would fold more)
Private Function AsciiLower(ByVal s As String) As String
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c >= 65 And c <= 90 Then Mid$(s, i, 1) = ChrW$(c + 32)
    Next
    AsciiLower = s
End Function

' 0 if equal (C: Q_strcasecmp)
Public Function Q_strcasecmp(ByVal s1 As String, ByVal s2 As String) As Long
    If AsciiLower(s1) = AsciiLower(s2) Then Q_strcasecmp = 0 Else Q_strcasecmp = -1
End Function

' 0 if the first n characters are equal, ignoring case
Public Function Q_strncasecmp(ByVal s1 As String, ByVal s2 As String, ByVal n As Long) As Long
    If AsciiLower(Left$(s1, n)) = AsciiLower(Left$(s2, n)) Then Q_strncasecmp = 0 Else Q_strncasecmp = -1
End Function

' 0 if the first count characters are equal
Public Function Q_strncmp(ByVal s1 As String, ByVal s2 As String, ByVal count As Long) As Long
    If Left$(s1, count) = Left$(s2, count) Then Q_strncmp = 0 Else Q_strncmp = -1
End Function

Private Function CharAt(ByVal s As String, ByVal i As Long) As Long
    ' C: str[i] (0-based, 0 past the end)
    If i < Len(s) Then CharAt = AscW(Mid$(s, i + 1, 1)) Else CharAt = 0
End Function

Public Function Q_atoi(ByVal s As String) As Long
    Dim val As Long, sign As Long, c As Long, p As Long

    If CharAt(s, 0) = 45 Then          ' '-'
        sign = -1
        p = 1
    Else
        sign = 1
    End If

    val = 0

    ' check for hex
    If CharAt(s, p) = 48 And (CharAt(s, p + 1) = 120 Or CharAt(s, p + 1) = 88) Then
        p = p + 2
        Do
            c = CharAt(s, p)
            p = p + 1
            If c >= 48 And c <= 57 Then
                val = val * 16 + c - 48
            ElseIf c >= 97 And c <= 102 Then
                val = val * 16 + c - 97 + 10
            ElseIf c >= 65 And c <= 70 Then
                val = val * 16 + c - 65 + 10
            Else
                Q_atoi = val * sign
                Exit Function
            End If
        Loop
    End If

    ' check for character
    If CharAt(s, p) = 39 Then
        Q_atoi = sign * CharAt(s, p + 1)
        Exit Function
    End If

    ' assume decimal
    Do
        c = CharAt(s, p)
        p = p + 1
        If c < 48 Or c > 57 Then
            Q_atoi = val * sign
            Exit Function
        End If
        val = val * 10 + c - 48
    Loop
End Function

Public Function Q_atof(ByVal s As String) As Single
    Dim val As Double, sign As Long, c As Long, p As Long
    Dim decimal_ As Long, total As Long

    If CharAt(s, 0) = 45 Then
        sign = -1
        p = 1
    Else
        sign = 1
    End If

    val = 0

    ' check for hex
    If CharAt(s, p) = 48 And (CharAt(s, p + 1) = 120 Or CharAt(s, p + 1) = 88) Then
        p = p + 2
        Do
            c = CharAt(s, p)
            p = p + 1
            If c >= 48 And c <= 57 Then
                val = (val * 16) + c - 48
            ElseIf c >= 97 And c <= 102 Then
                val = (val * 16) + c - 97 + 10
            ElseIf c >= 65 And c <= 70 Then
                val = (val * 16) + c - 65 + 10
            Else
                Q_atof = val * sign
                Exit Function
            End If
        Loop
    End If

    ' check for character
    If CharAt(s, p) = 39 Then
        Q_atof = sign * CharAt(s, p + 1)
        Exit Function
    End If

    ' assume decimal
    decimal_ = -1
    total = 0
    Do
        c = CharAt(s, p)
        p = p + 1
        If c = 46 Then
            decimal_ = total
        Else
            If c < 48 Or c > 57 Then Exit Do
            val = val * 10 + c - 48
            total = total + 1
        End If
    Loop

    If decimal_ = -1 Then
        Q_atof = val * sign
        Exit Function
    End If
    Do While total > decimal_
        val = val / 10
        total = total - 1
    Loop

    Q_atof = val * sign
End Function

' C library atof: leading white space, a sign, digits with a decimal point,
' an exponent; stops at the first other character (VB's Val would read
' "1 2" as 12 and "&H10" as hex)
Public Function C_atof(ByVal s As String) As Double
    Dim p As Long, c As Long, neg As Boolean, n As Double, k As Long, e As Long, eneg As Boolean
    Dim seen As Boolean

    Do While CharAt(s, p) = 32 Or (CharAt(s, p) >= 9 And CharAt(s, p) <= 13)
        p = p + 1
    Loop
    c = CharAt(s, p)
    If c = 45 Or c = 43 Then
        neg = (c = 45)
        p = p + 1
    End If
    Do
        c = CharAt(s, p)
        If c < 48 Or c > 57 Then Exit Do
        n = n * 10 + c - 48
        seen = True
        p = p + 1
    Loop
    If CharAt(s, p) = 46 Then
        p = p + 1
        Do
            c = CharAt(s, p)
            If c < 48 Or c > 57 Then Exit Do
            n = n * 10 + c - 48
            k = k + 1
            seen = True
            p = p + 1
        Loop
    End If
    If seen Then
        c = CharAt(s, p)
        If c = 101 Or c = 69 Then
            c = CharAt(s, p + 1)
            If c = 45 Or c = 43 Then
                eneg = (c = 45)
                p = p + 1
            End If
            c = CharAt(s, p + 1)
            If c >= 48 And c <= 57 Then
                p = p + 1
                Do
                    c = CharAt(s, p)
                    If c < 48 Or c > 57 Then Exit Do
                    If e < 400 Then e = e * 10 + c - 48
                    p = p + 1
                Loop
                If eneg Then e = -e
            End If
        End If
    End If
    e = e - k
    If e > 308 Then
        n = n * 1E+308 * 10# ^ (e - 308)
    ElseIf e > 0 Then
        n = n * 10# ^ e
    ElseIf e < -300 Then
        n = n / 1E+300 / 10# ^ (-e - 300)
    ElseIf e < 0 Then
        n = n / 10# ^ (-e)
    End If
    If neg Then n = n * -1#              ' (not -n: VB's negation of 0 gives 0, not C's -0)
    C_atof = n
End Function

' C library atoi: leading white space, a sign, decimal digits
Public Function C_atoi(ByVal s As String) As Long
    Dim p As Long, c As Long, neg As Boolean, n As Long

    Do While CharAt(s, p) = 32 Or (CharAt(s, p) >= 9 And CharAt(s, p) <= 13)
        p = p + 1
    Loop
    c = CharAt(s, p)
    If c = 45 Or c = 43 Then
        neg = (c = 45)
        p = p + 1
    End If
    Do
        c = CharAt(s, p)
        If c < 48 Or c > 57 Then Exit Do
        n = n * 10 + c - 48
        p = p + 1
    Loop
    If neg Then n = -n
    C_atoi = n
End Function

' C's (int) of a float: truncates toward zero. Out of range (and NaN) gives
' 0x80000000 as on x86, where VB would raise error 6; a NaN comparison
' leaves the FPU's invalid flag set, so that case clears it.
Public Function ToInt(ByVal x As Double) As Long
    If x < 2147483648# And x > -2147483649# Then
        ToInt = Fix(x)
    Else
        ToInt = &H80000000
        Sys_ClearFPU
    End If
End Function

'============================================================================
' printf: C format strings for the printf family

' C: va (format, ...). A VB format argument list is passed on as a Variant
' holding the ParamArray: Dim a As Variant: a = args.
Public Function va(ByVal fmt As String, ParamArray args() As Variant) As String
    Dim a As Variant
    a = args
    va = vsprintf(fmt, a)
End Function

Private Function PadField(ByVal s As String, ByVal width As Long, ByVal leftjust As Boolean, ByVal zeropad As Boolean) As String
    Dim signch As String
    If Len(s) >= width Then
        PadField = s
    ElseIf leftjust Then
        PadField = s & Space$(width - Len(s))
    ElseIf zeropad Then
        If Left$(s, 1) = "-" Or Left$(s, 1) = "+" Or Left$(s, 1) = " " Then
            signch = Left$(s, 1)
            s = Mid$(s, 2)
        End If
        PadField = signch & String$(width - Len(s) - Len(signch), "0") & s
    Else
        PadField = Space$(width - Len(s)) & s
    End If
End Function

' |x| with prec decimals, rounded half to even like the C library
Private Function FmtFixed(ByVal x As Double, ByVal prec As Long) As String
    Dim scaled As Double, r As Double, f As Double, s As String
    If prec > 15 Then prec = 15
    scaled = x * (10# ^ prec)
    If scaled >= 1E+27 Then
        FmtFixed = Trim$(Str$(x))
        Exit Function
    End If
    r = Int(scaled)
    f = scaled - r
    If f > 0.5 Then
        r = r + 1
    ElseIf f = 0.5 Then
        If r - 2 * Int(r / 2) <> 0 Then r = r + 1
    End If
    s = CStr(CDec(r))
    If prec > 0 Then
        If Len(s) < prec + 1 Then s = String$(prec + 1 - Len(s), "0") & s
        s = Left$(s, Len(s) - prec) & "." & Right$(s, prec)
    End If
    FmtFixed = s
End Function

' |x| as d.ddde+XX
Private Function FmtExp(ByVal x As Double, ByVal prec As Long, ByVal upper As Boolean) As String
    Dim e As Long, m As String, es As String
    If x <> 0 Then
        e = Int(Log(x) / Log(10#))
        If x / (10# ^ e) >= 10 Then e = e + 1
        If x / (10# ^ e) < 1 Then e = e - 1
    End If
    m = FmtFixed(x / (10# ^ e), prec)
    If Left$(m, 2) = "10" Then              ' rounding carried into a new digit
        e = e + 1
        m = FmtFixed(x / (10# ^ e), prec)
    End If
    If e < 0 Then es = "-" Else es = "+"
    es = es & Format$(Abs(e), "00")
    If upper Then FmtExp = m & "E" & es Else FmtExp = m & "e" & es
End Function

Private Function StripZeros(ByVal s As String) As String
    Dim epos As Long, mant As String, tail As String
    epos = InStr(s, "e")
    If epos = 0 Then epos = InStr(s, "E")
    If epos > 0 Then
        mant = Left$(s, epos - 1)
        tail = Mid$(s, epos)
    Else
        mant = s
    End If
    If InStr(mant, ".") > 0 Then
        Do While Right$(mant, 1) = "0"
            mant = Left$(mant, Len(mant) - 1)
        Loop
        If Right$(mant, 1) = "." Then mant = Left$(mant, Len(mant) - 1)
    End If
    StripZeros = mant & tail
End Function

' Interprets a C format string: escapes, then %[flags][width][.prec][hlL]conv
' with d i u x X o c s f e E g G p %.
Public Function vsprintf(ByVal fmt As String, args As Variant) As String
    Dim out As String, i As Long, n As Long, c As Long, k As Long
    Dim argi As Long, nargs As Long
    Dim leftjust As Boolean, zeropad As Boolean, plus As Boolean, space_ As Boolean, alt As Boolean
    Dim width As Long, prec As Long, s As String, d As Double, l As Long, neg As Boolean
    Dim a As Variant

    If IsArray(args) Then nargs = UBound(args) + 1 Else nargs = 0
    n = Len(fmt)
    i = 1
    Do While i <= n
        c = AscW(Mid$(fmt, i, 1))
        If c = 92 And i < n Then                          ' backslash escape
            i = i + 1
            c = AscW(Mid$(fmt, i, 1))
            Select Case c
            Case 110: out = out & vbLf                      ' \n
            Case 116: out = out & vbTab                     ' \t
            Case 114: out = out & vbCr                      ' \r
            Case 48 To 55                                   ' \ooo
                k = 0
                l = 0
                Do While k < 3 And i <= n
                    c = AscW(Mid$(fmt, i, 1))
                    If c < 48 Or c > 55 Then Exit Do
                    l = l * 8 + c - 48
                    i = i + 1
                    k = k + 1
                Loop
                i = i - 1
                out = out & ChrW$(l And 255)
            Case Else: out = out & ChrW$(c)                 ' \\ \' ...
            End Select
            i = i + 1
        ElseIf c = 37 And i < n Then                      ' %
            i = i + 1
            leftjust = False: zeropad = False: plus = False: space_ = False: alt = False
            Do While i <= n
                c = AscW(Mid$(fmt, i, 1))
                If c = 45 Then
                    leftjust = True
                ElseIf c = 48 Then
                    zeropad = True
                ElseIf c = 43 Then
                    plus = True
                ElseIf c = 32 Then
                    space_ = True
                ElseIf c = 35 Then
                    alt = True
                Else
                    Exit Do
                End If
                i = i + 1
            Loop
            width = 0
            If c = 42 Then                                  ' *
                If argi < nargs Then width = CLng(args(argi))
                argi = argi + 1
                i = i + 1
            Else
                Do While i <= n
                    c = AscW(Mid$(fmt, i, 1))
                    If c < 48 Or c > 57 Then Exit Do
                    width = width * 10 + c - 48
                    i = i + 1
                Loop
            End If
            prec = -1
            If i <= n Then
                If AscW(Mid$(fmt, i, 1)) = 46 Then
                    i = i + 1
                    prec = 0
                    If i <= n Then
                        If AscW(Mid$(fmt, i, 1)) = 42 Then
                            If argi < nargs Then prec = CLng(args(argi))
                            argi = argi + 1
                            i = i + 1
                        End If
                    End If
                    Do While i <= n
                        c = AscW(Mid$(fmt, i, 1))
                        If c < 48 Or c > 57 Then Exit Do
                        prec = prec * 10 + c - 48
                        i = i + 1
                    Loop
                End If
            End If
            Do While i <= n                                 ' length modifiers
                c = AscW(Mid$(fmt, i, 1))
                If c <> 104 And c <> 108 And c <> 76 Then Exit Do
                i = i + 1
            Loop
            If i > n Then Exit Do
            c = AscW(Mid$(fmt, i, 1))
            i = i + 1

            If c = 37 Then
                out = out & "%"
            Else
                If argi < nargs Then a = args(argi) Else a = Empty
                argi = argi + 1
                Select Case c
                Case 100, 105, 117                          ' d i u
                    Select Case VarType(a)
                    Case vbSingle, vbDouble, vbCurrency, vbDecimal: d = Fix(a)
                    Case vbBoolean: d = -CLng(a)
                    Case vbString: d = Val(a)
                    Case vbEmpty: d = 0
                    Case Else: d = CDbl(a)
                    End Select
                    If d > 2147483647# Or d < -2147483648# Then d = 0
                    l = d
                    If c = 117 And l < 0 Then d = CDbl(l) + 4294967296# Else d = l
                    neg = (d < 0)
                    s = CStr(CDec(Abs(d)))
                    If prec >= 0 Then
                        If Len(s) < prec Then s = String$(prec - Len(s), "0") & s
                        zeropad = False
                    End If
                    If neg Then
                        s = "-" & s
                    ElseIf plus Then
                        s = "+" & s
                    ElseIf space_ Then
                        s = " " & s
                    End If
                    out = out & PadField(s, width, leftjust, zeropad)
                Case 120, 88, 111, 112                      ' x X o p
                    Select Case VarType(a)
                    Case vbSingle, vbDouble, vbCurrency, vbDecimal: l = Fix(a)
                    Case vbBoolean: l = -CLng(a)
                    Case vbString: l = Val(a)
                    Case vbEmpty: l = 0
                    Case Else: l = CLng(a)
                    End Select
                    If c = 111 Then s = LCase$(Oct$(l)) Else s = LCase$(Hex$(l))
                    If c = 88 Then s = UCase$(s)
                    If c = 112 Then s = String$(8 - Len(s), "0") & UCase$(s)
                    If prec >= 0 Then
                        If Len(s) < prec Then s = String$(prec - Len(s), "0") & s
                    End If
                    If alt And l <> 0 Then
                        If c = 120 Then s = "0x" & s
                        If c = 88 Then s = "0X" & s
                        If c = 111 Then s = "0" & s
                    End If
                    out = out & PadField(s, width, leftjust, zeropad)
                Case 99                                     ' c
                    If VarType(a) = vbString Then s = Left$(a, 1) Else s = ChrW$(CLng(a) And 255)
                    out = out & PadField(s, width, leftjust, False)
                Case 115                                    ' s
                    s = CStr(a & "")
                    If prec >= 0 Then s = Left$(s, prec)
                    out = out & PadField(s, width, leftjust, False)
                Case 102, 70, 101, 69, 103, 71              ' f F e E g G
                    d = CDbl(a)
                    CopyMemory l, ByVal VarPtr(d) + 4, 4   ' the sign bit, so -0 prints as -0 too
                    neg = (l < 0)
                    d = Abs(d)
                    If prec < 0 Then prec = 6
                    If c = 102 Or c = 70 Then
                        s = FmtFixed(d, prec)
                        If alt And prec = 0 Then s = s & "."
                    ElseIf c = 101 Or c = 69 Then
                        s = FmtExp(d, prec, c = 69)
                    Else
                        If prec = 0 Then prec = 1
                        ' the exponent X of the %e form after rounding decides
                        s = FmtExp(d, prec - 1, c = 71)
                        l = Val(Mid$(s, InStr(LCase$(s), "e") + 1))
                        If l >= -4 And l < prec Then s = FmtFixed(d, prec - 1 - l)
                        If Not alt Then s = StripZeros(s)
                    End If
                    If neg Then
                        s = "-" & s
                    ElseIf plus Then
                        s = "+" & s
                    ElseIf space_ Then
                        s = " " & s
                    End If
                    out = out & PadField(s, width, leftjust, zeropad)
                Case Else
                    out = out & "%" & ChrW$(c)
                End Select
            End If
        Else
            out = out & ChrW$(c)
            i = i + 1
        End If
    Loop
    vsprintf = out
End Function

'============================================================================
' MESSAGE IO FUNCTIONS
' Handles byte ordering and avoids alignment errors

'
' writing functions
'

' (QuakeVB) the values are Double: C converts the floats that callers pass
' to the int parameter by truncation, ToInt does the same here
Public Sub MSG_WriteChar(sb As sizebuf_t, ByVal c As Double)
    Dim buf As Long
    buf = SZ_GetSpace(sb, 1)
    sb.data(buf) = ToInt(c) And 255
End Sub

Public Sub MSG_WriteByte(sb As sizebuf_t, ByVal c As Double)
    Dim buf As Long
    buf = SZ_GetSpace(sb, 1)
    sb.data(buf) = ToInt(c) And 255
End Sub

Public Sub MSG_WriteShort(sb As sizebuf_t, ByVal f As Double)
    Dim buf As Long, c As Long
    c = ToInt(f)
    buf = SZ_GetSpace(sb, 2)
    sb.data(buf) = c And &HFF
    sb.data(buf + 1) = (c And &HFF00&) \ &H100
End Sub

Public Sub MSG_WriteLong(sb As sizebuf_t, ByVal f As Double)
    Dim buf As Long, c As Long
    c = ToInt(f)
    buf = SZ_GetSpace(sb, 4)
    CopyMemory sb.data(buf), c, 4
End Sub

Public Sub MSG_WriteFloat(sb As sizebuf_t, ByVal f As Single)
    Dim buf As Long
    buf = SZ_GetSpace(sb, 4)
    CopyMemory sb.data(buf), f, 4
End Sub

Public Sub MSG_WriteString(sb As sizebuf_t, ByVal s As String)
    Dim buf As Long
    buf = SZ_GetSpace(sb, Len(s) + 1)
    StrToBuf s, sb.data, buf
    sb.data(buf + Len(s)) = 0
End Sub

Public Sub MSG_WriteCoord(sb As sizebuf_t, ByVal f As Single)
    MSG_WriteShort sb, ToInt(f * 8)
End Sub

Public Sub MSG_WriteAngle(sb As sizebuf_t, ByVal f As Single)
    MSG_WriteByte sb, ((ToInt(f) * 256) \ 360) And 255
End Sub

'
' reading functions
'

Public Sub MSG_BeginReading()
    msg_readcount = 0
    msg_badread = False
End Sub

' returns -1 and sets msg_badread if no more characters are available
Public Function MSG_ReadChar() As Long
    Dim c As Long
    If msg_readcount + 1 > net_message.cursize Then
        msg_badread = True
        MSG_ReadChar = -1
        Exit Function
    End If
    c = net_message.data(msg_readcount)
    If c > 127 Then c = c - 256
    msg_readcount = msg_readcount + 1
    MSG_ReadChar = c
End Function

Public Function MSG_ReadByte() As Long
    If msg_readcount + 1 > net_message.cursize Then
        msg_badread = True
        MSG_ReadByte = -1
        Exit Function
    End If
    MSG_ReadByte = net_message.data(msg_readcount)
    msg_readcount = msg_readcount + 1
End Function

Public Function MSG_ReadShort() As Long
    Dim c As Long
    If msg_readcount + 2 > net_message.cursize Then
        msg_badread = True
        MSG_ReadShort = -1
        Exit Function
    End If
    c = net_message.data(msg_readcount) + net_message.data(msg_readcount + 1) * 256&
    If c > 32767 Then c = c - 65536
    msg_readcount = msg_readcount + 2
    MSG_ReadShort = c
End Function

Public Function MSG_ReadLong() As Long
    If msg_readcount + 4 > net_message.cursize Then
        msg_badread = True
        MSG_ReadLong = -1
        Exit Function
    End If
    MSG_ReadLong = GetLong(net_message.data, msg_readcount)
    msg_readcount = msg_readcount + 4
End Function

Public Function MSG_ReadFloat() As Single
    MSG_ReadFloat = GetFloat(net_message.data, msg_readcount)
    msg_readcount = msg_readcount + 4
End Function

Public Function MSG_ReadString() As String
    Dim l As Long, c As Long, start As Long
    start = msg_readcount
    l = 0
    Do
        c = MSG_ReadChar()
        If c = -1 Or c = 0 Then Exit Do
        l = l + 1
    Loop While l < 2048 - 1
    MSG_ReadString = BytesToStr(net_message.data, start, l)
End Function

Public Function MSG_ReadCoord() As Single
    MSG_ReadCoord = MSG_ReadShort() * (1# / 8)
End Function

Public Function MSG_ReadAngle() As Single
    MSG_ReadAngle = MSG_ReadChar() * (360# / 256)
End Function

'===========================================================================

Public Sub SZ_Alloc(buf As sizebuf_t, ByVal startsize As Long)
    If startsize < 256 Then startsize = 256
    ReDim buf.data(0 To startsize - 1)
    buf.maxsize = startsize
    buf.cursize = 0
End Sub

Public Sub SZ_Free(buf As sizebuf_t)
    buf.cursize = 0
End Sub

Public Sub SZ_Clear(buf As sizebuf_t)
    buf.cursize = 0
End Sub

' returns the offset of the new space in buf.data
Public Function SZ_GetSpace(buf As sizebuf_t, ByVal length_ As Long) As Long
    If buf.cursize + length_ > buf.maxsize Then
        If Not buf.allowoverflow Then Sys_Error "SZ_GetSpace: overflow without allowoverflow set"

        If length_ > buf.maxsize Then Sys_Error "SZ_GetSpace: %i is > full buffer size", length_

        buf.overflowed = True
        Con_Printf "SZ_GetSpace: overflow"
        SZ_Clear buf
    End If

    SZ_GetSpace = buf.cursize
    buf.cursize = buf.cursize + length_
End Function

' copies length bytes from data(ofs ..)
Public Sub SZ_Write(buf As sizebuf_t, data() As Byte, ByVal ofs As Long, ByVal length_ As Long)
    Dim p As Long
    p = SZ_GetSpace(buf, length_)
    If length_ > 0 Then CopyMemory buf.data(p), data(ofs), length_
End Sub

' strcats onto the sizebuf
Public Sub SZ_Print(buf As sizebuf_t, ByVal data As String)
    Dim l As Long, p As Long
    l = Len(data) + 1

    If buf.cursize = 0 Then
        p = SZ_GetSpace(buf, l)                ' no trailing 0 (C reads data[-1] here)
    ElseIf buf.data(buf.cursize - 1) <> 0 Then
        p = SZ_GetSpace(buf, l)                ' no trailing 0
    Else
        p = SZ_GetSpace(buf, l - 1) - 1        ' write over trailing 0
    End If
    StrToBuf data, buf.data, p
    buf.data(p + l - 1) = 0
End Sub

'============================================================================

Public Function COM_SkipPath(ByVal pathname As String) As String
    Dim i As Long, last As Long
    last = 1
    For i = 1 To Len(pathname)
        If Mid$(pathname, i, 1) = "/" Then last = i + 1
    Next
    COM_SkipPath = Mid$(pathname, last)
End Function

Public Function COM_StripExtension(ByVal in_ As String) As String
    Dim p As Long
    p = InStr(in_, ".")
    If p > 0 Then COM_StripExtension = Left$(in_, p - 1) Else COM_StripExtension = in_
End Function

Public Function COM_FileExtension(ByVal in_ As String) As String
    Dim p As Long
    p = InStr(in_, ".")
    If p = 0 Then
        COM_FileExtension = ""
    Else
        COM_FileExtension = Left$(Mid$(in_, p + 1), 7)
    End If
End Function

Public Function COM_FileBase(ByVal in_ As String) As String
    Dim s As Long, s2 As Long
    ' 1-based positions; s = 0 stands for C's s == in
    s = Len(in_)
    Do While s > 1
        If Mid$(in_, s, 1) = "." Then Exit Do
        s = s - 1
    Loop
    s2 = s
    Do While s2 >= 1
        If Mid$(in_, s2, 1) = "/" Then Exit Do
        s2 = s2 - 1
    Loop
    ' C walks s2 backwards from s while *s2 != '/'; stopping at in[-1]
    ' would be undefined, here s2 = 0 then
    If s - s2 < 2 Then
        COM_FileBase = "?model?"
    Else
        COM_FileBase = Mid$(in_, s2 + 1, s - s2 - 1)
    End If
End Function

' if path doesn't have a .EXT, append extension (extension should include the .)
Public Function COM_DefaultExtension(ByVal path As String, ByVal extension As String) As String
    Dim src As Long
    src = Len(path)
    Do While src > 1
        If Mid$(path, src, 1) = "/" Then Exit Do
        If Mid$(path, src, 1) = "." Then
            COM_DefaultExtension = path       ' it has an extension
            Exit Function
        End If
        src = src - 1
    Loop
    COM_DefaultExtension = path & extension
End Function

' Parse a token out of a zero-terminated byte buffer. data(p) is where to
' start; returns the position after the token, or -1 at the end of the data
' (C: data = COM_Parse (data), NULL at the end). The token is in com_token.
Public Function COM_Parse(data() As Byte, ByVal p As Long) As Long
    Dim c As Long, start As Long, l As Long

    com_token = ""

    If p < 0 Then
        COM_Parse = -1
        Exit Function
    End If

    ' skip whitespace
skipwhite:
    Do
        c = data(p)
        If c > 32 Then Exit Do
        If c = 0 Then
            COM_Parse = -1                    ' end of file
            Exit Function
        End If
        p = p + 1
    Loop

    ' skip // comments
    If c = 47 Then
        If data(p + 1) = 47 Then
            Do While data(p) <> 0 And data(p) <> 10
                p = p + 1
            Loop
            GoTo skipwhite
        End If
    End If

    ' handle quoted strings specially
    If c = 34 Then
        p = p + 1
        start = p
        Do
            c = data(p)
            p = p + 1
            If c = 34 Or c = 0 Then
                com_token = BytesToStr(data, start, p - 1 - start)
                COM_Parse = p
                Exit Function
            End If
        Loop
    End If

    ' parse single characters
    If c = 123 Or c = 125 Or c = 41 Or c = 40 Or c = 39 Or c = 58 Then
        com_token = ChrW$(c)
        COM_Parse = p + 1
        Exit Function
    End If

    ' parse a regular word
    start = p
    l = 0
    Do
        p = p + 1
        l = l + 1
        c = data(p)
        If c = 123 Or c = 125 Or c = 41 Or c = 40 Or c = 39 Or c = 58 Then Exit Do
    Loop While c > 32

    com_token = BytesToStr(data, start, l)
    COM_Parse = p
End Function

' Returns the position (1 to argc-1) in the program's argument list
' where the given parameter apears, or 0 if not present
Public Function COM_CheckParm(ByVal parm As String) As Long
    Dim i As Long
    For i = 1 To com_argc - 1
        If parm = com_argv(i) Then
            COM_CheckParm = i
            Exit Function
        End If
    Next
    COM_CheckParm = 0
End Function

' Looks for the pop.txt file and verifies it.
' Sets the "registered" cvar.
' Immediately exits out if an alternate game was attempted to be started without
' being registered.
Private Sub COM_CheckRegistered()
    Dim h As Long, check(0 To 255) As Byte, i As Long

    COM_OpenFile "gfx/pop.lmp", h
    static_registered = 0

    If h = -1 Then
        Con_Printf "Playing shareware version.\n"
        If com_modified Then Sys_Error "You must have the registered version to use modified games"
        Exit Sub
    End If

    Sys_FileRead h, check, 0, 256
    COM_CloseFile h

    For i = 0 To 127
        If pop(i) <> ((check(i * 2) * 256& + check(i * 2 + 1)) And &HFFFF&) Then Sys_Error "Corrupted data file."
    Next

    Cvar_Set "cmdline", com_cmdline
    Cvar_Set "registered", "1"
    static_registered = 1
    Con_Printf "Playing registered version.\n"
End Sub

Private Sub InitPop()
    Dim t As Variant, i As Long
    t = Array( _
        &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H6600&, &H0, &H0, &H0, &H6600&, &H0, _
        &H0, &H66, &H0, &H0, &H0, &H0, &H67, &H0, &H0, &H6665&, &H0, &H0, &H0, &H0, &H65, &H6600&, _
        &H63, &H6561&, &H0, &H0, &H0, &H0, &H61, &H6563&, &H64, &H6561&, &H0, &H0, &H0, &H0, &H61, &H6564&, _
        &H64, &H6564&, &H0, &H6469&, &H6969&, &H6400&, &H64, &H6564&, &H63, &H6568&, &H6200&, &H64, &H6864&, &H0, &H6268&, &H6563&, _
        &H0, &H6567&, &H6963&, &H64, &H6764&, &H63, &H6967&, &H6500&, &H0, &H6266&, &H6769&, &H6A68&, &H6768&, &H6A69&, &H6766&, &H6200&, _
        &H0, &H62, &H6566&, &H6666&, &H6666&, &H6666&, &H6562&, &H0, &H0, &H0, &H62, &H6364&, &H6664&, &H6362&, &H0, &H0, _
        &H0, &H0, &H0, &H62, &H6662&, &H0, &H0, &H0, &H0, &H0, &H0, &H61, &H6661&, &H0, &H0, &H0, _
        &H0, &H0, &H0, &H0, &H6500&, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H6400&, &H0, &H0, &H0)
    For i = 0 To 127
        pop(i) = t(i)
    Next
End Sub

Public Sub COM_InitArgv(ByVal argc As Long, argv() As String)
    Dim safe As Boolean, i As Long, j As Long, safeargvs As Variant

    safeargvs = Array("-stdvid", "-nolan", "-nosound", "-nocdaudio", "-nojoy", "-nomouse", "-dibonly")

    ' reconstitute the command line for the cmdline externally visible cvar
    com_cmdline = ""
    For j = 0 To argc - 1
        If j >= MAX_NUM_ARGVS Then Exit For
        com_cmdline = com_cmdline & argv(j)
        If Len(com_cmdline) < CMDLINE_LENGTH - 1 Then
            com_cmdline = com_cmdline & " "
        Else
            Exit For
        End If
    Next
    com_cmdline = Left$(com_cmdline, CMDLINE_LENGTH - 1)

    safe = False

    com_argc = 0
    Do While com_argc < MAX_NUM_ARGVS And com_argc < argc
        com_argv(com_argc) = argv(com_argc)
        If argv(com_argc) = "-safe" Then safe = True
        com_argc = com_argc + 1
    Loop

    If safe Then
        ' force all the safe-mode switches. Note that we reserved extra space in
        ' case we need to add these, so we don't need an overflow check
        For i = 0 To NUM_SAFE_ARGVS - 1
            com_argv(com_argc) = safeargvs(i)
            com_argc = com_argc + 1
        Next
    End If

    com_argv(com_argc) = " "             ' argvdummy

    standard_quake = True
    If COM_CheckParm("-rogue") Then
        rogue = True
        standard_quake = False
    End If

    If COM_CheckParm("-hipnotic") Then
        hipnotic = True
        standard_quake = False
    End If
End Sub

Public Sub COM_Init(ByVal basedir As String)
    InitPop
    static_registered = 1

    ' (the byte swapping functions are fixed: Windows is little-endian)

    cv_registered = Cvar_RegisterVariable("registered", "0")
    cv_cmdline = Cvar_RegisterVariable("cmdline", "0", False, True)
    Cmd_AddCommand "path", XC_COM_Path_f

    COM_InitFilesystem
    COM_CheckRegistered
End Sub

' just for debugging
Public Function memsearch(start() As Byte, ByVal count As Long, ByVal search As Long) As Long
    Dim i As Long
    For i = 0 To count - 1
        If start(i) = search Then
            memsearch = i
            Exit Function
        End If
    Next
    memsearch = -1
End Function

'=============================================================================
' QUAKE FILESYSTEM

Public Sub COM_Path_f()
    Dim s As Long
    Con_Printf "Current search path:\n"
    For s = num_searchpaths To 1 Step -1
        If searchpaths(s).pack Then
            Con_Printf "%s (%i files)\n", packs(searchpaths(s).pack).filename, packs(searchpaths(s).pack).numfiles
        Else
            Con_Printf "%s\n", searchpaths(s).filename
        End If
    Next
End Sub

' The filename will be prefixed by the current game directory
Public Sub COM_WriteFile(ByVal filename As String, data() As Byte, ByVal length_ As Long)
    Dim handle As Long, name As String

    name = com_gamedir & "/" & filename

    handle = Sys_FileOpenWrite(name)
    If handle = -1 Then
        Sys_Printf "COM_WriteFile: failed on %s\n", name
        Exit Sub
    End If

    Sys_Printf "COM_WriteFile: %s\n", name
    Sys_FileWrite handle, data, 0, length_
    Sys_FileClose handle
End Sub

' Only used for CopyFile
Public Sub COM_CreatePath(ByVal path As String)
    Dim ofs As Long
    For ofs = 2 To Len(path)
        If Mid$(path, ofs, 1) = "/" Then
            ' create the directory
            Sys_mkdir Left$(path, ofs - 1)
        End If
    Next
End Sub

' Copies a file over from the net to the local cache, creating any directories
' needed.  This is for the convenience of developers using ISDN from home.
Public Sub COM_CopyFile(ByVal netpath As String, ByVal cachepath As String)
    Dim in_ As Long, out As Long, remaining As Long, count As Long
    Dim buf(0 To 4095) As Byte

    remaining = Sys_FileOpenRead(netpath, in_)
    COM_CreatePath cachepath             ' create directories up to the cache file
    out = Sys_FileOpenWrite(cachepath)

    Do While remaining > 0
        If remaining < 4096 Then count = remaining Else count = 4096
        Sys_FileRead in_, buf, 0, count
        Sys_FileWrite out, buf, 0, count
        remaining = remaining - count
    Loop

    Sys_FileClose in_
    Sys_FileClose out
End Sub

' Finds the file in the search path.
' Sets com_filesize and one of handle or file (wantfile: C passes file, not handle;
' both are Win32 file handles here, a "file" is a new handle of its own)
Private Function COM_FindFile(ByVal filename As String, handle As Long, ByVal wantfile As Boolean) As Long
    Dim search As Long, netpath As String, cachepath As String
    Dim pak As Long, i As Long, findtime As Long, cachetime As Long

    '
    ' search through the path, one element at a time
    '
    search = num_searchpaths
    If proghack Then
        ' gross hack to use quake 1 progs with quake 2 maps
        If filename = "progs.dat" Then search = search - 1
    End If

    Do While search >= 1
        ' is the element a pak file?
        If searchpaths(search).pack Then
            ' look through all the pak file elements
            pak = searchpaths(search).pack
            For i = 0 To packs(pak).numfiles - 1
                If packs(pak).files(i).name = filename Then
                    ' found it!
                    Sys_Printf "PackFile: %s : %s\n", packs(pak).filename, filename
                    If Not wantfile Then
                        handle = packs(pak).handle
                        Sys_FileSeek packs(pak).handle, packs(pak).files(i).filepos
                    Else
                        ' open a new file on the pakfile
                        Sys_FileOpenRead packs(pak).filename, handle
                        If handle <> -1 Then Sys_FileSeek handle, packs(pak).files(i).filepos
                    End If
                    com_filesize = packs(pak).files(i).filelen
                    COM_FindFile = com_filesize
                    Exit Function
                End If
            Next
        Else
            ' check a file in the directory tree
            If static_registered = 0 Then
                ' if not a registered version, don't ever go beyond base
                If InStr(filename, "/") > 0 Or InStr(filename, "\") > 0 Then GoTo nextsearch
            End If

            netpath = searchpaths(search).filename & "/" & filename

            findtime = Sys_FileTime(netpath)
            If findtime = -1 Then GoTo nextsearch

            ' see if the file needs to be updated in the cache
            If Len(com_cachedir) = 0 Then
                cachepath = netpath
            Else
                If Len(netpath) < 2 Or Mid$(netpath, 2, 1) <> ":" Then
                    cachepath = com_cachedir & netpath
                Else
                    cachepath = com_cachedir & Mid$(netpath, 3)
                End If

                cachetime = Sys_FileTime(cachepath)

                If cachetime < findtime Then COM_CopyFile netpath, cachepath
                netpath = cachepath
            End If

            Sys_Printf "FindFile: %s\n", netpath
            com_filesize = Sys_FileOpenRead(netpath, handle)
            COM_FindFile = com_filesize
            Exit Function
        End If
nextsearch:
        search = search - 1
    Loop

    Sys_Printf "FindFile: can't find %s\n", filename

    handle = -1
    com_filesize = -1
    COM_FindFile = -1
End Function

' filename never has a leading slash, but may contain directory walks
' returns a handle and a length
' it may actually be inside a pak file
Public Function COM_OpenFile(ByVal filename As String, handle As Long) As Long
    COM_OpenFile = COM_FindFile(filename, handle, False)
End Function

' If the requested file is inside a packfile, a new file handle will be opened
' into the file (C: FILE *). Close it with Sys_FileClose.
Public Function COM_FOpenFile(ByVal filename As String, file As Long) As Long
    COM_FOpenFile = COM_FindFile(filename, file, True)
End Function

' If it is a pak file handle, don't really close it
Public Sub COM_CloseFile(ByVal h As Long)
    Dim s As Long
    For s = num_searchpaths To 1 Step -1
        If searchpaths(s).pack Then
            If packs(searchpaths(s).pack).handle = h Then Exit Sub
        End If
    Next
    Sys_FileClose h
End Sub

' Filename are reletive to the quake directory.
' Allways appends a 0 byte. Fills buf(0 To len) and returns False if the
' file was not found (C: returns NULL). The hunk, temp, cache and stack
' variants are all the same here.
Public Function COM_LoadFile(ByVal path As String, buf() As Byte) As Boolean
    Dim h As Long, l As Long

    ' look for it in the filesystem or pack files
    l = COM_OpenFile(path, h)
    If h = -1 Then
        COM_LoadFile = False
        Exit Function
    End If

    ReDim buf(0 To l)
    buf(l) = 0

    Draw_BeginDisc
    Sys_FileRead h, buf, 0, l
    COM_CloseFile h
    Draw_EndDisc

    COM_LoadFile = True
End Function

Public Function COM_LoadHunkFile(ByVal path As String, buf() As Byte) As Boolean
    COM_LoadHunkFile = COM_LoadFile(path, buf)
End Function

Public Function COM_LoadTempFile(ByVal path As String, buf() As Byte) As Boolean
    COM_LoadTempFile = COM_LoadFile(path, buf)
End Function

' Takes an explicit (not game tree related) path to a pak file.
'
' Loads the header and directory, adding the files at the beginning
' of the list so they override previous pack files. Returns the pack's
' index, or 0 if the file cannot be opened (C: NULL).
Private Function COM_LoadPackFile(ByVal packfile As String) As Long
    Dim header(0 To 11) As Byte, i As Long, numpackfiles As Long
    Dim packhandle As Long, info() As Byte, crc As Long
    Dim dirofs As Long, dirlen As Long

    If Sys_FileOpenRead(packfile, packhandle) = -1 Then
        COM_LoadPackFile = 0
        Exit Function
    End If
    Sys_FileRead packhandle, header, 0, 12
    If header(0) <> 80 Or header(1) <> 65 Or header(2) <> 67 Or header(3) <> 75 Then   ' "PACK"
        Sys_Error "%s is not a packfile", packfile
    End If
    dirofs = GetLong(header, 4)
    dirlen = GetLong(header, 8)

    numpackfiles = dirlen \ 64

    If numpackfiles > MAX_FILES_IN_PACK Then Sys_Error "%s has %i files", packfile, numpackfiles

    If numpackfiles <> PAK0_COUNT Then com_modified = True    ' not the original file

    ReDim info(0 To dirlen)
    Sys_FileSeek packhandle, dirofs
    Sys_FileRead packhandle, info, 0, dirlen

    ' crc the directory to check for modifications
    CRC_Init crc
    For i = 0 To dirlen - 1
        CRC_ProcessByte crc, info(i)
    Next
    If crc <> PAK0_CRC Then com_modified = True

    num_packs = num_packs + 1
    ReDim Preserve packs(1 To num_packs)
    packs(num_packs).filename = packfile
    packs(num_packs).handle = packhandle
    packs(num_packs).numfiles = numpackfiles
    If numpackfiles > 0 Then ReDim packs(num_packs).files(0 To numpackfiles - 1)
    ' parse the directory
    For i = 0 To numpackfiles - 1
        packs(num_packs).files(i).name = BufToStrN(info, i * 64, 56)
        packs(num_packs).files(i).filepos = GetLong(info, i * 64 + 56)
        packs(num_packs).files(i).filelen = GetLong(info, i * 64 + 60)
    Next

    Con_Printf "Added packfile %s (%i files)\n", packfile, numpackfiles
    COM_LoadPackFile = num_packs
End Function

Private Sub AddSearchPath(ByVal filename As String, ByVal pack As Long)
    num_searchpaths = num_searchpaths + 1
    ReDim Preserve searchpaths(1 To num_searchpaths)
    searchpaths(num_searchpaths).filename = filename
    searchpaths(num_searchpaths).pack = pack
End Sub

' Sets com_gamedir, adds the directory to the head of the path,
' then loads and adds pak1.pak pak2.pak ...
Private Sub COM_AddGameDirectory(ByVal dir As String)
    Dim i As Long, pak As Long

    com_gamedir = dir

    '
    ' add the directory to the search path
    '
    AddSearchPath dir, 0

    '
    ' add any pak files in the format pak0.pak pak1.pak, ...
    '
    i = 0
    Do
        pak = COM_LoadPackFile(dir & "/pak" & i & ".pak")
        If pak = 0 Then Exit Do
        AddSearchPath "", pak
        i = i + 1
    Loop
End Sub

Private Sub COM_InitFilesystem()
    Dim i As Long, j As Long, basedir As String, pak As Long

    '
    ' -basedir <path>
    ' Overrides the system supplied base directory (under GAMENAME)
    '
    i = COM_CheckParm("-basedir")
    If i <> 0 And i < com_argc - 1 Then
        basedir = com_argv(i + 1)
    Else
        basedir = host_parms.basedir
    End If

    j = Len(basedir)
    If j > 0 Then
        If Right$(basedir, 1) = "\" Or Right$(basedir, 1) = "/" Then basedir = Left$(basedir, j - 1)
    End If

    '
    ' -cachedir <path>
    ' Overrides the system supplied cache directory (NULL or /qcache)
    ' -cachedir - will disable caching.
    '
    i = COM_CheckParm("-cachedir")
    If i <> 0 And i < com_argc - 1 Then
        If Left$(com_argv(i + 1), 1) = "-" Then
            com_cachedir = ""
        Else
            com_cachedir = com_argv(i + 1)
        End If
    ElseIf Len(host_parms.cachedir) > 0 Then
        com_cachedir = host_parms.cachedir
    Else
        com_cachedir = ""
    End If

    '
    ' start up with GAMENAME by default (id1)
    '
    COM_AddGameDirectory basedir & "/" & GAMENAME

    If COM_CheckParm("-rogue") Then COM_AddGameDirectory basedir & "/rogue"
    If COM_CheckParm("-hipnotic") Then COM_AddGameDirectory basedir & "/hipnotic"

    '
    ' -game <gamedir>
    ' Adds basedir/gamedir as an override game
    '
    i = COM_CheckParm("-game")
    If i <> 0 And i < com_argc - 1 Then
        com_modified = True
        COM_AddGameDirectory basedir & "/" & com_argv(i + 1)
    End If

    '
    ' -path <dir or packfile> [<dir or packfile>] ...
    ' Fully specifies the exact serach path, overriding the generated one
    '
    i = COM_CheckParm("-path")
    If i Then
        com_modified = True
        num_searchpaths = 0
        Do
            i = i + 1
            If i >= com_argc Then Exit Do
            If Left$(com_argv(i), 1) = "+" Or Left$(com_argv(i), 1) = "-" Or Len(com_argv(i)) = 0 Then Exit Do

            If COM_FileExtension(com_argv(i)) = "pak" Then
                pak = COM_LoadPackFile(com_argv(i))
                If pak = 0 Then Sys_Error "Couldn't load packfile: %s", com_argv(i)
                AddSearchPath "", pak
            Else
                AddSearchPath com_argv(i), 0
            End If
        Loop
    End If

    If COM_CheckParm("-proghack") Then proghack = True
End Sub

' (QuakeVB) makes the array at ppsa (VarPtrArray (arr)) a view of count
' elements of elemsize bytes at pData, through the descriptor sa (which must
' live as long as the view); ClearView detaches it again (before End)
Public Sub SetView(ByVal ppsa As Long, sa As SAFEARRAY1D, ByVal pData As Long, ByVal count As Long, ByVal elemsize As Long)
    Dim p As Long
    sa.cDims = 1
    sa.fFeatures = &H11                  ' FADF_AUTO Or FADF_FIXEDSIZE: VB never frees or resizes it
    sa.cbElements = elemsize
    sa.cLocks = 0
    sa.pvData = pData
    sa.cElements = count
    sa.lLbound = 0
    p = VarPtr(sa)
    CopyMemory ByVal ppsa, p, 4
End Sub

Public Sub ClearView(ByVal ppsa As Long)
    Dim p As Long
    CopyMemory ByVal ppsa, p, 4
End Sub
