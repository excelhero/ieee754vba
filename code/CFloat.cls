VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFloat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type DoubleTemplate
    Value As Double
End Type

Private Type DoubleByteArrayTemplate
    Elements(0 To 7) As Byte
End Type

Private Type DoubleStringTemplate
    Value As String * 4
End Type


Private Type SingleTemplate
    Value As Single
End Type

Private Type SingleByteArrayTemplate
    Elements(0 To 3) As Byte
End Type

Private Type SingleStringTemplate
    Value As String * 2
End Type




Private Const DOUBLE_SUBNORMAL_MIN# = 4.94065645841247E-324


Public PrivateMessage$



Public Function PI_VariantDecimal()
    PI_VariantDecimal = CDec("3.1415926535897932384626433833")
End Function



Public Function Reinterpret_ByteArrayAsFloat(b() As Byte, VBAType&)
    Select Case VBAType
        Case vbDouble
            Static dtm As DoubleTemplate
            Static dst As DoubleStringTemplate
            dst.Value = b
            LSet dtm = dst
            On Error Resume Next
            Reinterpret_ByteArrayAsFloat = dtm.Value
            On Error GoTo 0
        Case vbSingle
            Static stm As SingleTemplate
            Static sst As SingleStringTemplate
            sst.Value = b
            LSet stm = sst
            On Error Resume Next
            Reinterpret_ByteArrayAsFloat = stm.Value
            On Error GoTo 0
    End Select
End Function


Public Function Power#(number#, thepower#)
    Power = number ^ thepower
End Function


Public Function Root#(number#, theroot#)
    Root = number ^ (1 / theroot)
End Function



Public Function Ceil(ByVal n)
    If NotNumeric(n, Ceil) Then Exit Function
    If NanInfZero(n, Ceil) Then Exit Function
    If n > -1 And n < 0 Then Ceil = NegativeZero: Exit Function
     
    Ceil = n
    If Int(n) <> n Then Ceil = Int(n + 1)
End Function

Public Function Floor(ByVal n)
    If NotNumeric(n, Floor) Then Exit Function
    If NanInfZero(n, Floor) Then Exit Function
    
    Floor = Int(n)
End Function

Private Function NanInfZero(n, report) As Boolean
    Dim s$
    s = Left$(CStr(n), 4)
    NanInfZero = InStr(s, "nan") + InStr(s, "inf") Or s = "-0" Or s = "0"
    report = n
End Function

Private Function NotNumeric(n, report) As Boolean
    Dim x#
    If IsNumeric(n) = False Then
        On Error Resume Next
        x = -(0 / 0)
        report = x
        On Error GoTo 0
        NotNumeric = True
    End If
End Function

Private Function NegativeZero#()
    NegativeZero = -DOUBLE_SUBNORMAL_MIN / 2
End Function


Public Function FloatNanPayloadImplant(n, payload, oCaller As Object, Optional negate As Boolean)
    Dim j&, nbits&(), pbits&()
    
    FloatNanPayloadImplant = n
    If VarType(n) <> vbDouble And VarType(n) <> vbSingle Then Exit Function
    If oCaller.Func_isNaN((n)) = False Then Exit Function
    If IsNumeric(payload) = False Then Exit Function
    If payload > oCaller.Spec_NanPayloadMaxValue Then Exit Function
    
    nbits = CFloat.FloatToBitArray(n, oCaller)
    ReDim pbits(LBound(nbits) To UBound(nbits))
    VariantDecimalIntegerToBitArrayFill Int(payload), pbits
    CFloat.BitArrayBitArrayLogicalOR nbits, pbits, 1, oCaller.Spec_SignificandHiBit - 1
    
    If negate Then nbits(UBound(nbits)) = 1
    FloatNanPayloadImplant = oCaller.Reinterpret_ByteArrayAsFloat(CFloat.BitArrayToByteArray(nbits))
End Function

Public Sub BitArrayBitArrayLogicalOR(abits&(), bbits&(), Optional lobit&, Optional hibit&)
    Dim j&, lo&, hi&
    lo = lobit: If lo = 0 Then lo = LBound(bbits)
    hi = hibit: If hi = 0 Then hi = UBound(bbits)
    
    For j = lo To hi
        abits(j) = bbits(j)
    Next
End Sub


Public Function BitSetAll(bits&(), Optional lobit&, Optional hibit&) As Boolean
    Dim j&, lo&, hi&
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
    For j = lo To hi
        If bits(j) = 0 Then Exit Function
    Next
    BitSetAll = True
End Function

Public Function BitSetAny(bits&(), Optional lobit&, Optional hibit&) As Boolean
    Dim j&, lo&, hi&
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
    For j = lo To hi
        If bits(j) Then
            BitSetAny = True
            Exit Function
        End If
    Next
End Function

Public Function BitSetNone(bits&(), Optional lobit&, Optional hibit&) As Boolean
    Dim j&, lo&, hi&
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
    For j = lo To hi
        If bits(j) Then Exit Function
    Next
    BitSetNone = True
End Function

Public Function BitSetOneOnly(bits&(), Optional lobit&, Optional hibit&) As Boolean
    Dim c&, j&, lo&, hi&
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
    For j = lo To hi
        If bits(j) Then c = c + 1
    Next
    BitSetOneOnly = c = 1
End Function


Public Function BitArrayToHexString$(bits&(), Optional lobit&, Optional hibit&)
    Dim c&, i&, hNdx&, nibble&, h() As Byte, lo&, hi&
    
    Static bHexChars() As Byte, pow2() As Byte
    
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
    
    Const HEX_CHARS$ = "0123456789ABCDEF"
    If (Not Not bHexChars) = 0 Then bHexChars = StrConv(HEX_CHARS, vbFromUnicode)
    If (Not Not pow2) = 0 Then pow2 = ChrW$(&H201) & ChrW$(&H804)

    
    hNdx = CFloat.Ceil((hi - lo + 1) / 4) - 1
    ReDim h(0 To hNdx)

    For i = lo To hi
        If bits(i) = 1 Then nibble = nibble + pow2(c)
        c = c + 1
        If c = 4 Then
            h(hNdx) = bHexChars(nibble)
            hNdx = hNdx - 1
            nibble = 0
            c = 0
        End If
    Next
    If c Then h(hNdx) = bHexChars(nibble)
    BitArrayToHexString = StrConv(h, vbUnicode)
End Function

Public Sub HexStringToBitArrayFill(ByVal h$, bits&(), Optional lobit&, Optional hibit&)
    Dim i&, j&, d&, n&, p&, lo&, hi&
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
    
    i = lo - 1
    For p = Len(h) To 1 Step -1
        n = InStr("123456789abcdef", Mid$(h, p, 1))
        For j = 1 To 4 'nibble
            i = i + 1
            bits(i) = 0
            If n Then
                d = Int(n / 2)
                bits(i) = n - 2 * d
                n = d
            End If
            If i = hi Then GoTo Done
        Next
    Next
Done:
End Sub

Public Sub VariantDecimalIntegerToBitArrayFill(ByVal n, bits&(), Optional lobit&, Optional hibit&)
    Dim i&, d, lo&, hi&
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
    n = CDec(n)
    For i = lo To hi
        d = Int(n / 2)
        bits(i) = n - 2 * d
        n = d
    Next
End Sub

Public Sub BinaryStringToBitArrayFill(ByVal bin$, bits&(), Optional lobit&, Optional hibit&)
    Dim i&, p&, lo&, hi&
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
    p = Len(bin)
    For i = lo To hi
        If p = 0 Then Exit For
        bits(i) = -Not ((Mid$(bin, p, 1) = 0))
        p = p - 1
    Next
End Sub

Public Function ByteArrayToBitArray(b() As Byte, Optional lobyte&, Optional hibyte&) As Long()
    Dim c&, i&, j&, byt As Byte, lo&, hi&
    Dim a&()
    
    lo = lobyte: If lo = 0 Then lo = LBound(b)
    hi = hibyte: If hi = 0 Then hi = UBound(b)
    
    ReDim a(1 To 8 * (hi - lo + 1))
    
    For j = lo To hi
        byt = b(j)
        For i = 7 To 0 Step -1
            c = c + 1
            If byt And 1 Then a(c) = 1
            byt = byt \ 2
        Next
    Next

    ByteArrayToBitArray = a
End Function

Public Sub ByteArrayToBitArrayFill(b() As Byte, bits&(), Optional lobit&, Optional hibit&)
    Dim c&, i&, j&, byt As Byte, lo&, hi&
    
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
        
    c = lo - 1
    For j = LBound(b) To UBound(b)
        byt = b(j)
        For i = 7 To 0 Step -1
            c = c + 1
            bits(c) = 0
            If byt And 1 Then bits(c) = 1
            byt = byt \ 2
            If c = hi Then GoTo Done
        Next
    Next
Done:
End Sub


Public Function BitArrayToVariantDecimalInteger(bits&(), Optional lobit& = 1, Optional hibit&)
    Dim i&, c&, lo&, hi&, n#, decint

    decint = CDec(0&)
    lo = lobit: If lo = 0 Then lo = 1
    hi = hibit: If hi = 0 Then hi = UBound(bits)

    For i = lo To hi
        If bits(i) Then decint = decint + PowerOf02VariantDecimal(i - lobit)
    Next
    BitArrayToVariantDecimalInteger = decint
End Function

Public Function BitArrayToFloatExponentlInteger(bits&(), exp_bias&, Optional exp_lobit&, Optional exp_hibit&)
    BitArrayToFloatExponentlInteger = CFloat.BitArrayToVariantDecimalInteger(bits, exp_lobit, exp_hibit) - exp_bias
End Function




Public Function ByteArrayToHexString$(b() As Byte, Optional ByVal prefix$)
    Dim j&, ndx&, hexByte$, hexstr$

    ndx = UBound(b)
    hexstr = String$((ndx + 1) * 2, "0")
    For j = 1 To Len(hexstr) - 1 Step 2
        hexByte = Hex(b(ndx))
        Mid$(hexstr, j, 2) = hexByte
        If LenB(hexByte) = 2 Then Mid$(hexstr, j, 2) = "0" & hexByte
        ndx = ndx - 1
    Next

    ByteArrayToHexString = hexstr
    If LenB(prefix) Then ByteArrayToHexString = prefix & ByteArrayToHexString
End Function


Public Function HexStringToByteArray(hexstr$) As Byte()
    Dim c&, i&, hexStrLength&, h$, b() As Byte
    
    Const HEX_ZERO____POSITIVE$ = "0000000000000000"
    
    hexstr = Replace(hexstr, " ", vbNullString)
    
    hexStrLength = Len(hexstr)
    h = HEX_ZERO____POSITIVE
    Mid$(h, 16 - Len(hexstr) + 1, Len(hexstr)) = hexstr

    ReDim b(0 To Len(h) / 2 - 1)
    c = 1
    For i = UBound(b) To 0 Step -1
        b(i) = CLng("&h" & Mid$(h, c, 2))
        c = c + 2
    Next
    HexStringToByteArray = b
End Function


Public Function BitArrayToBinaryString$(bits&(), Optional lobit&, Optional hibit&)
    Dim c&, j&, lo&, hi&
    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)

    c = hi - lo + 1
    BitArrayToBinaryString = String$(c, "0")
    
    For j = lo To hi
        Mid$(BitArrayToBinaryString, c, 1) = bits(j)
        c = c - 1
    Next
End Function


Public Function BitArrayToByteArray(bits&(), Optional lobit&, Optional hibit&) As Byte()
    Dim c&, j&, k&, z#, byt As Byte, lo&, hi&
    Dim a() As Byte

    lo = lobit: If lo = 0 Then lo = LBound(bits)
    hi = hibit: If hi = 0 Then hi = UBound(bits)
        
    z = (hi - lo + 1) / 8
    If Int(z) <> z Then z = Int(z + 1)
    ReDim a(0 To z - 1)
    k = lo Mod 8 - 1
    For j = lo To hi
        If bits(j) Then byt = byt + PowerOf02Double(k)
        k = k + 1
        If j Mod 8 = 0 Then
            a(c) = byt
            c = c + 1
            byt = 0
            k = 0
        End If
    Next
    If byt Then a(c) = byt
    BitArrayToByteArray = a
End Function


Public Function BinToHex$(ByVal bin$, Optional groupBy& = 1)
    Dim c&, i&, j&, hNdx&, nibble&, s$
    Dim b() As Byte, h() As Byte
    Static bHexChars() As Byte, pow2() As Byte
    
    Const HEX_CHARS$ = "0123456789ABCDEF"
    If (Not Not bHexChars) = 0 Then bHexChars = StrConv(HEX_CHARS, vbFromUnicode)
    If (Not Not pow2) = 0 Then pow2 = ChrW$(&H201) & ChrW$(&H804)

    b = StrConv(bin, vbFromUnicode)
    ReDim h(0 To -Int(-Len(bin) / 4) - 1)
    hNdx = UBound(h)
    
    For i = UBound(b) To 0 Step -1
        If b(i) = 49& Then nibble = nibble + pow2(c)
        c = c + 1
        If c = 4 Then
            h(hNdx) = bHexChars(nibble)
            hNdx = hNdx - 1
            nibble = 0
            c = 0
        End If
    Next
    If c Then h(hNdx) = bHexChars(nibble)
    BinToHex = StrConv(h, vbUnicode)
    
    If groupBy > 1 Then
        i = Len(BinToHex) + 1
        Do
            i = i - groupBy
            If i < 1 Then
                s = " " & Mid$(BinToHex, 1, i + groupBy - 1) & s
                Exit Do
            End If
            s = " " & Mid$(BinToHex, i, groupBy) & s
        Loop While i
        BinToHex = Trim$(s)
    End If
End Function


Public Function HexToDec(hexstr$)
    Dim c&, i&, bitVal
    Static hexVals&(), b() As Byte
    
    If (Not Not hexVals) = 0 Then
        ReDim hexVals(48 To 102)
        For i = 48 To 57
            hexVals(i) = c
            c = c + 1
        Next
        For i = 65 To 70
            hexVals(i) = c
            c = c + 1
        Next
        c = 10
        For i = 97 To 102
            hexVals(i) = c
            c = c + 1
        Next
    End If
    
    bitVal = CDec(1)
    HexToDec = CDec(0)
    b = StrConv(hexstr, vbFromUnicode)
    For i = UBound(b) To 0 Step -1
        HexToDec = HexToDec + hexVals(b(i)) * bitVal
        bitVal = bitVal * 16&
    Next
End Function


Public Function RemoveFromString$(ByVal s$, arr, Optional bFirstMatchOnly = False)
    Dim j&, e$
    If IsArray(arr) Then
        For j = LBound(arr) To UBound(arr)
            e = arr(j)
            If InStr(1, s, e, vbTextCompare) Then
                s = Replace(s, e, vbNullString, Compare:=vbTextCompare)
                If bFirstMatchOnly Then Exit For
            End If
        Next
    End If
    RemoveFromString = s
End Function


Public Function RemoveDecorationsFromHexString$(ByVal h$)
    If InStr(h, " ") Then h = Replace(h, " ", vbNullString)
    h = RemoveFromString(h, Array("&h", "0x", "#x", "0h", "h'", "u+", "16r", "&#x"), True)
    RemoveDecorationsFromHexString = h
End Function


Public Function ForceCase$(s$, theCase&)
    Select Case theCase
        Case Is < 0: ForceCase = LCase$(s)
        Case Is > 0: ForceCase = UCase$(s)
        Case Is = 0: ForceCase = s
    End Select
End Function


Public Function ForceStringToDecimalDigitsOnly$(ByVal s$)
    Dim i&, p&, max&, t&
    Dim b() As Byte, res() As Byte
    Static keep() As Boolean

    Const VALS$ = "0.123456789"
    
    If (Not Not keep) = 0 Then
        ReDim keep(0 To 255)
        For i = 1 To Len(VALS)
            keep(Asc(Mid$(VALS, i, 1))) = 1
        Next
    End If
    
    max = Len(s)
    ReDim res(0 To max)
    b = StrConv(s, vbFromUnicode)
    For i = 0 To Len(s) - 1
        t = b(i)
        If keep(t) Then
            res(p) = t
            p = p + 1
        End If
    Next
    ForceStringToDecimalDigitsOnly = Left$(StrConv(res, vbUnicode), p)
End Function

Public Function ForceStringToHexDigitsOnly$(ByVal s$)
    Dim i&, p&, max&, t&
    Dim b() As Byte, res() As Byte
    Static keep() As Boolean

    Const VALS$ = "0123456789 ABCDEFabcdef"
    
    If (Not Not keep) = 0 Then
        ReDim keep(0 To 255)
        For i = 1 To Len(VALS)
            keep(Asc(Mid$(VALS, i, 1))) = 1
        Next
    End If
    
    s = RemoveDecorationsFromHexString(s)
    max = Len(s)
    ReDim res(0 To max)
    b = StrConv(s, vbFromUnicode)
    For i = 0 To Len(s) - 1
        t = b(i)
        If keep(t) Then
            res(p) = t
            p = p + 1
        End If
    Next
    ForceStringToHexDigitsOnly = Left$(StrConv(res, vbUnicode), p)
End Function


Public Function RemoveDecorationsFromBinaryString$(ByVal b$)
    b = RemoveFromString(b, Array("0b"), True)
    RemoveDecorationsFromBinaryString = b
End Function

Public Function ForceStringToBinaryDigitsOnly$(ByVal s$)
    Dim i&, p&, max&, t&
    Dim b() As Byte, res() As Byte
    Static keep() As Boolean

    Const VALS$ = "01"
    
    If (Not Not keep) = 0 Then
        ReDim keep(0 To 255)
        For i = 1 To Len(VALS)
            keep(Asc(Mid$(VALS, i, 1))) = 1
        Next
    End If
    
    s = RemoveDecorationsFromBinaryString(s)
    max = Len(s)
    ReDim res(0 To max)
    b = StrConv(s, vbFromUnicode)
    For i = 0 To Len(s) - 1
        t = b(i)
        If keep(t) Then
            res(p) = t
            p = p + 1
        End If
    Next
    ForceStringToBinaryDigitsOnly = Left$(StrConv(res, vbUnicode), p)
End Function


Public Function DecToBin$(ByVal n, Optional padDigits& = 0)
    Dim DIGIT&
    n = CDec(n)
    If n < 0 Then Exit Function
    Do
        DIGIT = Mod2(n, 2)
        If DIGIT <> 0 Then DIGIT = 1
        DecToBin = DIGIT & DecToBin
        n = Int(n * 0.5)
    Loop While n
    If padDigits Then DecToBin = Left$(String$(padDigits, "0"), padDigits - Len(DecToBin)) & DecToBin
End Function


Public Function Mod2(n, divisor)
    Mod2 = CDec(n) - divisor * Int(n / divisor)
End Function


Public Function ByteArrayToBinaryString$(b() As Byte, Optional FieldSpaces&, Optional ByteSpaces&, Optional ByVal SubnormalAlt&)
    Dim c&, i&, j&, k&, n&, p&, byt As Byte, t$
    
    n = UBound(b) - LBound(b) + 1
    k = n * 8
    c = k + (3 - 1) * FieldSpaces
    c = c + (n - 1) * ByteSpaces
    ByteArrayToBinaryString = Space$(c)
    
    If SubnormalAlt Then
        t = ByteArrayToBinaryString(b)
        SubnormalAlt = Len(t) - InStr(9, t, "1")
        If SubnormalAlt = 0 Then SubnormalAlt = 999
    End If
    
    For j = 0 To UBound(b)
        byt = b(j)
        For i = 7 To 0 Step -1
            Mid$(ByteArrayToBinaryString, c, 1) = byt And 1
            byt = byt \ 2
            c = c - 1
            p = p + 1
            If 0 = SubnormalAlt Then
                If FieldSpaces Or ByteSpaces Then c = c - -FieldSpaces * ((k = 64 And (p = 52 Or p = 63)) Or (k = 32 And (p = 23 Or p = 31))) - -ByteSpaces * (i = 0)
            Else
                If FieldSpaces Or ByteSpaces Then c = c - -FieldSpaces * ((k = 64 And (p = SubnormalAlt Or p = 63)) Or (k = 32 And (p = SubnormalAlt Or p = 31))) - -ByteSpaces * (i = 0)
            End If
        Next
    Next

End Function


Public Sub BuildPositiveInfinity(ByRef n#, oCaller As Object)
    Dim b() As Byte, lo&, hi&, himin0&, himin1&
    
    Select Case 5
        Case vbDouble: hi = 7: himin0 = &H7F: himin1 = &HF0
        Case vbSingle: hi = 3: himin0 = &H7F: himin1 = &H80
    End Select
    
    ReDim b(0 To hi)
    b(hi - 0) = himin0
    b(hi - 1) = himin1
    n = oCaller.Reinterpret_ByteArrayAsFloat(b)
End Sub

 
Public Function HexStringToVariantFloat#(h$, oCaller As Object)
    HexStringToVariantFloat = oCaller.Reinterpret_ByteArrayAsFloat(HexStringToByteArray(h))
End Function



Public Function RightJustify$(ByVal s$, ByVal fieldwidth&, ByVal prepend&)
    If LenB(s) Then RightJustify = Space$(prepend) & Left$(Space$(fieldwidth), fieldwidth - Len(s)) & s
End Function


Public Function SymbolJustify$(ByVal s$, ByVal spaces$)
    If Left$(s, 1) = "-" Then spaces = Mid$(spaces, 2)
    If Left$(s, 1) = "+" Then spaces = Mid$(spaces, 2)
    If LenB(s) Then SymbolJustify = spaces & s
End Function


Public Function FloatToBitArray(n, oCaller As Object) As Long()
    Dim b() As Byte
    b = oCaller.Reinterpret_FloatByRefAsByteArray(n)
    FloatToBitArray = ByteArrayToBitArray(b)
End Function

Public Function FloatFromFieldValues_Calculated(ByVal s&, ByVal e&, ByVal m$, oCaller As Object, Optional bits& = 64, Optional ZeroBitENotation As Boolean = True)
    Dim res$, tmp
    
    If s = -1 Or s = 1 Then
        
        If e >= oCaller.Spec_ExponentMin + 1 And e <= oCaller.Spec_ExponentMax - 1 Then
            If m = "1" Then m = "1.0"
            If m = "0" Then m = "0.0"
            If Mid$(m, 2, 1) = "." Then
                tmp = Mid$(m, 1, 1)
                If tmp = "1" Or tmp = "0" Then
                    m = ForceStringToDecimalDigitsOnly(m)
                    m = Left$(m, 2) & Replace(Split(m, ".")(1), ".", vbNullString)
                    If m = "0.0" Then tmp = 0: GoTo CheckNegative
                    
                    tmp = Left$(m, 2)
                    If tmp = "1." Or (tmp = "0." And e = oCaller.Spec_ExponentMin + 1) Then
                        res = Me.PowerOf02Exact(e)
                        If Val(m) Then res = Me.ExactProduct(res, m, ZeroBitENotation)
                        If s = -1 Then res = "-" & res
                        Select Case bits
                            Case 64:    FloatFromFieldValues_Calculated = CDbl(res)
                            Case 32:    FloatFromFieldValues_Calculated = CSng(res)
                            Case 0:     FloatFromFieldValues_Calculated = res
                            Case Else:  FloatFromFieldValues_Calculated = "The 'bits' argument can only be 64, 32, or 0."
                        End Select
                        Exit Function
                    End If
                    
                End If
            End If
        Else
            If e = oCaller.Spec_ExponentMax And Val(m) = 1 Then tmp = oCaller.Sample_Infinity: GoTo CheckNegative
        End If
    
    End If

    tmp = oCaller.Sample_QNaN
    
    
CheckNegative:
    If s = -1 Then
        FloatFromFieldValues_Calculated = oCaller.Func_floatNegate(tmp)
    Else
        FloatFromFieldValues_Calculated = tmp
    End If

End Function



Public Function FloatFromFields_Assembled(s&, e&, m, oCaller As Object, Optional bits& = 64)

End Function



Function FloatByRefToFloatingLiteral$(ByRef n, oCaller As Object, Optional lowercase As Boolean = True)
                                                   
    Dim expo, m$, sgnf$, pref$, pval$, bIsSubnormalPow2 As Boolean, bits&()
    
    bits = FloatToBitArray(n, oCaller)
    With oCaller
        pref = Left$("-", bits(.Spec_SignBit)) & "0x" & -BitSetAny(bits, 1, .Spec_ExponentHiBit)
        expo = BitArrayToFloatExponentlInteger(bits, .Spec_ExponentBias_Base10, .Spec_ExponentLoBit, .Spec_ExponentHiBit)
        expo = FloatingLiteralAdjustmentForMinimumExponent(bits, expo, .Spec_ExponentMin, .Spec_SignificandHiBit, .Spec_ExponentMinBiased_SubnormalExpansion, bIsSubnormalPow2)
        pval = Left$("+", -(expo >= 0)) & expo
        
        If bIsSubnormalPow2 Then
            sgnf = "0"
        Else
            m = Right$(BitArrayToBinaryString(bits), .Spec_SignificandHiBit)
            If .Spec_VBAType = vbSingle Then m = m & "0"
            sgnf = TrimTrailingZeros(BinToHex(m))
        End If
        
        FloatByRefToFloatingLiteral = ForceCase(Replace(Replace(pref & "." & sgnf & "p" & pval, ".0p", "p"), "0p" & .Spec_ExponentMin, "0p+0"), --lowercase)
    End With

End Function

Private Function FloatingLiteralAdjustmentForMinimumExponent(bits&(), ByVal exponent, spec_minexp&, significand_hibit&, expmin_subnormalexpansion&, ispow2subnormal As Boolean)
    Dim j&, c&
        
    If exponent = spec_minexp Then
        For j = significand_hibit To 1 Step -1
            If bits(j) Then Exit For
            c = c + 1
        Next
        exponent = exponent - c
        If exponent < expmin_subnormalexpansion Then exponent = 0
        ispow2subnormal = BitSetOneOnly(bits)
    End If
    
    FloatingLiteralAdjustmentForMinimumExponent = exponent
End Function


Function Trisect&(ByVal s$, a$, b$, c$, ByVal delim_1$, ByVal delim_2$)
    Dim p&
    
    If Bisect(s, a, b, delim_1) Then Bisect b, b, c, delim_2
    
    If LenB(a) Then p = p + 1
    If LenB(b) Then p = p + 1
    If LenB(c) Then p = p + 1
    
    Trisect = p
End Function



Function Bisect&(ByVal s$, a$, b$, ByVal delim$)
    Dim p&
    
    a = vbNullString
    b = vbNullString
    p = InStr(1, s, delim, vbTextCompare)
    If p = 0 Then Exit Function
    
    a = Mid$(s, 1, p - 1)
    b = Mid$(s, p + Len(delim) + 0)
    
    If LenB(a) Then Bisect = Bisect + 1
    If LenB(b) Then Bisect = Bisect + 1
End Function




Public Sub HexRadixSplit(ByVal hex_str$, hex_int$, hex_frc$, Optional frac_digits& = 13)
    Dim a$, b$, c$
    
    hex_int = vbNullString
    hex_frc = vbNullString
    
    If 0 = Bisect(hex_str, a, b, ".") Then c = hex_str Else c = a
    hex_int = ForceStringToHexDigitsOnly(c)
    
    hex_frc = ForceStringToHexDigitsOnly(b)
    hex_frc = Left$(hex_frc & String$(frac_digits, "0"), frac_digits)
End Sub




Public Sub FloatingLiteralToVariantByRef(ByVal fl$, ByRef n, oCaller As Object)
    Dim a$, b$, c$, s&, e&, m, em, q&, h$, bias&, emin&, bmax&
    Dim hex_int$, hex_frc$, dec_int#, dec_frc#


'----------------------------------------------------------------------------------------------------------
'
'    Sample Hex Floating Point Literal:  "-0x1.921fb54442d18p+1" --> -3.141592653589793116
'                                                           |
'                                             base16        |            base10
'    sign(opt)  hex-prefix  hex-significand-constant  2power-marker  2^exponent  floating-suffix(opt)
'    |-------|  |--------|  |----------------------|  |-----------|  |--------|  |------------------|
'        |          |            |                          |            |               |
'        |          |            |                          |            |               |
'        -         0x           1.921fb54442d18             p           +1              d,f
'                                |
'                                |
'              hex-digit-sequence.hex-digit-sequence
'                           (opt).(opt)
'
'----------------------------------------------------------------------------------------------------------



    
'        sign                                         binary64 representation of pi
'           |
'           |  exponent    significand (mantissa)
'           ||----------||--------------------------------------------------------|
'           11000000 00001001 00100001 11111011 01010100 01000100 00101101 00011000
'           -------- -------- -------- -------- -------- -------- -------- --------
'                 C0       09       21       FB       54       44       2D       18
    
    
    
    
'        sign     binary32 representation of pi
'           |
'           | exponent   significand (mantissa)
'           ||-------||-----------------------|
'           11000000 01001001 00001111 11011011
'           -------- -------- -------- --------
'                 C0       49       0F       DB
    
    
'----------------------------------------------------------------------------------------------------------
    

    
    
    With oCaller
        
        n = 0
        h = Trim$(fl)
        h = Replace(h, " ", "")
        If 0 = LenB("h") Then Exit Sub
        If 0 = Trisect(h, a, b, c, "0x", "p") Then n = .Sample_QNaN: Exit Sub
         
        
        HexRadixSplit b, hex_int, hex_frc, .Spec_SignificandHexDigits
         
         
        dec_int = CDbl(CFloat.HexToDec(hex_int))
        dec_frc = CDbl(CFloat.HexToDec(hex_frc))
         
        e = Val(c)
        If dec_int = 0 And dec_frc = 0 Then e = 0
        s = --(a = "-")
        m = b
             
             
        emin = .Spec_ExponentMin
        bmax = .Spec_ExponentMaxBiased
        bias = .Spec_ExponentBias_Base10
         
        If dec_int = 1 And e >= emin Then 'assemble
             
            If e <> 0 Or dec_int > 0 Or dec_frc > 0 Then
                If dec_int = 0 And e = (emin + 1) Then e = emin
                If e >= emin Or (dec_int + dec_frc) > 0 Then e = e + bias
            End If
            If e > bmax Then GoTo Not_A_Number
                     
            
            Select Case .Spec_VBAType
                
                Case vbDouble
                    s = s * -&H800
                    n = HexStringToVariantFloat(VBA.Hex(s + e) & hex_frc, oCaller)
                     
                Case vbSingle
                    q = e And 1
                    e = ((e And -2) \ 2) * &H1000000
                    em = e + (Val("&h" & hex_frc) And -2) \ 2
                    n = HexStringToVariantFloat(VBA.Hex(-(s <> 0) * &H80000000 + em + q * &H800000), oCaller)
                 
            End Select
             
        
        Else 'calculate
                          
            If e > bias Then GoTo Not_A_Number
            If e < .Spec_ExponentMinBiased_SubnormalExpansion Then GoTo zero
                        
            On Error Resume Next
                n = (-1) ^ s * (dec_int + dec_frc / PowerOf16Double(Len(hex_frc))) * PowerOf02Double(e)
                If Err Then GoTo Not_A_Number
            On Error GoTo 0
             
        End If
        
        Exit Sub
    


Not_A_Number:
        n = .Sample_QNaN(HexToDec(hex_frc), s <> 0)
        Exit Sub


zero:
        If s Then n = .Sample_NegativeZero
    
    End With
    
End Sub


Function ULP_Distance(a, b, oCaller As Object)
    Dim max#, nan#, x As Object, y As Object
    
    If a = b Then ULP_Distance = 0: Exit Function
    
    max = oCaller.Sample_NormalMax
    nan = oCaller.Sample_QNaN
    
    If VarType(a) <> vbDouble And VarType(a) <> vbSingle Then ULP_Distance = nan: Exit Function
    If VarType(b) <> vbDouble And VarType(b) <> vbSingle Then ULP_Distance = nan: Exit Function
    If VarType(a) <> VarType(b) Then ULP_Distance = nan: Exit Function
        
    Set x = oCaller.Compare
    Set y = oCaller.Compare.Compare
    
    x.Value = a
    y.Value = b
    
    If x.Value_Is_NaN Then ULP_Distance = max: Exit Function
    If y.Value_Is_NaN Then ULP_Distance = max: Exit Function
    
    If x.Value_Is_Infinity Then ULP_Distance = max: Exit Function
    If y.Value_Is_Infinity Then ULP_Distance = max: Exit Function
    
    If x.Raw_SignBit <> y.Raw_SignBit Then ULP_Distance = max: Exit Function
    
    ULP_Distance = Abs(CDec(x.Raw_Entire_Decimal) - y.Raw_Entire_Decimal)
    
End Function


Function PrecisionComp&(a, b, oCaller As Object)
    
    Dim diff, bitsdiff&, max_expon&, precision&, subn As Boolean
    Dim a_exponent&, b_exponent&, d_exponent&
    
    If VarType(a) <> vbDouble And VarType(a) <> vbSingle Then Exit Function
    If VarType(b) <> vbDouble And VarType(b) <> vbSingle Then Exit Function
    If VarType(a) <> VarType(b) Then Exit Function
    
        
    With oCaller
    
        precision = .Spec_PrecisionBits
        max_expon = .Spec_ExponentMaxBiased
        
        If VarType(a) = vbDouble Then
            a_exponent = .Func_floatExponent(CDbl(a))
            b_exponent = .Func_floatExponent(CDbl(b))
            If a_exponent = max_expon Or b_exponent = max_expon Then
                If .Func_floatSignificand(CDbl(a)) Then Exit Function
                If .Func_floatSignificand(CDbl(b)) Then Exit Function
            End If
        Else
            a_exponent = .Func_floatExponent(CSng(a))
            b_exponent = .Func_floatExponent(CSng(b))
            If a_exponent = max_expon Or b_exponent = max_expon Then
                If .Func_floatSignificand(CSng(a)) Then Exit Function
                If .Func_floatSignificand(CSng(b)) Then Exit Function
            End If
        End If
        
        
        On Error Resume Next
          If a = b Then On Error GoTo 0: PrecisionComp = precision: Exit Function
          diff = Abs(a - b): If Err Then Err.Clear: Exit Function
        On Error GoTo 0
        
        d_exponent = .Func_floatExponent(Abs(a - b))
        bitsdiff = (((a_exponent + b_exponent) - 1) And -2) \ 2 - d_exponent
        subn = (a_exponent = 0) And (b_exponent = 0)
            
        Select Case True
            Case subn And Abs(a - b) > 0:     PrecisionComp = 1 + bitsdiff + 1 + precision + 0 - .Func_floatExponent(Abs(a - b) * PowerOf02Double(precision - 0))
            Case d_exponent = 0:        PrecisionComp = bitsdiff + 1 + precision + 0 - .Func_floatExponent(Abs(a - b) * PowerOf02Double(precision - 0))
            Case bitsdiff > precision:  PrecisionComp = 0
            Case bitsdiff > 0:          PrecisionComp = bitsdiff + 1
            Case bitsdiff = 0:          PrecisionComp = -(a_exponent = b_exponent)
        End Select

    End With
    
End Function




Public Function TrimTrailingZeros$(s$)
    Dim p&
    For p = Len(s) To 1 Step -1
        If Mid$(s, p, 1) <> "0" Then Exit For
    Next
    If p = 0 Then p = 1
    TrimTrailingZeros = Left$(s, p)
End Function

Public Function TrimLeadingZeros$(s$)
    Dim p&
    For p = 1 To Len(s)
        If Mid$(s, p, 1) <> "0" Then Exit For
    Next
    TrimLeadingZeros = Mid$(s, p)
End Function


Public Function SignificandIntegerToBase10DecimalString$(si, bits&(), sig_bit_lo&, sig_bit_hi&, Optional subnormal As Boolean)

    VariantDecimalIntegerToBitArrayFill si, bits, sig_bit_lo, sig_bit_hi
    SignificandIntegerToBase10DecimalString = BitArrayToMantissaBase10DecimalString(bits, -(Not subnormal))

End Function


Public Function MantissaHexadecimalInteger(bits&(), mantissa_hibit&)
    MantissaHexadecimalInteger = TrimLeadingZeros(ByteArrayToHexString(BitArrayToByteArray(bits, 1, mantissa_hibit)))
End Function


Public Function BitArrayToMantissaBase10DecimalString$(bits&(), Optional leadinginteger& = -1, Optional custom_hibit& = -1)
    Dim i&, c&, lo&, hi&, intval$
    Static Cached_MantissaDigits, tot&()
           
    If custom_hibit > -1 Then
        hi = custom_hibit
    Else
        Select Case UBound(bits)
            Case 64: hi = 52
            Case 32: hi = 23
            Case Else: Exit Function
        End Select
    End If
    lo = 1
        
           
    If Not IsArray(Cached_MantissaDigits) Then Cached_MantissaDigits = CreateCache_MantissaDigits
           
    ReDim tot(1 To 9)
    For i = lo To hi
        If bits(i) Then tot = Base1E6Add(Cached_MantissaDigits(hi - c), tot)
        c = c + 1
    Next
    
    intval = "0"
    If leadinginteger > -1 Then intval = leadinginteger
    BitArrayToMantissaBase10DecimalString = intval & "." & Base1E6ToBase10(tot, hi, , True)
    'Stop
End Function


Public Function Base1E6ToBase10$(op1, Optional padZeros = 0, Optional radixpointpos& = 0, Optional bIsFraction As Boolean)
    Dim j&, t$
    
    Base1E6ToBase10 = String$(6 * UBound(op1), "0")
    
    For j = 1 To UBound(op1)
        t = CStr(op1(j))
        Mid$(Base1E6ToBase10, Len(Base1E6ToBase10) - j * 6 + 1 + 6 - Len(t)) = t
    Next
    
    
    If padZeros Then Base1E6ToBase10 = Mid$(Base1E6ToBase10, 3, padZeros)
    
    
    If radixpointpos Then
        Base1E6ToBase10 = Left$(Base1E6ToBase10, Len(Base1E6ToBase10) - radixpointpos) & "." & Mid$(Base1E6ToBase10, Len(Base1E6ToBase10) - radixpointpos + 1)
        Base1E6ToBase10 = CFloat.TrimLeadingZeros(Base1E6ToBase10)
        If Left$(Base1E6ToBase10, 1) = "." Then Base1E6ToBase10 = "0" & Base1E6ToBase10
        Base1E6ToBase10 = Me.TrimTrailingZeros(Base1E6ToBase10)
        Exit Function
    End If

    
    If padZeros Then
        If padZeros > Len(Base1E6ToBase10) Then
            t = String(padZeros, "0")
            Mid$(t, 53 - Len(Base1E6ToBase10)) = Base1E6ToBase10
            Base1E6ToBase10 = t
        Else
            Base1E6ToBase10 = Right$(Base1E6ToBase10, padZeros)
        End If
    End If
    
    Dim c&, q&
    For j = 1 To UBound(op1)
        If op1(j) = 0 Then
            c = c + 6
        Else
            t = op1(j)
            For q = Len(t) To 1 Step -1
                If Mid$(t, q, 1) = 0 Then
                    c = c + 1
                Else
                    GoTo TheExit
                End If
            Next
        End If
    Next
    
TheExit:
    Base1E6ToBase10 = Me.TrimTrailingZeros(Base1E6ToBase10)
    If Not bIsFraction Then Base1E6ToBase10 = Base1E6ToBase10 & String$(c, "0")
    
End Function


Public Function Base1E6Add(op1, op2) As Long()
    Dim j&, s&, carry&, tot&()
    
    ReDim tot(1 To UBound(op1))
    
    For j = 1 To UBound(op1)
        s = op1(j) + op2(j) + carry
        tot(j) = s
        carry = 0
        If s >= 1000000 Then
            tot(j) = s - 1000000
            carry = s \ 1000000
        End If
    Next
    
    Base1E6Add = tot
End Function


Public Function Base1E6Subtract$(ByVal a$, ByVal b$, Optional UseENotationLongerThanDigits& = 17)
    Dim j&, s&, borrow&, tot&(), x&(), y&(), t$, bNeg As Boolean
        
    If InStr(a, "inf") Then Base1E6Subtract = "inf":  Exit Function
    If InStr(b, "inf") Then Base1E6Subtract = "-inf": Exit Function
    If InStr(a, "nan") Then Base1E6Subtract = "nan":  Exit Function
    If InStr(b, "nan") Then Base1E6Subtract = "nan":  Exit Function
        
    If b > a Then t = b: b = a: a = t: bNeg = True
        
    If InStr(a, ".") = 0 And InStr(a, ".") = 0 Then
        
        If LenB(a) <> LenB(b) Then
            If LenB(a) < LenB(b) Then
                a = String$(Len(b) - Len(a), "0") & a
            Else
                b = String$(Len(a) - Len(b), "0") & b
            End If
        End If
        
    Else
        Dim aa, bb, radixpointpos&
        
        If InStr(a, ".") = 0 Then a = a & ".0"
        If InStr(b, ".") = 0 Then b = b & ".0"
        
        aa = Split(a, ".")
        bb = Split(b, ".")
        
        If LenB(aa(0)) <> LenB(bb(0)) Then
            If LenB(aa(0)) < LenB(bb(0)) Then
                aa(0) = String$(Len(bb(0)) - Len(aa(0)), "0") & aa(0)
            Else
                bb(0) = String$(Len(aa(0)) - Len(bb(0)), "0") & bb(0)
            End If
        End If
        
        If LenB(aa(1)) <> LenB(bb(1)) Then
            If LenB(aa(1)) < LenB(bb(1)) Then
                aa(1) = aa(1) & String$(Len(bb(1)) - Len(aa(1)), "0")
            Else
                bb(1) = bb(1) & String$(Len(aa(1)) - Len(bb(1)), "0")
            End If
        End If
        
        a = aa(0) & aa(1)
        b = bb(0) & bb(1)
        radixpointpos = Len(aa(1))
        
    End If
    
    
    If Right$(a, 2) = ".0" Then a = Left$(a, Len(a) - 2)
    If Right$(b, 2) = ".0" Then b = Left$(b, Len(b) - 2)
    
    If Left$(a, 2) = "0-" Then a = Split(a, "-")(1)
    If Left$(b, 2) = "0-" Then b = Split(b, "-")(1)
    
    x = CFloat.Base10ToBase1E6x(a)
    y = CFloat.Base10ToBase1E6x(b)
    
            
    ReDim tot(1 To UBound(x))
    
    For j = 1 To UBound(x)
        borrow = 0
        If y(j) > x(j) Then
            borrow = 1000000
            If UBound(x) >= j + 1 Then
                x(j + 1) = x(j + 1) - 1
                s = borrow + x(j) - y(j)
            Else
                s = x(j) - y(j)
            End If
        Else
            s = x(j) - y(j)
        End If
        
        tot(j) = s
    Next
    
    '666666666676767676jglkdjfg        I think it's the    x.0  on the first term causing the error.
    '? CFloat.Base1E6Subtract("1660093443604480.0", "1660093475923250", False)
    
    Base1E6Subtract = CFloat.Base1E6ToBase10(tot, , radixpointpos)
    Base1E6Subtract = Replace(Base1E6Subtract, "-", vbNullString)
    Base1E6Subtract = CFloat.TrimLeadingZeros(Base1E6Subtract)
    
    If Left$(Base1E6Subtract, 1) = "." Then Base1E6Subtract = "0" & Base1E6Subtract
    
    If bNeg Then Base1E6Subtract = "-" & Base1E6Subtract
    If Len(Base1E6Subtract) > UseENotationLongerThanDigits Then Base1E6Subtract = E_Notation(Base1E6Subtract)
End Function



Public Function Base10ToBase1E6(n$) As Long()
    Dim i&, z&, d&()
    
    ReDim d(1 To -Int(-(Len(n) / 6)))
    z = Len(n) + 1
    For i = 1 To 9
        z = z - 6
        If z > 0 Then
            d(i) = Mid$(n, z, 6)
        Else
            d(i) = Mid$(n, 1, 6 - (1 - z))
        End If
    Next
    Base10ToBase1E6 = d
End Function

Public Function Base10ToBase1E6x(n$) As Long()
    Dim i&, z&, d&(), t$
    
    ReDim d(1 To -Int(-(Len(n) / 6)))
    z = Len(n) + 1
    For i = 1 To UBound(d)
        z = z - 6
        If z > 1 Then
            d(i) = Mid$(n, z, 6)
        Else
            t = Mid$(n, 1, 6 - (1 - z))
            If t = "-" Then t = 0
            d(i) = t
        End If
    Next
    Base10ToBase1E6x = d
End Function






Function ExactProduct$(ByVal a$, ByVal b$, Optional ENotation As Boolean)
    Dim d&, i&, j&, k&, p&, q&, f$
    
    If a = "0" Or b = "0" Then ExactProduct = "0": Exit Function
    If LenB(a) = 0 Or LenB(b) = 0 Then ExactProduct = 0: Exit Function
    
    d = InStr(a, "."): If d Then k = Len(a) - d:      a = Replace(a, ".", vbNullString)
    d = InStr(b, "."): If d Then k = Len(b) - d + k:  b = Replace(b, ".", vbNullString)
    
    a = Trim$(a)
    b = Trim$(b)
    
    For i = Len(a) To 1 Step -1
        For j = Len(b) To 1 Step -1
            ExactProduct = ExactSum(ExactProduct, Mid$(a, i, 1) * Mid$(b, j, 1), p + q)
            q = q + 1
        Next
        q = 0
        p = p + 1
    Next
        
    If k Then
        f$ = Me.TrimTrailingZeros(Right(ExactProduct, k))
        ExactProduct = Left(ExactProduct, Len(ExactProduct) - k)
        If Len(f) > 0 Then ExactProduct = Me.TrimLeadingZeros(ExactProduct) & "." & f
        If Left$(ExactProduct, 1) = "." Then ExactProduct = "0" & ExactProduct
    End If
        
    If ENotation Then ExactProduct = E_Notation(ExactProduct)
        
End Function

Function ExactSum$(a$, b$, Optional ndx&)
'a$ and b$ must be whole numbers
'ndx& is used by ExactProduct$()

    Dim i&, j&, sum&
    
    sum = 0
    j = Len(b) + 1
    i = Len(a) + 1 - ndx
    
    ExactSum = Mid$(a, i)

    While i > 1 Or j > 1
        i = i - 1: j = j - 1
        If i > 0 Then sum = sum + Mid$(a, i, 1)
        If j > 0 Then sum = sum + Mid$(b, j, 1)
        ExactSum = sum Mod 10 & ExactSum
        sum = -(sum > 9)
    Wend
    
    If sum Then ExactSum = sum & ExactSum
      
End Function





Public Function CompareTypes$(ByVal n#)

    Dim t$
    
    t = t & vbLf
    t = t & CompareTypes_BitPattern(n, True)
    t = t & CompareTypes_Evaluation(n)
    t = t & CompareTypes_SubnormalAltRepresentation(n)
    t = t & CompareTypes_Binade(n)
    t = t & CompareTypes_FloatNumberLine(n)
    t = t & CompareTypes_QuantizationError(n)
    
    't = t & "? " & CSingle.Value_Exact & " - " & CSingle & vbLf & vbLf
    
    't = t & CSingle.Value_QuantizationError_DoubleRef(n)
    
    CompareTypes = t

End Function


Public Sub SetTypesEqualToValue(n#)
    
    CSingle.Mutate n
    CDouble.Mutate n

End Sub


Public Function CompareTypes_BitPattern$(n#, Optional include_conncectors As Boolean)

    Dim s$, d$
    
    'CSingle.Mutate n
    'CDouble.Mutate n
    SetTypesEqualToValue n
    
    s = CSingle.Raw_Entire_Binary(1)
    d = CDouble.Raw_Entire_Binary(1)
    s = Left$(s, 1) & "    " & Mid$(s, 3)

    s = "   0b" & s
    d = "   0b" & d
    
    s = Left$(s, 6) & "        " & Mid$(s, 8)
    d = Left$(d, 6) & "        " & Mid$(d, 8)
    
    s = Left$(s, 25) & "       " & Mid$(s, 26) & Space$(35) & CSingle.Value_HexFloatingPointLiteral
    d = Left$(d, 25) & "       " & Mid$(d, 26) & Space$(6) & CDouble.Value_HexFloatingPointLiteral
    
    s = s & Space$(118 - Len(s)) & "0x" & CSingle.Raw_Entire_Hexadecimal
    d = d & Space$(118 - Len(d)) & "0x" & CDouble.Raw_Entire_Hexadecimal
    

    Mid$(s, 15, 1) = Mid$(s, 18, 1): Mid$(s, 18, 1) = " "

    CompareTypes_BitPattern = s & vbLf & d & vbLf
    
    If include_conncectors Then
        CompareTypes_BitPattern = CompareTypes_BitPattern & "     |        |                  |" & Space$(60) & vbLf
        CompareTypes_BitPattern = CompareTypes_BitPattern & "     |        |                  |" & Space$(60) & vbLf
    End If
    
End Function

Public Function CompareTypes_Evaluation$(n#)

    Dim s$, d$, p&
    
    'CSingle.Mutate n
    'CDouble.Mutate n
    SetTypesEqualToValue n
    
    s = CSingle.Value_Formula_Evaluation
    d = CDouble.Value_Formula_Evaluation
        
    s = Replace(s, "*", " * ")
    p = InStr(1 + 0, s, "*")
    p = InStr(p + 1, s, "*")
    s = Left$(s, p - 1) & Space$(29 - p) & "*" & Mid$(s, p + 1)
    s = s & Space$(91 - Len(s)) & CSingle & " = " & CSingle.Value_IntegerRatio & vbLf
    
    d = Replace(d, "*", " * ")
    p = InStr(1 + 0, d, "*")
    p = InStr(p + 1, d, "*")
    d = Left$(d, p - 1) & Space$(29 - p) & "*" & Mid$(d, p + 1)
    d = d & Space$(91 - Len(d)) & CDouble & " = " & CDouble.Value_IntegerRatio & vbLf
        
    CompareTypes_Evaluation = s & d & vbLf

End Function


Public Function CompareTypes_Binade$(n#)

    Dim s$, d$, p&, q&, t$
    
    'CSingle.Mutate n
    'CDouble.Mutate n
    SetTypesEqualToValue n
    
    s = Replace(CSingle.Value_BinadeIntervalVerbose, "|", "")
    d = Replace(CDouble.Value_BinadeIntervalVerbose, "|", "")
    
    
    p = InStr(1, d, "Index")
    If p > InStr(1, s, "Index") Then
        s = Replace(s, "Index", Space$(p - InStr(1, s, "Index")) & "Index")
    End If
    
    p = InStr(1, d, "of")
    s = Replace(s, "of", Space$(p - InStr(1, s, "of")) & "of")
    
    p = InStrRev(d, " ")
    q = InStrRev(s, " ")
    If p <> q Then s = Left$(s, q) & Space$(p - q - 1) & Mid$(s, q)

    If LenB(s) <> LenB(d) Then
        If LenB(s) < LenB(d) Then
            t = t & "  " & s & Space$(Len(d) - Len(s) + 2) & "ULP: " & CSingle.Value_ULP & vbLf
            t = t & "  " & d & Space$(2) & "ULP: " & CDouble.Value_ULP & vbLf & vbLf
        End If
    Else
        t = t & "  " & s & Space$(2) & "ULP: " & CSingle.Value_ULP & vbLf
        t = t & "  " & d & Space$(2) & "ULP: " & CDouble.Value_ULP & vbLf & vbLf
    End If
    
    CompareTypes_Binade = t

End Function


Public Function FormatRemoveSciNotation$(n#)
    Dim e&, m$, v
    
    Select Case True
        Case InStr(n, "E-")
            v = Split(CStr(n), "-")
            m = v(0)
            m = Left$(m, Len(m) - 1)
            m = Replace(m, ".", vbNullString)
            e = v(1)
            FormatRemoveSciNotation = "0." & String$(e - 1, "0") & m
            
        Case InStr(n, "E+")
            v = Split(CStr(n), "+")
            m = v(0)
            m = Left$(m, Len(m) - 1)
            m = Replace(m, ".", vbNullString)
            e = v(1)
            FormatRemoveSciNotation = m & String$(e - Len(m) + 1, "0")
            
        Case Else
            FormatRemoveSciNotation = n
    End Select
End Function


Public Function CompareTypes_QuantizationError$(n#)

    Dim s$, d$, p&, t$, z$
    
    'CSingle.Mutate n
    'CDouble.Mutate n
    SetTypesEqualToValue n

    
    s = CSingle.Value_QuantizationError_DoubleRef(n, 21)
        
    s = "Quantization Error: " & s
    z = CSingle.Value_Exact(False): If Len(z) > 21 Then z = CSingle.Value_Exact(True)
    s = s & "    Exact Stored Value: " & z & vbLf
    
    
    
    
    d = CDouble.Value_QuantizationError(21)
    
    d = "Quantization Error: " & d
    z = CDouble.Value_Exact(False): If Len(z) > 21 Then z = CDouble.Value_Exact(True)
    d = d & "    Exact Stored Value: " & z & vbLf
    
    
    
    
    If InStr(1, d, "Exact Stored Value") >= InStr(1, s, "Exact Stored Value") Then
        p = InStr(1, d, "Exact Stored Value")
        s = Replace(s, "Exact Stored Value", Space$(p - InStr(1, s, "Exact Stored Value")) & "Exact Stored Value")
    Else
        p = InStr(1, s, "Exact Stored Value")
        d = Replace(d, "Exact Stored Value", Space$(p - InStr(1, d, "Exact Stored Value")) & "Exact Stored Value")
    End If
    
    t = t & s & d & vbLf

    CompareTypes_QuantizationError = t
    
End Function




Public Function CompareTypes_SubnormalAltRepresentation$(n#)

    Dim b$, s$, d$, j$, p&, q&, t$, z$
    
    'CSingle.Mutate n
    'CDouble.Mutate n
    SetTypesEqualToValue n

    s = vbNullString
    If CSingle.Value_Is_Subnormal Then
        b = Trim$(CSingle.Raw_Entire_Binary(1, , True))
        b = "     " & Left$(b, 1) & "        " & Mid$(b, 3)
        p = InStr(9, b, "1 ")
        j = "     |        |                                  "
        If p Then
            Mid$(j, p, 3) = "| |"
        Else
            p = Len(b)
            Mid$(j, p, 1) = "|"
        End If
        
        z = Replace(CSingle.Value_Formula_Evaluation(, , True), "* 2", "  * 2") & "   =   " & CSingle
        q = InStr(z, "* 1")
        z = Left$(z, q - 1) & Space$(p - q - 2) & Mid$(z, q)
    
        s = s & "--binary32 Subnormal Alternative Normalized Interpretation---------------------------" & vbLf
        s = s & b & vbLf
        s = s & j & vbLf
        s = s & z & vbLf
        s = s & vbLf
        s = s & "  " & Replace(CSingle.Value_BinadeIntervalVerbose(True), "|", vbNullString) & vbLf
        s = s & "-------------------------------------------------------------------------------------" & vbLf
        s = s & vbLf
    End If
    
    
    's = vbNullString
    If CDouble.Value_Is_Subnormal Then
        b = Trim$(CDouble.Raw_Entire_Binary(1, , True))
        b = "     " & Left$(b, 1) & "        " & Mid$(b, 3)
        p = InStr(9, b, "1 ")
        j = "     |        |                                                                       "
        If p Then
            Mid$(j, p, 3) = "| |"
        Else
            p = Len(b)
            Mid$(j, p, 1) = "|"
        End If
        
        z = Replace(CDouble.Value_Formula_Evaluation(, , True), "* 2", "  * 2") & "   =   " & CDouble
        q = InStr(z, "* 1")
        z = Left$(z, q - 1) & Space$(p - q - 2) & Mid$(z, q)
    
        s = s & "--binary64 Subnormal Alternative Normalized Interpretation-------------------------------------------------" & vbLf
        s = s & b & vbLf
        s = s & j & vbLf
        s = s & z & vbLf
        s = s & vbLf
        s = s & "  " & Replace(CDouble.Value_BinadeIntervalVerbose(True), "|", vbNullString) & vbLf
        s = s & "-----------------------------------------------------------------------------------------------------------" & vbLf
        s = s & vbLf
    End If

    CompareTypes_SubnormalAltRepresentation = s
    
End Function


Public Function CompareTypes_FloatNumberLine$(n#)

    Dim t$
    
    'CSingle.Mutate n
    'CDouble.Mutate n
    SetTypesEqualToValue n

    t = t & FloatNumberLine(CSingle.Value) & vbLf
    t = t & FloatNumberLine(CDouble.Value) & vbLf & vbLf
    
    CompareTypes_FloatNumberLine = t

End Function



Public Function FloatNumberLine$(n)
    Dim s$, p&, zero&, subna&, subnb&, normmin&, normmax&, one&, extentbeforeone&, extentpastone&
    Dim subnmidvalue, normminvalue, normunderone, maxmagexp&
    Dim legendleftmax$, legendleftmin$, legendrightmin$, legendrightmax$

    Select Case VarType(n)
        Case vbDouble
        
            subnmidvalue = 2 ^ ((-1074 + -1023) / 2)
            normminvalue = CDouble.Sample_NormalMin.Value
            normunderone = 2 ^ -4
            maxmagexp = 308
            legendleftmax = "-E+" & maxmagexp
            legendleftmin = "-E-" & maxmagexp
            legendrightmin = "+E-" & maxmagexp
            legendrightmax = "+E+" & maxmagexp
            
            
        
        Case vbSingle
        
            subnmidvalue = 2 ^ ((-149 + -127) / 2)
            normminvalue = CSingle.Sample_NormalMin.Value
            normunderone = 2 ^ -2
            maxmagexp = 38
            legendleftmax = " -E+" & maxmagexp
            legendleftmin = " -E-" & maxmagexp
            legendrightmin = "+E-" & maxmagexp & " "
            legendrightmax = "+E+" & maxmagexp & " "
        
        Case Else
            Stop
        
    End Select


    zero = 67
    subna = 72
    subnb = subna + 1
    normmin = subnb + 1
    extentbeforeone = 12
    extentpastone = 42
    one = normmin + extentbeforeone
    normmax = one + extentpastone


    s = s & "                                                                                                                                     " & vbLf
    s = s & "     |        ‹    ‹    ‹    ‹  ‹  ‹  ‹ ‹ ‹ ‹‹‹|«««««««««««|……    |    ……|»»»»»»»»»»»|››› › › ›  ›  ›  ›    ›    ›    ›        |     " & vbLf
    's = s & "     |                                         |           |      |      |           |                                         |     " & vbLf
    s = s & "-+++++                                        -1      ------      0      +-----      +1                                        ++++++" & vbLf

    s = Replace(s, "-+++++", legendleftmax)
    s = Replace(s, "------", legendleftmin)
    s = Replace(s, "+-----", legendrightmin)
    s = Replace(s, "++++++", legendrightmax)

    If VarType(n) = vbSingle Then
        If CSingle.Func_isNaN(CSng(n)) Then
            p = normmax + 3
            GoTo TheExit
        End If
        If CSingle.Func_isInfinity(CSng(n)) Then
            p = normmax + 1
            GoTo TheExit
        End If
    End If
    
    If VarType(n) = vbDouble Then
        If CDouble.Func_isNaN(CDbl(n)) Then
            p = normmax + 3
            GoTo TheExit
        End If
        If CDouble.Func_isInfinity(CDbl(n)) Then
            p = normmax + 1
            GoTo TheExit
        End If
    End If
        
        
    If n = 0 Then
        p = zero
    Else
    
        Select Case Abs(n)
            Case Is < subnmidvalue:   p = subna
            Case Is < normminvalue:   p = subnb
            
            Case Is < normunderone
                p = one - 1 - (extentbeforeone - 1) * EaseOutSine(((Log(Abs(n)) / Log(10)) / (-maxmagexp / (one - normmin))) / (extentbeforeone + 2))
            
            Case Else
                p = one + extentpastone * EaseOutSine(((Log(Abs(n)) / Log(10)) / (maxmagexp / (normmax - one))) / extentpastone)
            
        End Select
        
    End If
    
    
TheExit:
    If Left(n, 1) = "-" Then p = zero - (p - zero)
    Mid$(s, p, 1) = "x"
    FloatNumberLine = s

End Function


Function EaseOutSine#(x#)
    EaseOutSine = Sin((x * CFloat.PI_VariantDecimal) / 2)
End Function

Function EaseOutExpo#(x#)
    If x = 1 Then EaseOutExpo = 1: Exit Function
    EaseOutExpo = 1 - CFloat.Power(2, -10 * x)
End Function













'---------------------------------------------------------------------------
' Power Caches
'---------------------------------------------------------------------------


Public Function E_Notation$(ByVal n$, Optional lowercase As Boolean)
    Dim p&, marker$, n_original$, bNeg As Boolean

    n_original = n
    
    n = Trim$(n)
    bNeg = Left$(n, 1) = "-"
    If bNeg Then n = Mid$(n, 2)
    
    
    n = ForceStringToDecimalDigitsOnly(n)
    n = TrimLeadingZeros(n)
    
    marker = "E": If lowercase Then marker = "e"
    
    p = InStr(n, ".")
    If InStr(p + 1, n, ".") Then E_Notation = n_original: Exit Function
    
    If p = 0 Then
        p = Len(n) - 1
        E_Notation = Left$(n, 1) & "." & TrimTrailingZeros(Mid$(n, 2)) & marker & "+" & p
        E_Notation = Replace(E_Notation, ".E", "E")
        If bNeg Then E_Notation = "-" & E_Notation
        Exit Function
    End If

    If p = 1 Then
        n = "0" & n
        For p = 3 To Len(n)
            If Mid$(n, p, 1) <> "0" Then Exit For
        Next
        E_Notation = Mid$(n, p, 1) & "." & TrimTrailingZeros(Mid$(n, p + 1)) & marker & "-" & p - 2
        E_Notation = Replace(E_Notation, ".E", "E")
        If bNeg Then E_Notation = "-" & E_Notation
        Exit Function
    End If

    If p > 2 Then
        E_Notation = Left$(n, 1) & "." & Mid$(n, 2, p - 2) & TrimTrailingZeros(Mid$(n, p + 1)) & marker & "+" & p - 2
        E_Notation = Replace(E_Notation, ".E", "E")
        If bNeg Then E_Notation = "-" & E_Notation
        Exit Function
    End If
    
    E_Notation = Replace(n, ".E", "E")
    If bNeg Then E_Notation = "-" & E_Notation

End Function



Public Function DegreesToRadians#(d#)
    DegreesToRadians = d * PI_VariantDecimal / 180
End Function

Public Function RadiansToDegrees#(d#)
    RadiansToDegrees = d * 180 / PI_VariantDecimal
End Function



Public Function RandomIntegerBetween(lo, hi)
    RandomIntegerBetween = CFloat.Floor((hi - lo + 1) * Rnd + lo)
End Function

Public Function PowerOf02Exact$(ByVal n&)
    Debug.Assert n >= -1074
    Debug.Assert n <= 1023 + 1
    
    Dim i&
    Static CachedIncremental_PowersOf02Exact$()

    If (Not Not CachedIncremental_PowersOf02Exact) = 0 Then
        CachedIncremental_PowersOf02Exact = CreateCacheIncremental_PowersOf02Exact
    End If

    PowerOf02Exact = CachedIncremental_PowersOf02Exact(n)
    If LenB(PowerOf02Exact) Then Exit Function
    
      
    CachedIncremental_PowersOf02Exact(n) = 1
    If n > 0 Then
        For i = 1 To n
            CachedIncremental_PowersOf02Exact(n) = ExactProduct(CachedIncremental_PowersOf02Exact(n), 2)
        Next
    Else
        For i = 1 To Abs(n)
            CachedIncremental_PowersOf02Exact(n) = ExactProduct(CachedIncremental_PowersOf02Exact(n), 0.5)
        Next
    End If
    
    PowerOf02Exact = CachedIncremental_PowersOf02Exact(n)

End Function
Private Function CreateCacheIncremental_PowersOf02Exact() As String()
    Dim cache$()
    
    ReDim cache(-1074 To 1024)

    CreateCacheIncremental_PowersOf02Exact = cache
End Function



Public Function PowerOf02VariantDecimal(ByVal n&)
    Debug.Assert n >= -93
    Debug.Assert n <= 95
    
    'Note: More accurate than PowerOf02Double() when:   +50 <= n <=   +95 ***
    'Note: More accurate than PowerOf02Double() when:   -20 >= n >=   -48 ***
    'Note: Equal accuracy w/  PowerOf02Double() when:   -23 <= n <=   +49
    'Note: PowerOf02Double() can handle larger range: -1074 <= n <= +1023
    
    Static Cached_PowersOf02VariantDecimal()

    If (Not Not Cached_PowersOf02VariantDecimal) = 0 Then
        Cached_PowersOf02VariantDecimal = CreateCache_PowersOf02VariantDecimal
    End If

    PowerOf02VariantDecimal = Cached_PowersOf02VariantDecimal(n)
End Function
Private Function CreateCache_PowersOf02VariantDecimal() As Variant()
    Dim i&, n&, cache()
    
    ReDim cache(-93 To 95)
    For n = LBound(cache) To UBound(cache)
        cache(n) = CDec(1)
        If n > 0 Then
            For i = 1 To n
                cache(n) = cache(n) * 2
            Next
        Else
            For i = 1 To Abs(n)
                cache(n) = cache(n) * 0.5
            Next
        End If
    Next
    CreateCache_PowersOf02VariantDecimal = cache
End Function




Public Function PowerOf02Double#(ByVal n&)
    Debug.Assert n >= -1074
    Debug.Assert n <= 1023
    
    Static Cached_PowersOf02Double#()

    If (Not Not Cached_PowersOf02Double) = 0 Then
        Cached_PowersOf02Double = CreateCache_PowersOf02Double
    End If

    PowerOf02Double = Cached_PowersOf02Double(n)
End Function
Private Function CreateCache_PowersOf02Double() As Double()
    Dim c&, cache#()
    
    ReDim cache(-1074 To 1023)
    For c = LBound(cache) To UBound(cache)
        cache(c) = 2 ^ c
    Next

    CreateCache_PowersOf02Double = cache
End Function




Public Function PowerOf16Double#(ByVal n&)
    Debug.Assert n >= 0
    Debug.Assert n <= 23
    
    Static Cached_PowersOf16Double#()

    If (Not Not Cached_PowersOf16Double) = 0 Then
        Cached_PowersOf16Double = CreateCache_PowersOf16Double
    End If

    PowerOf16Double = Cached_PowersOf16Double(n)
End Function
Private Function CreateCache_PowersOf16Double() As Double()
    Dim c&, cache#()
    
    ReDim cache(0 To 23)
    For c = LBound(cache) To UBound(cache)
        cache(c) = 16 ^ c
    Next
    cache(23) = CDec("4951760157141521099596496896")

    CreateCache_PowersOf16Double = cache
End Function





Public Function CreateCache_MantissaDigits()
    Dim i&, j&, s$, t$
    Dim op1&(), tot&(), mDigits()
    Const TEMPLATE$ = "0000000000000000000000000000000000000000000000000000"

    ReDim op1(1 To 9)
    ReDim tot(1 To 9)
    ReDim mDigits(1 To 52)

    op1(1) = 1
    For j = 1 To 5
        tot = Base1E6Add(op1, tot)
    Next
    s = Base1E6ToBase10(tot)
    s = Me.TrimLeadingZeros(s)
    t = TEMPLATE
    'Mid$(t, 1) = Mid$(s, 3)
    Mid$(t, 1) = s
    mDigits(1) = Base10ToBase1E6(t)
    op1 = tot
    
    For i = 2 To 52
        For j = 1 To 4
            tot = Base1E6Add(op1, tot)
        Next
        s = Base1E6ToBase10(tot)
        s = Me.TrimLeadingZeros(s)
        t = TEMPLATE
        Mid$(t, i - Len(s) + 1) = s
        mDigits(i) = Base10ToBase1E6(t)
        op1 = tot
    Next
    
    CreateCache_MantissaDigits = mDigits
End Function


