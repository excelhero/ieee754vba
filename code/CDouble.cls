VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Type FloatTemplate
    Value As Double
End Type

Private Type ByteArrayTemplate
    Elements(0 To 7) As Byte
End Type

Private Type StringTemplate
    Value As String * 4
End Type

Private Type Long2Template
    lo As Long
    hi As Long
End Type

Private theValue As FloatTemplate
Private theByteArray As ByteArrayTemplate
Private theBitArray&()
Private cached&



Private Const VBA_TYPE_NUMB& = vbDouble
Private Const VBA_TYPE_NAME$ = "Double"
Private Const IEEE_754_NAME$ = "binary64"
Private Const STORAGE_BITS& = 64
Private Const STORAGE_BYTES& = 8
Private Const BYTE_______HI& = 7

Private Const RAW_HEX_DIGITS& = 16
Private Const RAW_INTEGER_MAX_DECIMAL$ = "18446744073709551615"
Private Const RAW_INTEGER_MAX_DECIMAL_SANS_SIGN_BIT$ = "9221120237041090560"
Private Const RAW_MANTISSA_MAX_DECIMAL$ = "4503599627370495"
Private Const RAW_EXPONENT_MAX& = 2047
Private Const EXPONENT_MIN_SUBNORMAL_EXPANSION& = -1074
Private Const EXPONENT_MIN& = -1023
Private Const EXPONENT_MAX& = 1024
Private Const EXPONENT_BIAS& = 1023
Private Const EXPONENT_SIZE& = 11
Private Const EXPONENT_HEX_DIGITS& = 3
Private Const EXPONENT_BIT_LO& = 53
Private Const EXPONENT_BIT_HI& = 63
Private Const EXPONENT_BIAS_FORMULA$ = "Bias = 2^(k-1) - 1, where k = " & EXPONENT_SIZE & "-bit exponent" & vbLf & "Bias = 2^(" & EXPONENT_SIZE & "-1) - 1" & vbLf & "Bias = " & EXPONENT_BIAS
Private Const MANTISSA_SIZE& = 52
Private Const MANTISSA_HEX_DIGITS& = 13
Private Const MANTISSA_BIT_LO& = 1
Private Const MANTISSA_BIT_HI& = 52
Private Const QUIET_BIT& = MANTISSA_BIT_HI
Private Const SIGN_BIT& = STORAGE_BITS
Private Const VALUE_FORMULA_NORMAL_BASE10$ = "(-1)^sign_bit * 2^(exponent - bias) * 1.mantissa"
Private Const VALUE_FORMULA_SUBNRM_BASE10$ = "(-1)^sign_bit * 2^(1 - bias) * 0.mantissa"
Private Const VALUE_FORMULA_NORMAL_BASE02$ = "(-1)^sign_bit * 10^(exponent_base2 - bias_base2) * 1.mantissa_base2"
Private Const VALUE_FORMULA_SUBNRM_BASE02$ = "(-1)^sign_bit * 10^(00000000001 - bias_base2) * 0.mantissa_base2"

Private Const NAN_PAYLOAD_MAX_VALUE$ = "2251799813685247"
Private Const NAN_PAYLOAD_MAX_BYTES& = 6

Private Const HEX_INFINITY__POSITIVE$ = "7FF0000000000000"
Private Const HEX_ZERO______NEGATIVE$ = "8000000000000000"
Private Const HEX_SNAN______POSITIVE$ = "7FF0000000000001"
Private Const HEX_QNAN______POSITIVE$ = "7FF8000000000001"
Private Const HEX_QNAN_INDETERMINATE$ = "FFF8000000000001"

Private Const HEX_NORM_MAX__POSITIVE$ = "7FEFFFFFFFFFFFFF"
Private Const HEX_NORM_MIN__POSITIVE$ = "0010000000000000"
Private Const HEX_SUBNORMMAXPOSITIVE$ = "000FFFFFFFFFFFFF"
Private Const HEX_SUBNORMMINPOSITIVE$ = "0000000000000001"

Private Const HEX_MAX_CONSEC_INT_POS$ = "4340000000000000"
Private Const HEX_MAX_CONSEC_INT_NEG$ = "C340000000000000"

Private Const HEX_MAX_SAFE___INT_POS$ = "433FFFFFFFFFFFFF"
Private Const HEX_MAX_SAFE___INT_NEG$ = "C33FFFFFFFFFFFFF"


Private Const HEX_PI________POSITIVE$ = "400921FB54442D18"
Private Const HEX_E_________POSITIVE$ = "4005BF0A8B145769"




Public Function Func_calculateQuotient(op1#, op2#)
    Dim s_op1, e_op1_b, e_op1_d, m_op1, i_op1
    Dim s_op2, e_op2_b, e_op2_d, m_op2, i_op2
    Dim s_res, e_res, m_res, m_size_x2&
    Dim s$, tmp$, blank$, headertemplate$, newrowtemplate$, rounded$, normshift$
    Dim bRequiresNormalization As Boolean, bSubnormal As Boolean, bUnderflow As Boolean, bOverflow As Boolean
    Dim res_fromops, res_frombitpattern, mutatebitpattern$
    Dim quot#, fields
    
    
    Const LINE_WIDTH& = 199, TAB1& = 27, TAB2& = 40, TAB3& = 60
    'Const LINE_WIDTH& = 143, TAB1& = 27, TAB2& = 40, TAB3& = 57
    
    m_size_x2 = 2 * MANTISSA_SIZE
    
    On Error Resume Next
        quot = op1 / op2
    On Error GoTo 0
    
    
    With CFloat
        .EnumerateFloat op1, Me, s_op1, e_op1_b, e_op1_d, m_op1, i_op1
        .EnumerateFloat op2, Me, s_op2, e_op2_b, e_op2_d, m_op2, i_op2
        
        
        s_res = s_op1 Xor s_op2
        
        e_res = e_op1_d - e_op2_d + EXPONENT_BIAS
        If e_res < 0 Then bUnderflow = True
        If e_res >= RAW_EXPONENT_MAX Then bOverflow = True
        
        m_res = .ExactProduct(CDec(.BinToDec(i_op1 & m_op1)) / CDec(.BinToDec(i_op2 & m_op2)), .PowerOf02Exact(m_size_x2))
        'm_res = (CDec(.BinToDec(i_op1 & m_op1)) / CDec(.BinToDec(i_op2 & m_op2))) * .PowerOf02Exact(m_size_x2)
        
        m_res = .DecToBinHuge(m_res)
        If Len(m_res) = m_size_x2 Then
            m_res = "0" & m_res
            bRequiresNormalization = True
        End If
        
        
        ReDim fields(1 To 3)
        fields(1) = Array(TAB1, TAB1 + 5 - 1, "Right")
        fields(2) = Array(TAB2, TAB2 + 13 - 1, "Split")
        fields(3) = Array(TAB3, TAB3 + 139 - 1, "Decimal", TAB3 + 2)
        
        blank = String$(LINE_WIDTH, " ") & vbLf
        newrowtemplate = .StringOverwrites(blank, "|", TAB1 + 7, "|", TAB3 - 5)
        headertemplate = .StringOverwrites(String$(LINE_WIDTH, "_") & vbLf, "|", TAB1 + 7, "|", TAB3 - 5)
    


        s = s & .StringOverwrites(blank, "a = " & op1, 1)
        s = s & .StringOverwrites(blank, "b = " & op2, 1)
        s = s & blank
        s = s & .StringOverwrites(blank, "a / b = " & quot, 1)
        s = s & blank
    
    
    
        'Bits___________________Sign__|_______Exponent__|_______Mantissa____________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Bits", 1, "Sign", TAB1, "Exponent", TAB2, "Mantissa", TAB3)
        s = s & .StringOverwriteX(newrowtemplate, fields, "a", 2, s_op1, TAB1, e_op1_b, TAB2, m_op1, TAB3)
        s = s & .StringOverwriteX(newrowtemplate, fields, "b", 2, s_op2, TAB1, e_op2_b, TAB2, m_op2, TAB3)
        s = s & blank
        
        
                
        'Quotient Sign___________Sign__|_________________|___________________________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Quotient Sign", 1, "Sign", TAB1)
        s = s & .StringOverwriteX(newrowtemplate, fields, "a", 2, s_op1, TAB1)
        s = s & .StringOverwriteX(newrowtemplate, fields, "b", 2, "xor " & s_op2, TAB1)
        s = s & .StringOverwriteX(newrowtemplate, fields, "-----", TAB1)
        s = s & .StringOverwriteX(newrowtemplate, fields, "result", 2, s_res, TAB1)
        s = s & blank
                        
        
        
        'Quotient Exponent_____________|_______Exponent__|___________________________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Quotient Exponent", 1, "Exponent", TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "a", 2, e_op1_b, TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "b", 2, "- " & e_op2_b, TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "bias " & EXPONENT_BIAS, 2, "+ " & "0" & String$(EXPONENT_SIZE - 1, "1"), TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, String$(EXPONENT_SIZE + 2, "-"), TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "result", 2, IIf(bUnderflow, "-", " ") & .DecToBin(Abs(e_res), IIf(e_res <= RAW_EXPONENT_MAX, EXPONENT_SIZE, 1 + EXPONENT_SIZE)), TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "base10", 3, e_res, TAB2)
        s = s & blank
    
    
    
        'Quotient Mantissa_____________|_________________|_______Mantissa____________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Quotient Mantissa", 1, "Mantissa", TAB3)
        s = s & .StringOverwriteX(newrowtemplate, fields, "a", 2, i_op1 & "." & m_op1, TAB3)
        s = s & .StringOverwriteX(newrowtemplate, fields, "b", 2, i_op2 & "." & m_op2, TAB3, "÷", TAB3 - 2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "---------------------------------------------------------¦---------------------------------------------------", TAB3 - 2)
        
                tmp = Left$(m_res, Len(m_res) - m_size_x2) & "." & Right$(m_res, m_size_x2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "result", 2, tmp, TAB3)
    
        Select Case True
            Case e_res < 0
                If e_res > (EXPONENT_MIN_SUBNORMAL_EXPANSION - 1) Then bSubnormal = True
                normshift = " * 2^(" & e_res & ")"
                m_res = String$(-e_res - -0, "0") & m_res
                s = s & .StringOverwriteX(newrowtemplate, fields, "subnormalization:" & normshift, 2, "0." & m_res, TAB3)
                
            Case e_res = 0
                If Len(m_res) < (m_size_x2 + 2) Then
                    bSubnormal = True
                    normshift = " * 2^(-1)"
                    s = s & .StringOverwriteX(newrowtemplate, fields, "subnormalization:" & normshift, 2, "0." & m_res & normshift, TAB3)
                Else
                    bSubnormal = False
                    bRequiresNormalization = True
                    normshift = " * 2^(-1)"
                    m_res = Mid$(m_res, 2)
                    s = s & .StringOverwriteX(newrowtemplate, fields, "normalization:" & normshift, 2, Left$(m_res, 1) & "." & m_res & normshift, TAB3)
                End If
                
            Case e_res = 1
                Select Case Mid$(m_res, 1, 1)
                    Case "1": bSubnormal = False
                    Case "0": bSubnormal = True
                End Select
                m_res = Mid$(m_res, 2)
                
                
            Case bOverflow
                    If (Not bRequiresNormalization) Or (e_res > RAW_EXPONENT_MAX) Then
                        m_res = String$(MANTISSA_SIZE, "0")
                        s = s & .StringOverwriteX(newrowtemplate, fields, "overflow", 2, m_res, TAB3)
                    Else
                        m_res = Mid$(m_res, 3)
                        normshift = " * 2^(1)"
                        tmp = "1." & m_res & normshift
                        bOverflow = False
                    End If
                    
            Case Else
                If bRequiresNormalization Then
                    If e_res > 0 Then
                        m_res = Mid$(m_res, 3)
                        normshift = " * 2^(1)"
                        tmp = "1." & m_res & normshift
                    Else
                        'Does this ever happen????
                        Stop
                        '''normshift = " * 2^(-1)"
                        '''tmp = "1." & Mid$(m_res, 3) & normshift
                    End If
                    s = s & .StringOverwriteX(newrowtemplate, fields, "normalization:" & normshift, 2, tmp, TAB3)
                Else
                    m_res = Mid$(m_res, 2)
                End If
                
        End Select
    
        If Not bOverflow Then
            rounded = .RoundToNearestNeighbor(m_res, MANTISSA_SIZE)
            If e_res < 0 Then
                tmp = "0." & rounded
            ElseIf e_res = 0 Or e_res = 1 Then
                If bSubnormal Then
                    tmp = "0." & rounded
                Else
                    tmp = "1." & rounded
                End If
            Else
                tmp = "1." & rounded & IIf(bRequiresNormalization, " * 2^(1)", vbNullString)
            End If
            s = s & .StringOverwriteX(newrowtemplate, fields, "rounded", 2, tmp, TAB3)
        Else
            rounded = m_res
        End If
        s = s & .StringOverwriteX(newrowtemplate, fields, "encoding" & IIf(bUnderflow, " (underflow)", vbNullString), 2, rounded, TAB3)
        s = s & blank
    
    
    
        'Quotient Exp Normalization____|_______Exponent__|___________________________________________________________________________________________
        If (bRequiresNormalization Or bOverflow) Then
            s = s & .StringOverwriteX(headertemplate, fields, "Quotient Exp Normalization", 1, "Exponent", TAB2)
            s = s & .StringOverwriteX(newrowtemplate, fields, "base10", 2, e_res, TAB2)
    
            If Not bOverflow Then
                If bUnderflow Then
                    s = s & .StringOverwriteX(newrowtemplate, fields, "mantissa shift:" & normshift, 2, "- " & e_res, TAB2)
                Else
                    s = s & .StringOverwriteX(newrowtemplate, fields, "mantissa shift:" & normshift, 2, "- 1", TAB2)
                End If
                s = s & .StringOverwriteX(newrowtemplate, fields, String$(EXPONENT_SIZE + 2, "-"), TAB2)
                s = s & .StringOverwriteX(newrowtemplate, fields, "result", 2, e_res - IIf(bUnderflow, e_res, 1), TAB2)
                s = s & .StringOverwriteX(newrowtemplate, fields, "base2", 3, .DecToBin(e_res - IIf(bUnderflow, -e_res, -bRequiresNormalization), EXPONENT_SIZE), TAB2)
            
            Else
                s = s & .StringOverwriteX(newrowtemplate, fields, "overflow", 2, String$(EXPONENT_SIZE, "1"), TAB2)
            End If
        End If
        s = s & blank
        
        
        
        'Quotient Encoded________Sign__|_______Exponent__|_______Mantissa____________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Quotient Encoded", 1, "Sign", TAB1, "Exponent", TAB2, "Mantissa", TAB3)
        Select Case True
            Case bUnderflow Or bSubnormal:  e_res = String$(EXPONENT_SIZE, "0")
            Case bOverflow:                 e_res = String$(EXPONENT_SIZE, "1")
            Case Else:                      e_res = .DecToBin(e_res - -bRequiresNormalization, EXPONENT_SIZE)
        End Select
        s = s & .StringOverwriteX(newrowtemplate, fields, "bits", 2, s_res, TAB1, e_res, TAB2, rounded, TAB3)
        s = s & blank
        
        
                                        
                                        
       '------------------------------------------------------------------------------------------------------------------------------------------------------
    
        mutatebitpattern = "0b " & s_res & " " & e_res & " " & rounded
        Me.Compare.Compare.Mutate mutatebitpattern
        res_frombitpattern = Me.Compare.Compare.Value
        res_fromops = quot
        
        
        s = s & .StringOverwriteX(blank, fields, "a / b = CDouble.Mutate(""0b", 1, s_res, TAB1, e_res, TAB2, rounded & """) = " & (res_fromops = res_frombitpattern), TAB3)
        s = s & blank
        s = s & .StringOverwrites(blank, "a / b = " & quot & ": ", 1, "0x" & CDouble.Mutate(quot).Raw_Entire_Hexadecimal, 35)
        s = s & .StringOverwrites(blank, "a = " & op1 & ": ", 1, "0x" & CDouble.Mutate(op1).Raw_Entire_Hexadecimal, 35)
        s = s & .StringOverwrites(blank, "b = " & op2 & ": ", 1, "0x" & CDouble.Mutate(op2).Raw_Entire_Hexadecimal, 35)
        s = s & blank
    
    
    End With
    
    Func_calculateQuotient = s
    
        
        




'                                 111110100100000000000000
'                               x 110000010000000000000000
'                               ---------------------------
'                                 000000000000000000000000
'                 111110100100000000000000
'           111110100100000000000000
'          111110100100000000000000
'         -------------------------------------------------
'         101111001010101001000000000000000000000000000000



End Function




Public Function Func_calculateProduct(op1#, op2#)
    Dim s_op1, e_op1_b, e_op1_d, m_op1, i_op1
    Dim s_op2, e_op2_b, e_op2_d, m_op2, i_op2
    Dim s_res, e_res, m_res, m_size_x2&
    Dim s$, tmp$, blank$, headertemplate$, newrowtemplate$, rounded$, normshift$
    Dim bRequiresNormalization As Boolean, bSubnormal As Boolean, bUnderflow As Boolean, bOverflow As Boolean
    Dim res_fromops, res_frombitpattern, mutatebitpattern$
    Dim prod#, fields
    
    Const LINE_WIDTH& = 199, TAB1& = 27, TAB2& = 40, TAB3& = 60
    
    m_size_x2 = 2 * MANTISSA_SIZE
    
    On Error Resume Next
        prod = op1 * op2
    On Error GoTo 0
    
    
    With CFloat
        .EnumerateFloat op1, Me, s_op1, e_op1_b, e_op1_d, m_op1, i_op1
        .EnumerateFloat op2, Me, s_op2, e_op2_b, e_op2_d, m_op2, i_op2
        
        
        s_res = s_op1 Xor s_op2
        
        e_res = e_op1_d + e_op2_d - EXPONENT_BIAS
        If e_res < 0 Then bUnderflow = True
        If e_res >= RAW_EXPONENT_MAX Then bOverflow = True
        
        m_res = .DecToBinHuge(.ExactProduct(CDec(.BinToDec(i_op1 & m_op1)), CDec(.BinToDec(i_op2 & m_op2))))
        'm_res = .DecToBin(CDec(.BinToDec(i_op1 & m_op1)) * CDec(.BinToDec(i_op2 & m_op2)))
        bRequiresNormalization = Len(m_res) > (m_size_x2 + 1)
        
        
        ReDim fields(1 To 3)
        fields(1) = Array(TAB1, TAB1 + 5 - 1, "Right")
        fields(2) = Array(TAB2, TAB2 + 13 - 1, "Split")
        fields(3) = Array(TAB3, TAB3 + 139 - 1, "Decimal", TAB3 + 2)
        
        blank = String$(LINE_WIDTH, " ") & vbLf
        newrowtemplate = .StringOverwrites(blank, "|", TAB1 + 7, "|", TAB3 - 5)
        headertemplate = .StringOverwrites(String$(LINE_WIDTH, "_") & vbLf, "|", TAB1 + 7, "|", TAB3 - 5)
    


        s = s & .StringOverwrites(blank, "a = " & op1, 1)
        s = s & .StringOverwrites(blank, "b = " & op2, 1)
        s = s & blank
        s = s & .StringOverwrites(blank, "a * b = " & prod, 1)
        s = s & blank
    
    
    
        'Bits___________________Sign__|_______Exponent__|_______Mantissa____________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Bits", 1, "Sign", TAB1, "Exponent", TAB2, "Mantissa", TAB3)
        s = s & .StringOverwriteX(newrowtemplate, fields, "a", 2, s_op1, TAB1, e_op1_b, TAB2, m_op1, TAB3)
        s = s & .StringOverwriteX(newrowtemplate, fields, "b", 2, s_op2, TAB1, e_op2_b, TAB2, m_op2, TAB3)
        s = s & blank
        
        
                
        'Product Sign___________Sign__|_________________|___________________________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Product Sign", 1, "Sign", TAB1)
        s = s & .StringOverwriteX(newrowtemplate, fields, "a", 2, s_op1, TAB1)
        s = s & .StringOverwriteX(newrowtemplate, fields, "b", 2, "xor " & s_op2, TAB1)
        s = s & .StringOverwriteX(newrowtemplate, fields, "-----", TAB1)
        s = s & .StringOverwriteX(newrowtemplate, fields, "result", 2, s_res, TAB1)
        s = s & blank
                        
        
        
        'Product Exponent_____________|_______Exponent__|___________________________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Product Exponent", 1, "Exponent", TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "a", 2, e_op1_b, TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "b", 2, "+ " & e_op2_b, TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "bias " & EXPONENT_BIAS, 2, "- " & "0" & String$(EXPONENT_SIZE - 1, "1"), TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, String$(EXPONENT_SIZE + 2, "-"), TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "result", 2, IIf(bUnderflow, "-", " ") & .DecToBin(Abs(e_res), IIf(e_res <= RAW_EXPONENT_MAX, EXPONENT_SIZE, 1 + EXPONENT_SIZE)), TAB2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "base10", 3, e_res, TAB2)
        s = s & blank
    
    
    
        'Product Mantissa_____________|_________________|_______Mantissa____________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Product Mantissa", 1, "Mantissa", TAB3)
        s = s & .StringOverwriteX(newrowtemplate, fields, "a", 2, i_op1 & "." & m_op1, TAB3)
        s = s & .StringOverwriteX(newrowtemplate, fields, "b", 2, i_op2 & "." & m_op2, TAB3, "x", TAB3 - 2)
        s = s & .StringOverwriteX(newrowtemplate, fields, "---------------------------------------------------------¦---------------------------------------------------", TAB3 - 2)
        
                tmp = Left$(m_res, Len(m_res) - m_size_x2) & "." & Right$(m_res, m_size_x2): If Len(tmp) < (m_size_x2 + 3) Then tmp = " " & tmp
        s = s & .StringOverwriteX(newrowtemplate, fields, "result", 2, tmp, TAB3)
    
        Select Case True
            Case e_res < 0
                If e_res > (EXPONENT_MIN_SUBNORMAL_EXPANSION - 1) Then bSubnormal = True
                normshift = " * 2^(" & e_res & ")"
                m_res = String$(-e_res - -bRequiresNormalization, "0") & m_res
                s = s & .StringOverwriteX(newrowtemplate, fields, "subnormalization:" & normshift, 2, "0." & m_res, TAB3)
                'Stop
                
            Case e_res = 0
                If Len(m_res) < (m_size_x2 + 2) Then
                    bSubnormal = True
                    normshift = " * 2^(-1)"
                    s = s & .StringOverwriteX(newrowtemplate, fields, "subnormalization:" & normshift, 2, "0." & m_res & normshift, TAB3)
                Else
                    bSubnormal = False
                    bRequiresNormalization = True
                    normshift = " * 2^(-1)"
                    m_res = Mid$(m_res, 2)
                    s = s & .StringOverwriteX(newrowtemplate, fields, "normalization:" & normshift, 2, Left$(m_res, 1) & "." & m_res & normshift, TAB3)
                End If
                
            Case bOverflow
                    m_res = String$(MANTISSA_SIZE, "0")
                    s = s & .StringOverwriteX(newrowtemplate, fields, "overflow", 2, m_res, TAB3)
            
            Case Else
                If bRequiresNormalization Then
                    If e_res > 0 Then
                        m_res = Mid$(m_res, 2)
                        normshift = " * 2^(1)"
                        tmp = "1." & m_res & normshift
                    Else
                        'Does this ever happen????
                        Stop
                        'normshift = " * 2^(-1)"
                        'tmp = "1." & Mid$(m_res, 2) & normshift
                    End If
                    s = s & .StringOverwriteX(newrowtemplate, fields, "normalization:" & normshift, 2, tmp, TAB3)
                Else
                    m_res = Mid$(m_res, 2)
                End If
                
        End Select
    
        If Not bOverflow Then
            rounded = .RoundToNearestNeighbor(m_res, MANTISSA_SIZE)
            If e_res < 0 Then
                tmp = "0." & rounded
            ElseIf e_res = 0 Then
                If bSubnormal Then
                    tmp = "0." & rounded
                Else
                    tmp = "1." & rounded
                End If
            Else
                tmp = "1." & rounded & IIf(bRequiresNormalization, " * 2^(1)", vbNullString)
            End If
            s = s & .StringOverwriteX(newrowtemplate, fields, "rounded", 2, tmp, TAB3)
        Else
            rounded = m_res
        End If
        s = s & .StringOverwriteX(newrowtemplate, fields, "encoding" & IIf(bUnderflow, " (underflow)", vbNullString), 2, rounded, TAB3)
        s = s & blank
    
    
    
        'Product Exp Normalization____|_______Exponent__|___________________________________________________________________________________________
        If (bRequiresNormalization Or bOverflow Or bUnderflow) Then
            s = s & .StringOverwriteX(headertemplate, fields, "Product Exp Normalization", 1, "Exponent", TAB2)
            s = s & .StringOverwriteX(newrowtemplate, fields, "base10", 2, e_res, TAB2)
    
            If Not bOverflow Then
                If bUnderflow Then
                    s = s & .StringOverwriteX(newrowtemplate, fields, "mantissa shift:" & normshift, 2, "- " & e_res, TAB2)
                Else
                    s = s & .StringOverwriteX(newrowtemplate, fields, "mantissa shift:" & normshift, 2, "+ 1", TAB2)
                End If
                s = s & .StringOverwriteX(newrowtemplate, fields, String$(EXPONENT_SIZE + 2, "-"), TAB2)
                s = s & .StringOverwriteX(newrowtemplate, fields, "result", 2, e_res + IIf(bUnderflow, -e_res, 1), TAB2)
                s = s & .StringOverwriteX(newrowtemplate, fields, "base2", 3, .DecToBin(e_res + IIf(bUnderflow, -e_res, -bRequiresNormalization), EXPONENT_SIZE), TAB2)
            
            Else
                s = s & .StringOverwriteX(newrowtemplate, fields, "overflow", 2, String$(EXPONENT_SIZE, "1"), TAB2)
            End If
        End If
        s = s & blank
        
        
        
        'Product Encoded________Sign__|_______Exponent__|_______Mantissa____________________________________________________________________________
        s = s & .StringOverwriteX(headertemplate, fields, "Product Encoded", 1, "Sign", TAB1, "Exponent", TAB2, "Mantissa", TAB3)
        Select Case True
            Case bUnderflow Or bSubnormal:  e_res = String$(EXPONENT_SIZE, "0")
            Case bOverflow:                 e_res = String$(EXPONENT_SIZE, "1")
            Case Else:                      e_res = .DecToBin(e_res + -bRequiresNormalization, EXPONENT_SIZE)
        End Select
        s = s & .StringOverwriteX(newrowtemplate, fields, "bits", 2, s_res, TAB1, e_res, TAB2, rounded, TAB3)
        s = s & blank
        
        
                                        
                                        
       '------------------------------------------------------------------------------------------------------------------------------------------------------
    
        mutatebitpattern = "0b " & s_res & " " & e_res & " " & rounded
        Me.Compare.Compare.Mutate mutatebitpattern
        res_frombitpattern = Me.Compare.Compare.Value
        res_fromops = prod
        
        
        s = s & .StringOverwriteX(blank, fields, "a * b = CDouble.Mutate(""0b", 1, s_res, TAB1, e_res, TAB2, rounded & """) = " & (res_fromops = res_frombitpattern), TAB3)
        s = s & blank
        s = s & .StringOverwrites(blank, "a * b = " & prod & ": ", 1, "0x" & CDouble.Mutate(prod).Raw_Entire_Hexadecimal, 35)
        s = s & .StringOverwrites(blank, "a = " & op1 & ": ", 1, "0x" & CDouble.Mutate(op1).Raw_Entire_Hexadecimal, 35)
        s = s & .StringOverwrites(blank, "b = " & op2 & ": ", 1, "0x" & CDouble.Mutate(op2).Raw_Entire_Hexadecimal, 35)
        s = s & blank
    
    
    End With
    
    Func_calculateProduct = s
    
        
        




'                                 111110100100000000000000
'                               x 110000010000000000000000
'                               ---------------------------
'                                 000000000000000000000000
'                 111110100100000000000000
'           111110100100000000000000
'          111110100100000000000000
'         -------------------------------------------------
'         101111001010101001000000000000000000000000000000



End Function






Public Sub Func_floatSplit(ByRef n#, Optional s, Optional e, Optional m, Optional decode As Boolean)
    Static lt As Long2Template
    Static ft As FloatTemplate
    ft.Value = n
    LSet lt = ft
    
    s = (lt.hi And &H80000000) \ &H80000000
    e = ((lt.hi And &H7FF00000) And -&H100000) \ &H100000
    
    m = CDec(lt.hi And &HFFFFF) * CFloat.PowerOf02Double(32)
    If lt.lo >= 0 Then
        m = m + lt.lo
    Else
        m = m + (lt.lo And &H7FFFFFFF) + &H7FFFFFFF + 1
    End If
    
    If decode Then
        Static b&(1 To STORAGE_BITS), subnormal As Boolean
        s = (-1) ^ s
        subnormal = 0 = e: If subnormal Then e = 1
        e = e - EXPONENT_BIAS
        m = CFloat.SignificandIntegerToBase10DecimalString(m, b, MANTISSA_BIT_LO, MANTISSA_BIT_HI, subnormal)
    End If
End Sub


Public Function Func_floatFromFieldValues_Calculated(ByVal s&, ByVal e&, ByVal m$, Optional exact As Boolean, Optional exactENotation As Boolean = True)
    Func_floatFromFieldValues_Calculated = CFloat.FloatFromFieldValues_Calculated(s, e, m, Me, -(exact = False) * STORAGE_BITS, exactENotation)
End Function


Public Function Func_floatNegate(ByVal n)
    Dim n_#, b() As Byte
    If Not IsNumeric(n) Then Exit Function
    n_ = n
    b = Reinterpret_FloatByRefAsByteArray(n_)
    b(BYTE_______HI) = b(BYTE_______HI) + &H80 * (-1) ^ (b(BYTE_______HI) > &H7F)
    Func_floatNegate = Reinterpret_ByteArrayAsFloat(b)
End Function

Public Function Func_floatSign&(ByRef n#)
    Static lt As Long2Template
    Static ft As FloatTemplate
    ft.Value = n
    LSet lt = ft
    
    Func_floatSign = (lt.hi And &H80000000) \ &H80000000
End Function

Public Function Func_floatSignValue&(ByRef n#)
    Func_floatSignValue = (-1) ^ Func_floatSign(n)
End Function


Public Function Func_floatExponent&(ByRef n#, Optional unbiased As Boolean)
    Static lt As Long2Template
    Static ft As FloatTemplate
    ft.Value = n
    LSet lt = ft
    
                 'exponent bitmask = &H7FF00000    ...then x >> 20
    Func_floatExponent = ((lt.hi And &H7FF00000) And -&H100000) \ &H100000
    
    If unbiased Then
        If Func_floatExponent = 0 Then Func_floatExponent = 1
        Func_floatExponent = Func_floatExponent - EXPONENT_BIAS
    End If
End Function


Public Function Func_floatSignificand(ByRef n#)
    Static lt As Long2Template
    Static ft As FloatTemplate, fs
    ft.Value = n
    LSet lt = ft
    
    fs = CDec(lt.hi And &HFFFFF) * CFloat.PowerOf02Double(32)
    If lt.lo >= 0 Then Func_floatSignificand = fs + lt.lo: Exit Function
    Func_floatSignificand = fs + (lt.lo And &H7FFFFFFF) + &H7FFFFFFF + 1
End Function

Public Function Func_nanPayload(ByRef n#)
    Static lt As Long2Template
    Static ft As FloatTemplate, fs
    ft.Value = n
    LSet lt = ft
    
    fs = CDec(lt.hi And &H7FFFF) * CFloat.PowerOf02Double(32)
    If lt.lo >= 0 Then Func_nanPayload = fs + lt.lo: Exit Function
    Func_nanPayload = fs + (lt.lo And &H7FFFFFFF) + &H7FFFFFFF + 1
End Function




Public Function Func_floatIntegerRatio$(n#, Optional num$, Optional den$)
    Dim e, m, mbits_used&

    If Me.Func_isZero(n) Then
        den = 1
        num = 0: If Me.Func_isZeroNegative(n) Then num = "-0"
        Func_floatIntegerRatio = num & " / " & den
        Exit Function
    End If
    
    If Me.Func_isInfinity(n) Then
        den = 0
        num = 1: If Me.Func_floatSign(n) Then num = -1
        Func_floatIntegerRatio = num & " / " & den
        Exit Function
    End If
    
    If Me.Func_isNaN(n) Then
        den = Me.Sample_QNaN(, Me.Func_floatSign(n))
        num = den
        Func_floatIntegerRatio = num & " / " & den
        Exit Function
    End If
    
    

    e = Me.Func_floatExponent(n, True)
    m = Me.Func_floatSignificand(n)
    
    If Me.Func_isNormal(n) Then
        den = CDec(CFloat.PowerOf02Exact(MANTISSA_SIZE))
        num = CFloat.ExactProduct(CFloat.PowerOf02Exact(e), m + den)
    Else
        mbits_used = Int(Log(m) / Log(2)) + 1
        den = CDec(CFloat.PowerOf02Exact(mbits_used))
        num = CFloat.ExactProduct(CFloat.PowerOf02Exact(e - (MANTISSA_SIZE - mbits_used)), m)
    End If
    
    If e < 0 Then

        num = "0" & CFloat.TrimLeadingZeros(num)

        Do While InStr(num, ".")
            num = CFloat.ExactProduct(num, 2)
            den = CFloat.ExactProduct(den, 2)
            If Right$(num, 2) = ".0" Then Exit Do
        Loop

        If Right$(num, 2) = ".0" Then num = Left$(num, Len(num) - 2)

        Func_floatIntegerRatio = String$(-(n < 0), "-") & num & " / " & den
    
        Exit Function
    End If
    
    
    
    Do
        num = CFloat.ExactProduct(num, 0.5)
        num = CFloat.TrimLeadingZeros(num)
        If Right$(num, 1) = 5 Then Exit Do
        
        den = CFloat.ExactProduct(den, 0.5)
        den = CFloat.TrimLeadingZeros(den)
        If Right$(den, 1) = 5 Then Exit Do
        DoEvents
    Loop While num > "1" And den > "1"
    
    If Right$(num, 2) = ".5" Then num = CFloat.ExactProduct(num, 2)
    If Right$(den, 2) = ".5" Then den = CFloat.ExactProduct(den, 2): num = CFloat.ExactProduct(num, 2)
    If Right$(num, 2) = ".0" Then num = Left$(num, Len(num) - 2)
    If Right$(den, 2) = ".0" Then den = Left$(den, Len(den) - 2)

    Func_floatIntegerRatio = String$(-(n < 0), "-") & num & " / " & den
End Function



Public Function Func_is2Power(x#) As Boolean
    If x = 0 Then Exit Function
    If Func_floatSign(x) Then Exit Function
    If Func_floatSignificand(x) Then Exit Function
    If Func_floatExponent(x) < RAW_EXPONENT_MAX Then Func_is2Power = True
End Function

Public Function Func_floatImplicitBit&(x#)
    Func_floatImplicitBit = -(0 <> Func_floatExponent(x))
End Function

Public Function Func_ulp#(x#)
    Me.Compare.Value = x
    Func_ulp = Me.Compare.Value_ULP
End Function


Public Function Func_isNegative(x#) As Boolean
    If Me.Func_floatSign(x) Then Func_isNegative = True
End Function

Public Function Func_isNaN(x#, Optional andNegative As Boolean) As Boolean
    If Func_isFinite(x) Then Exit Function
    If Func_floatSignificand(x) Then Func_isNaN = True
    If Func_isNaN And andNegative Then Func_isNaN = Func_floatSign(x)
End Function

Public Function Func_isInfinity(x#, Optional andNegative As Boolean) As Boolean
    If Func_isFinite(x) Then Exit Function
    If Func_floatSignificand(x) = 0 Then Func_isInfinity = True
    If Func_isInfinity And andNegative Then Func_isInfinity = Func_floatSign(x)
End Function

Public Function Func_isFinite(x#, Optional andNegative As Boolean) As Boolean
    If Func_floatExponent(x) < RAW_EXPONENT_MAX Then Func_isFinite = True
    If Func_isFinite And andNegative Then Func_isFinite = Func_floatSign(x)
End Function

Public Function Func_isSubnormal(x#, Optional andNegative As Boolean) As Boolean
    If x = 0 Then Exit Function
    If Func_floatExponent(x) = 0 Then Func_isSubnormal = True
    If Func_isSubnormal And andNegative Then Func_isSubnormal = Func_floatSign(x)
End Function

Public Function Func_isNormal(x#, Optional andNegative As Boolean) As Boolean
    If Not Func_isFinite(x) Then Exit Function
    If Func_floatExponent(x) Then Func_isNormal = True
    If Func_isNormal And andNegative Then Func_isNormal = Func_floatSign(x)
End Function

Public Function Func_isZero(x#) As Boolean
    If Func_floatExponent(x) Then Exit Function
    If Func_floatSignificand(x) = 0 Then Func_isZero = True
End Function

Public Function Func_isZeroNegative(x#) As Boolean
    If Not Func_isZero(x) Then Exit Function
    If Func_floatSign(x) = 1 Then Func_isZeroNegative = True
End Function

Public Function Func_isZeroPositive(x#) As Boolean
    If Not Func_isZero(x) Then Exit Function
    If Func_floatSign(x) = 0 Then Func_isZeroPositive = True
End Function





Public Property Get Spec_VBAType&()
    Spec_VBAType = VBA_TYPE_NUMB
End Property

Public Property Get Spec_VBATypeName$()
    Spec_VBATypeName = VBA_TYPE_NAME
End Property

Public Property Get Spec_IEEE754Name$()
    Spec_IEEE754Name = IEEE_754_NAME
End Property

Public Property Get Spec_BitsStorage&()
    Spec_BitsStorage = STORAGE_BITS
End Property

Public Property Get Spec_BytesStorage&()
    Spec_BytesStorage = STORAGE_BYTES
End Property

Public Property Get Spec_PrecisionBits&()
    Spec_PrecisionBits = MANTISSA_SIZE + 1
End Property

Public Property Get Spec_PrecisionDigitsBase10#()
    Spec_PrecisionDigitsBase10 = Spec_PrecisionBits& / (Log(10) / Log(2))
End Property

Public Property Get Spec_ExponentBits&()
    Spec_ExponentBits = EXPONENT_SIZE
End Property

Public Property Get Spec_ExponentMaxBiased&()
    Spec_ExponentMaxBiased = EXPONENT_MAX + EXPONENT_BIAS
End Property

Public Property Get Spec_ExponentMinBiased&()
    Spec_ExponentMinBiased = EXPONENT_MIN + EXPONENT_BIAS
End Property

Public Property Get Spec_ExponentMax&()
    Spec_ExponentMax = EXPONENT_MAX
End Property

Public Property Get Spec_ExponentMin&()
    Spec_ExponentMin = EXPONENT_MIN
End Property

Public Property Get Spec_ExponentMinBiased_SubnormalExpansion&()
    Spec_ExponentMinBiased_SubnormalExpansion = EXPONENT_MIN_SUBNORMAL_EXPANSION
End Property

Public Property Get Spec_ExponentBias_Base10&()
    Spec_ExponentBias_Base10 = EXPONENT_BIAS
End Property

Public Property Get Spec_ExponentBias_Base02$()
    Spec_ExponentBias_Base02 = CFloat.DecToBin(EXPONENT_BIAS, EXPONENT_SIZE)
End Property

Public Property Get Spec_ExponentBias_Base10Formula$()
    Spec_ExponentBias_Base10Formula = EXPONENT_BIAS_FORMULA
End Property

Public Property Get Spec_ExponentLoBit&()
    Spec_ExponentLoBit = EXPONENT_BIT_LO
End Property

Public Property Get Spec_ExponentHiBit&()
    Spec_ExponentHiBit = EXPONENT_BIT_HI
End Property

Public Property Get Spec_SignBit&()
    Spec_SignBit = SIGN_BIT
End Property

Public Property Get Spec_SignificandHiBit&()
    Spec_SignificandHiBit = MANTISSA_BIT_HI
End Property

Public Property Get Spec_SignificandLoBit&()
    Spec_SignificandLoBit = MANTISSA_BIT_LO
End Property

Public Property Get Spec_SignificandHexDigitCount&()
    Spec_SignificandHexDigitCount = MANTISSA_HEX_DIGITS&
End Property

Public Property Get Spec_SignificandMax()
    Spec_SignificandMax = CDec(RAW_MANTISSA_MAX_DECIMAL)
End Property

Public Property Get Spec_SignificandHexDigits&()
    Spec_SignificandHexDigits = MANTISSA_HEX_DIGITS
End Property

Public Property Get Spec_NanPayloadMaxValue()
    Spec_NanPayloadMaxValue = CDec(NAN_PAYLOAD_MAX_VALUE)
End Property


Public Property Get Value_Formula$(Optional level& = 2, Optional subnormal As Boolean = False, Optional base& = 10)
    
    If base = 10 Then
        Select Case Not subnormal
            Case True: Value_Formula = VALUE_FORMULA_NORMAL_BASE10
            Case Else: Value_Formula = VALUE_FORMULA_SUBNRM_BASE10
        End Select
        
        If level >= 1 Then Value_Formula = Replace(Value_Formula, "bias", Spec_ExponentBias_Base10)
        If level >= 2 Then Value_Formula = Replace(Value_Formula, "1 - " & Spec_ExponentBias_Base10, 1 - Spec_ExponentBias_Base10)
        If level < 3 Then Exit Property
    Else
        Select Case Not subnormal
            Case True: Value_Formula = VALUE_FORMULA_NORMAL_BASE02
            Case Else: Value_Formula = VALUE_FORMULA_SUBNRM_BASE02
        End Select
        
        If level >= 1 Then Value_Formula = Replace(Value_Formula, "bias_base2", Spec_ExponentBias_Base02)
        If level >= 2 Then Value_Formula = Replace(Value_Formula, String$(EXPONENT_SIZE - 1, "0") & "1 - " & Spec_ExponentBias_Base02, Value_Exponent_Base02)
        If level < 3 Then Exit Property
    End If
    
End Property


Public Property Get Value_Binade#(Optional return_exponent As Boolean, Optional subnormal_alt_representation As Boolean)
    Value_Binade = Func_binade(Me.Value, return_exponent, subnormal_alt_representation)
End Property

Public Property Get Value_BinadeNext#(Optional steps&, Optional return_exponent As Boolean, Optional subnormal_alt_representation As Boolean)
    Dim e&, i&
    
    If steps Then
        Compare.Value = Value
        If steps > 0 Then
            For i = 1 To steps
                Compare.Value = Compare.Value_BinadeNext
            Next
        Else
            For i = 1 To steps
                Compare.Value = Compare.Value_BinadePrev
            Next
        End If
        If Not return_exponent Then
            Value_BinadeNext = Compare.Value
        Else
            Value_BinadeNext = Compare.Value_Binade(True)
        End If
        Exit Property
    End If
    
    
    If Not return_exponent Then
        If Value <> 0 Then
            Value_BinadeNext = Me.Func_2PowerNext(Value)
        Else
            Value_BinadeNext = CFloat.PowerOf02Double(EXPONENT_MIN_SUBNORMAL_EXPANSION)
        End If
    Else
        If Value <> 0 Then
            Value_BinadeNext = Log(Me.Func_2PowerNext(Abs(Value), subnormal_alt_representation)) / Log(2)
        Else
            Value_BinadeNext = EXPONENT_MIN_SUBNORMAL_EXPANSION
        End If
    End If

End Property


Public Property Get Value_BinadePrev#(Optional steps&, Optional return_exponent As Boolean, Optional subnormal_alt_representation As Boolean)
    Dim e&, i&
    
    If steps Then
        Compare.Value = Value
        If steps > 0 Then
            For i = 1 To steps
                Compare.Value = Compare.Value_BinadePrev
            Next
        Else
            For i = 1 To steps
                Compare.Value = Compare.Value_BinadeNext
            Next
        End If
        If Not return_exponent Then
            Value_BinadePrev = Compare.Value
        Else
            Value_BinadePrev = Compare.Value_Binade(True)
        End If
        Exit Property
    End If
    
    
    If Not return_exponent Then
        If Value <> 0 Then
            Value_BinadePrev = Me.Func_2PowerPrev(Value)
        Else
            Value_BinadePrev = 0
        End If
    Else
        If Value <> 0 Then
            Value_BinadePrev = Log(Me.Func_2PowerPrev(Abs(Value)), subnormal_alt_representation) / Log(2)
        Else
            Value_BinadePrev = 0
        End If
    End If

End Property

Public Property Get Value_BinadeInterval$(Optional subnormal_alt_representation As Boolean)
    
    Dim lo#, hi#, lo_$, hi_$, t$
    
    lo = Value_Binade(True, subnormal_alt_representation)
    lo_ = "2^" & lo
    If Me.Value_Is_Negative Then lo_ = "-" & lo_
    If Me.Func_isNaN(lo) Then lo_ = lo
    If lo_ <> "nan" Then If lo < EXPONENT_MIN_SUBNORMAL_EXPANSION Then lo_ = 0: If Me.Value_Is_Negative Then lo_ = "-0"
    
    If Me.Func_isNaN(lo) Then
        hi_ = lo_
    Else
        hi = Value_BinadeNext(, True, subnormal_alt_representation)
        If Not Me.Func_isNaN(hi) Then
            hi_ = "2^" & hi
            If Me.Value_Is_Negative Then hi_ = "-" & hi_
        Else
            hi_ = hi
        End If
    End If
    
    t = "[" & lo_ & ", " & hi_
    If lo_ = "nan" And hi_ = "nan" Then t = t & "]" Else t = t & ")"
    Value_BinadeInterval = t
End Property

Public Property Get Value_BinadeIntervalVerbose$(Optional subnormal_alt_representation As Boolean)
    Value_BinadeIntervalVerbose = "Binade: " & Value_BinadeInterval(subnormal_alt_representation) & " | Index: " & Value_BinadeIndexVerbose(subnormal_alt_representation)
End Property

Public Property Get Value_BinadeIndex$(Optional subnormal_alt_representation As Boolean)
    If Not subnormal_alt_representation Then
        Value_BinadeIndex = Raw_Mantissa_Decimal
    Else
        Value_BinadeIndex = Raw_Mantissa_Decimal(Abs(EXPONENT_MIN_SUBNORMAL_EXPANSION) - Abs(Me.Value_Binade(True, True)))
    End If
End Property

Public Property Get Value_BinadeIndexVerbose$(Optional subnormal_alt_representation As Boolean)
    If Me.Value_Is_Finite Then
        If Value <> 0 Then
            If Not subnormal_alt_representation Then
                Value_BinadeIndexVerbose = Value_BinadeIndex & " of " & CDec(RAW_MANTISSA_MAX_DECIMAL) & ", " & Format$(Raw_Mantissa_Decimal / CDec(RAW_MANTISSA_MAX_DECIMAL), "0.00000%")
            Else
                Dim hibit&
                hibit = Abs(EXPONENT_MIN_SUBNORMAL_EXPANSION) - Abs(Me.Value_Binade(True, True))
                If hibit Then
                    Value_BinadeIndexVerbose = Value_BinadeIndex(True) & " of " & CDec(2 ^ hibit - 1) & ", " & Format$(Raw_Mantissa_Decimal(hibit) / CDec(2 ^ hibit - 1), "0.00000%")
                Else
                    Value_BinadeIndexVerbose = Value_BinadeIndex(True) & " of " & 1 & ", " & Format$(1, "0.00000%")
                End If
            End If
        Else
            Value_BinadeIndexVerbose = Value_BinadeIndex & " of " & 0 & ", " & Format$(0, "0.00000%")
        End If
    Else
        Value_BinadeIndexVerbose = Value_BinadeIndex & " of " & CDec(RAW_MANTISSA_MAX_DECIMAL) & ", " & Format$(Raw_Mantissa_Decimal / CDec(RAW_MANTISSA_MAX_DECIMAL), "0.00000%")
    End If
    
End Property


Public Property Get Value_Verbose$(Optional include_exact_value As Boolean, Optional subnormal_alt_representation As Boolean)
    Const GAP1& = 3
    Const GAP2$ = "    "
    With CFloat
        Value_Verbose = Value_Verbose & Raw_Entire_Binary(1)
        Value_Verbose = Value_Verbose & .RightJustify(Raw_Entire_Hexadecimal(prefix:="0x"), 18, GAP1)
        Value_Verbose = Value_Verbose & .RightJustify(Value_HexFloatingPointLiteral, 23, GAP1)
        Value_Verbose = Value_Verbose & .RightJustify(Value, 22, GAP1)
        Value_Verbose = Value_Verbose & Space$(GAP1) & Me.Value_IntegerRatio
        Value_Verbose = Value_Verbose & .SymbolJustify(Me.Value_BinadeIntervalVerbose(subnormal_alt_representation), GAP2)
        If include_exact_value Then Value_Verbose = Value_Verbose & .SymbolJustify(Value_Exact, GAP2)
        Value_Verbose = Value_Verbose & .SymbolJustify(Report_Threshold, GAP2)
    End With
End Property


Private Property Get Value_Sign_Symbol$(Optional suppressplus As Boolean = False)
    If Value_Sign > 0 Then
        If Not suppressplus Then Value_Sign_Symbol = "+"
    Else
        Value_Sign_Symbol = "-"
    End If
End Property

Public Property Get Value_Status$()
    Const DIGIT& = 2
    Static cache_$
    Dim s$
            
    If Not cached And DIGIT Then
        cached = cached Or DIGIT
        Select Case True
            Case Value_Is_Normal:    s = "Normal" & Report_NormalLimit & " | Binade" & Report_BinadeLimit & ": " & Value_BinadeInterval & " | ULP: " & Value_ULP
            Case Value_Is_Subnormal: s = "Subnormal" & Report_SubnormalLimit & " | Binade: " & Value_BinadeInterval & " | ULP: " & Value_ULP
            Case Value_Is_NaN:       s = "NaN " & IIf(Value_Is_QNaN, "Quiet" & IIf(Value_Is_QNaNIndeterminate, " Indeterminate", vbNullString), "Signaling")
            Case Value_Is_Infinity:  s = "Infinity"
            Case Value_Is_Zero:      s = "Zero"
        End Select
        cache_ = Value_Sign_Symbol & s
    End If
    Value_Status = cache_
End Property



Public Function Report_Threshold$()
    Dim s$
    Select Case True
        Case Value_Is_NormalMaxMagnitude:    s = "MaxNormal"
        Case Value_Is_NormalMinMagnitude:    s = "MinNormal"
        Case Value_Is_SubnormalMaxMagnitude: s = "MaxSubnormal"
        Case Value_Is_SubnormalMinMagnitude: s = "MinSubnormal"
        Case Value_Is_MaxSafeInteger:        s = "MaxSafeInteger"
        Case Value_Is_MaxConsecutiveInteger: s = "MaxConsecutiveInteger"
    End Select
    If Len(s) Then Report_Threshold = Value_Sign_Symbol & s
End Function



Private Function Report_NormalLimit$()
    Select Case True
        Case Value_Is_NormalMaxMagnitude: Report_NormalLimit = " Max Magnitude"
        Case Value_Is_NormalMinMagnitude: Report_NormalLimit = " Min Magnitude"
    End Select
End Function

Private Function Report_SubnormalLimit$()
    Select Case True
        Case Value_Is_SubnormalMaxMagnitude: Report_SubnormalLimit = " Max Magnitude"
        Case Value_Is_SubnormalMinMagnitude: Report_SubnormalLimit = " Min Magnitude"
    End Select
End Function

Private Function Report_BinadeLimit$()
    Select Case True
        Case Value_Is_NormalBinadeMaxMagnitude: Report_BinadeLimit = " Max Magnitude"
        Case Value_Is_NormalBinadeMinMagnitude: Report_BinadeLimit = " Min Magnitude"
    End Select
End Function

Public Property Get Value_Is_Negative() As Boolean
    If Me.Func_floatSign(Me.Value) Then Value_Is_Negative = True
End Property

Public Property Get Value_Is_Normal() As Boolean
    Dim e&
    e = Raw_Exponent_Biased_Decimal
    Value_Is_Normal = e > 0 And e < RAW_EXPONENT_MAX
End Property

Public Property Get Value_Is_NormalMinMagnitude() As Boolean
    Value_Is_NormalMinMagnitude = Value_Is_NormalBinadeMinMagnitude And (Value_Exponent_Base10 = EXPONENT_MIN + 1)
End Property

Public Property Get Value_Is_NormalMaxMagnitude() As Boolean
    Value_Is_NormalMaxMagnitude = Value_Is_NormalBinadeMaxMagnitude And (Value_Exponent_Base10 = EXPONENT_MAX - 1)
End Property

Public Property Get Value_Is_MaxSafeInteger() As Boolean
    Value_Is_MaxSafeInteger = Value = CFloat.HexStringToVariantFloat(HEX_MAX_SAFE___INT_POS, Me)
End Property

Public Property Get Value_Is_MaxConsecutiveInteger() As Boolean
    Value_Is_MaxConsecutiveInteger = Value = CFloat.HexStringToVariantFloat(HEX_MAX_CONSEC_INT_POS, Me)
End Property



Public Property Get Value_Is_NormalBinadeMinMagnitude() As Boolean
    If Value_Is_Normal Then Value_Is_NormalBinadeMinMagnitude = Raw_Mantissa_Decimal = 0
End Property

Public Property Get Value_Is_NormalBinadeMaxMagnitude() As Boolean
    If Value_Is_Normal Then Value_Is_NormalBinadeMaxMagnitude = Raw_Mantissa_Decimal = CDec(RAW_MANTISSA_MAX_DECIMAL)
End Property



Public Property Get Value_Is_2Power() As Boolean
    If Me.Value = 0 Then Exit Property
    If Me.Raw_SignBit Then Exit Property
    If Me.Raw_Mantissa_Decimal Then Exit Property
    If Me.Raw_Exponent_Biased_Decimal >= RAW_EXPONENT_MAX Then Exit Property
    Value_Is_2Power = True
End Property


Public Property Get Value_Is_Subnormal() As Boolean
    If Raw_Exponent_Biased_Decimal Then Exit Property
    If Raw_Mantissa_Decimal Then Value_Is_Subnormal = True
End Property

Public Property Get Value_Is_SubnormalMinMagnitude() As Boolean
    If Value_Is_Subnormal Then Value_Is_SubnormalMinMagnitude = Raw_Mantissa_Decimal = 1
End Property

Public Property Get Value_Is_SubnormalMaxMagnitude() As Boolean
    If Value_Is_Subnormal Then Value_Is_SubnormalMaxMagnitude = Raw_Mantissa_Decimal = CDec(RAW_MANTISSA_MAX_DECIMAL)
End Property

Public Property Get Value_Is_NaN() As Boolean
    If Raw_Exponent_Biased_Decimal <> RAW_EXPONENT_MAX Then Exit Property
    If Raw_Mantissa_Decimal Then Value_Is_NaN = True
End Property

Public Property Get Value_Is_SNaN() As Boolean
    Value_Is_SNaN = Value_Is_NaN And Raw_Bit(QUIET_BIT) = 0
End Property

Public Property Get Value_Is_QNaN() As Boolean
    Value_Is_QNaN = Value_Is_NaN And Raw_Bit(QUIET_BIT) = 1
End Property

Public Property Get Value_Is_QNaNIndeterminate() As Boolean
    Value_Is_QNaNIndeterminate = Value_Is_QNaN And Raw_Bit(SIGN_BIT) = 1
End Property

Public Property Get Value_Is_Finite() As Boolean
    If Raw_Exponent_Biased_Decimal < RAW_EXPONENT_MAX Then Value_Is_Finite = True
End Property

Public Property Get Value_Is_Infinity() As Boolean
    If Raw_Exponent_Biased_Decimal < RAW_EXPONENT_MAX Then Exit Property
    If Raw_Mantissa_Decimal = 0 Then Value_Is_Infinity = True
End Property


Public Property Get Value_Is_Zero() As Boolean
    If Value = 0 Then Value_Is_Zero = True
End Property

Public Property Get Value_Is_ZeroNegative() As Boolean
    Value_Is_ZeroNegative = Value = 0 And Value_Sign = -1
End Property

Public Property Get Value_Is_ZeroPositive() As Boolean
    Value_Is_ZeroPositive = Value = 0 And Value_Sign = 1
End Property



Public Property Get Value_Formula_Evaluation$(Optional level& = 3, Optional base& = 10, Optional SubnormalAlt As Boolean)
    Dim f$, b$, e$, m$, s$, bSubN As Boolean
        
    If Value_Is_Subnormal Then bSubN = True
    
    If level <= 2 Then
        If base = 10 Then
            If bSubN Then
                Value_Formula_Evaluation = Value_Formula(level, True)
            Else
                Value_Formula_Evaluation = Value_Formula(level, False)
            End If
        Else
            If bSubN Then
                Value_Formula_Evaluation = Value_Formula(level, True, 2)
            Else
                Value_Formula_Evaluation = Value_Formula(level, False, 2)
            End If
        End If
        Exit Property
    End If
    
    
    Select Case level
        Case 3
            s = "(-1)^" & Raw_SignBit
            e = "(e - b)": If bSubN Then e = "(1 - b)"
            Select Case base
                Case 10
                    e = Replace(e, "e", Raw_Exponent_Biased_Decimal)
                    e = Replace(e, "b", Spec_ExponentBias_Base10)
                    e = "2^" & e
                    m = Value_Mantissa_Base10(True)
                    If SubnormalAlt Then
                        If bSubN Then
                            e = "2^" & Me.Value_Binade(True, True)
                            m = Me.Value_Mantissa_Base10(False, Abs(EXPONENT_MIN_SUBNORMAL_EXPANSION) - Abs(Me.Value_Binade(True, True)))
                            Mid$(m, 1, 1) = "1"
                        End If
                    End If
                Case Else
                    If bSubN Then
                        e = Replace(e, "1", CFloat.DecToBin(1, Spec_ExponentBits))
                        e = Replace(e, "e", CFloat.DecToBin(1, Spec_ExponentBits))
                    End If
                    e = Replace(e, "e", Raw_Exponent_Biased_Binary)
                    e = "10^" & e
                    e = Replace(e, "b", Spec_ExponentBias_Base02)
                    m = Value_Mantissa_Base02(True)
            End Select
        Case Is >= 4
            s = Value_Sign
            e = "(e)"
            Select Case base
                Case 10
                    e = "2^" & e
                    e = Replace(e, "e", Value_Exponent_Base10)
                    m = Value_Mantissa_Base10(True)
                Case Else
                    e = "10^" & e
                    e = Replace(e, "e", Value_Exponent_Base02)
                    m = Value_Mantissa_Base02(True)
            End Select
            
            If level >= 5 Then
                If base = 10 Then e = CFloat.PowerOf02Exact(Value_Exponent_Base10)
            End If
            
            If level >= 6 Then
                If base = 10 Then
                    Value_Formula_Evaluation = Value_Sign_Symbol(True) & CFloat.ExactProduct(e, m, True)
                    Exit Property
                End If
            End If
        
    End Select
    
    f = "s * e * m"
    f = Replace(f, "s", s)
    f = Replace(f, "e", e)
    f = Replace(f, "m", m)
    Value_Formula_Evaluation = f
End Property



Public Property Get Compare() As CDouble
    Static o As CDouble
    If o Is Nothing Then Set o = New CDouble
    Set Compare = o
End Property


Public Property Get Value_PrecisionComp(n#)
    Value_PrecisionComp = Me.Func_precisionComp(Value, n)
End Property



Public Property Get Value_FindPrecisionComp#(ByVal p&)

    If p > Spec_PrecisionBits Then p = Spec_PrecisionBits
    
    Value_FindPrecisionComp = Value + (-1) ^ (p < 0) * Value_ULP * CFloat.PowerOf02Exact(MANTISSA_SIZE - Abs(p))

End Property




Public Property Get Value_QuantizationError(Optional UseENotationLongerThanDigits& = 21)
    Dim d$, z$
    
    If Me.Value_Is_NaN Then Value_QuantizationError = 0: Exit Function
    If Me.Value_Is_Zero Then Value_QuantizationError = 0: Exit Property
    
    If Me.Value_Is_NaN Or Me.Value_Is_Infinity Then
        d = 0
    Else
        d = Me.Func_formatRemoveSciNotation(Abs(Me.Value))
        If Right$(d, 2) = ".0" Then d = Left$(d, Len(z) - 2)
        z = Me.Value_Exact(False)
        z = Replace(z, "-", "")
        If Right$(z, 2) = ".0" Then z = Left$(z, Len(z) - 2)
        z = CFloat.Base1E6Subtract(z, d, UseENotationLongerThanDigits)
        d = z
    End If
    If Right$(d, 1) = "." Then d = Left$(d, Len(d) - 1)
    If d = "-0" Then d = "0"
    If d = "-0.0" Then d = "0"
    If Left$(d, 1) <> "-" Then d = " " & d
    Value_QuantizationError = d
End Property

Public Function Func_formatRemoveSciNotation$(n)
    Dim e&, m$, v
    
    Select Case True
        Case InStr(n, "E-")
            v = Split(CStr(n), "-")
            m = v(0)
            m = Left$(m, Len(m) - 1)
            m = Replace(m, ".", vbNullString)
            e = v(1)
            Func_formatRemoveSciNotation = "0." & String$(e - 1, "0") & m
            
        Case InStr(n, "E+")
            v = Split(CStr(n), "+")
            m = v(0)
            m = Left$(m, Len(m) - 1)
            m = Replace(m, ".", vbNullString)
            e = v(1)
            Func_formatRemoveSciNotation = m & String$(e - Len(m) + 1, "0")
            
        Case Else
            Func_formatRemoveSciNotation = n
    End Select
End Function


Public Function Func_x2Power#(x#, pow&)

    Func_x2Power = x * CFloat.PowerOf02Exact(pow)
    
End Function



Public Function Func_2PowerCurr#(x#)
    Dim n&
    
    If Me.Func_isNaN(x) Then Func_2PowerCurr = x: Exit Function
    n = Me.Func_floatExponent(x) - EXPONENT_BIAS
    If x = 0 Then Func_2PowerCurr = x: Exit Function
    
    If n >= Me.Spec_ExponentMax Then
        Func_2PowerCurr = Me.Sample_QNaN
    Else
        If n = EXPONENT_MIN Then n = CFloat.Floor(Log(Abs(x)) / Log(2))
        Func_2PowerCurr = CFloat.PowerOf02Double(n)
    End If
    
    If x < 0 Then Func_2PowerCurr = -Func_2PowerCurr
End Function

Public Function Func_2PowerNext#(x#, Optional subnormal_alt_representation As Boolean)
    Dim n&
    
    If Me.Func_isNaN(x) Then Func_2PowerNext = x: Exit Function
    If x = 0 Then Func_2PowerNext = x: Exit Function
    If Not subnormal_alt_representation Then
        n = Me.Func_floatExponent(x) - EXPONENT_BIAS + 1
    Else
        n = CStr(Log(Me.Func_2PowerCurr(Abs(x))) / Log(2)) + 1
    End If
    
    If n >= Me.Spec_ExponentMax Then
        Func_2PowerNext = Me.Sample_QNaN
        Exit Function
    End If
    
    Func_2PowerNext = CFloat.PowerOf02Double(n)
    If x < 0 Then Func_2PowerNext = -Func_2PowerNext
End Function


Public Function Func_2PowerPrev#(x#, Optional subnormal_alt_representation As Boolean)
    Dim n&
    
    If Me.Func_isNaN(x) Then Func_2PowerPrev = x: Exit Function
    If Me.Func_floatExponent(x) = Me.Spec_ExponentMaxBiased Then Func_2PowerPrev = x: Exit Function
    If x = 0 Then Func_2PowerPrev = x: Exit Function
    
    n = CFloat.Floor(Log(Abs(x)) / Log(2)) - 1
    
    If subnormal_alt_representation Then
        If n < EXPONENT_MIN_SUBNORMAL_EXPANSION Then n = EXPONENT_MIN_SUBNORMAL_EXPANSION
    Else
        If n <= Me.Spec_ExponentMin Then n = EXPONENT_MIN_SUBNORMAL_EXPANSION
    End If
    Func_2PowerPrev = CFloat.PowerOf02Double(n)
    If x < 0 Then Func_2PowerPrev = -Func_2PowerPrev
End Function



Public Function Func_binade#(x#, Optional return_exponent As Boolean, Optional subnormal_alt_representation As Boolean)
        
    If Me.Func_isNaN(x) Then Func_binade = x: Exit Function
    
    If Not Me.Func_isSubnormal(x) Then
        If Not return_exponent Then
            Func_binade = Me.Func_2PowerCurr(x)
        Else
            If x = 0 Then
                Func_binade = EXPONENT_MIN_SUBNORMAL_EXPANSION - 1
            Else
                Func_binade = Log(Me.Func_2PowerCurr(Abs(x))) / Log(2)
            End If
        End If
    Else
        If Not return_exponent Then
            If Not subnormal_alt_representation Then
                Func_binade = CFloat.PowerOf02Double(EXPONENT_MIN_SUBNORMAL_EXPANSION)
            Else
                Func_binade = Me.Func_2PowerCurr(x)
            End If
        Else
            If Not subnormal_alt_representation Then
                Func_binade = EXPONENT_MIN_SUBNORMAL_EXPANSION
            Else
                Func_binade = Log(Me.Func_2PowerCurr(Abs(x))) / Log(2)
            End If
        End If
    End If

End Function


Public Function Func_ulpDistance(a#, b#)

    Func_ulpDistance = CFloat.ULP_Distance(a, b, Me)
    
End Function


Public Property Get Value_ULP#()
    
    If Value_Sign < 0 Then
        Compare.Value = -Value
        Value_ULP = Compare.Value_ULP
        
    Else
    
        Select Case True
            Case Value_Is_Normal:       If Not Value = Sample_NormalMax Then Value_ULP = Value_NextUpDelta Else Value_ULP = Abs(Value_NextDownDelta)
            Case Value_Is_Subnormal:    Value_ULP = Value_NextUpDelta
            Case Value_Is_Infinity:     Value_ULP = Sample_Infinity
            Case Value_Is_NaN:          Value_ULP = Sample_QNaN
            Case Value_Is_Zero:         Value_ULP = Sample_SubnormalMin
        End Select
    
    End If
End Property

Public Property Get Value_ULP_Distance(target#)
    Value_ULP_Distance = CFloat.ULP_Distance(Value, target, Me)
End Property

Public Property Get Value_x2Power#(pow&)
    Value_x2Power = Me.Func_x2Power(Me.Value, pow)
End Property





Public Property Get Value_NextUp#(Optional steps&)
    Dim i&, bits&()
        
    If steps Then
        Compare.Value = Value
        If steps > 0 Then
            For i = 1 To steps
                Compare.Mutate_NextUp
            Next
        Else
            For i = 1 To steps
                Compare.Mutate_NextDown
            Next
        End If
        Value_NextUp = Compare.Value
        Exit Property
    End If
        
        
    If Value_Is_ZeroNegative Then Value_NextUp = 0: Exit Property
        
    ReDim bits(1 To STORAGE_BITS)

    If Value_Sign < 0 Then
        Compare.Value = Value
        Compare.Mutate_Negate
        CFloat.VariantDecimalIntegerToBitArrayFill Compare.Raw_Entire_Decimal - 1, bits
        bits(SIGN_BIT) = -(bits(SIGN_BIT) <> 1)
    Else
        CFloat.VariantDecimalIntegerToBitArrayFill Raw_Entire_Decimal + 1, bits
    End If
    Value_NextUp = Reinterpret_ByteArrayAsFloat(CFloat.BitArrayToByteArray(bits))
        
End Property



Public Function Mutate(Value) As CDouble
    On Error Resume Next
        Select Case True
        
            Case IsNumeric(Value)
                Me.Value = CDbl(Value)
                
            Case VarType(Value) = vbString
            
                Select Case True
                    Case InStr(1, Value, "0x", vbTextCompare)
                        If InStr(1, Value, "p", vbTextCompare) Then
                            Me.Value_HexFloatingPointLiteral = Value
                        Else
                            Me.Raw_Entire_Hexadecimal = Value
                        End If
                    Case Left$(LCase$(Value), 1) = "b"
                        Me.Raw_Entire_Binary = Value
                    Case Left$(LCase$(Value), 2) = "0b"
                        Me.Raw_Entire_Binary = Value
                End Select
                    
        End Select
    On Error GoTo 0
    Set Mutate = Me
End Function



Public Function Mutate_ToSample_Infinity(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_Infinity(negate)
    On Error GoTo 0
    Set Mutate_ToSample_Infinity = Me
End Function

Public Function Mutate_ToSample_MachineEpsilon(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_MachineEpsilon(negate)
    On Error GoTo 0
    Set Mutate_ToSample_MachineEpsilon = Me
End Function

Public Function Mutate_ToSample_MaxConsecutiveInteger(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_MaxConsecutiveInteger(negate)
    On Error GoTo 0
    Set Mutate_ToSample_MaxConsecutiveInteger = Me
End Function

Public Function Mutate_ToSample_MaxSafeInteger(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_MaxSafeInteger(negate)
    On Error GoTo 0
    Set Mutate_ToSample_MaxSafeInteger = Me
End Function

Public Function Mutate_ToSample_NormalMax(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_NormalMax(negate)
    On Error GoTo 0
    Set Mutate_ToSample_NormalMax = Me
End Function

Public Function Mutate_ToSample_NormalMin(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_NormalMin(negate)
    On Error GoTo 0
    Set Mutate_ToSample_NormalMin = Me
End Function

Public Function Mutate_ToSample_SubnormalMax(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_SubnormalMax(negate)
    On Error GoTo 0
    Set Mutate_ToSample_SubnormalMax = Me
End Function

Public Function Mutate_ToSample_SubnormalMin(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_SubnormalMin(negate)
    On Error GoTo 0
    Set Mutate_ToSample_SubnormalMin = Me
End Function

Public Function Mutate_ToSample_pi(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_pi(negate)
    On Error GoTo 0
    Set Mutate_ToSample_pi = Me
End Function

Public Function Mutate_ToSample_e(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_e(negate)
    On Error GoTo 0
    Set Mutate_ToSample_e = Me
End Function

Public Function Mutate_ToSample_ULPofOne(Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_ULPofOne(negate)
    On Error GoTo 0
    Set Mutate_ToSample_ULPofOne = Me
End Function

Public Function Mutate_ToSample_QNaN(Optional payload = 1, Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_QNaN(payload, negate)
    On Error GoTo 0
    Set Mutate_ToSample_QNaN = Me
End Function

Public Function Mutate_ToSample_SNaN(Optional payload = 1, Optional negate As Boolean) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_SNaN(payload, negate)
    On Error GoTo 0
    Set Mutate_ToSample_SNaN = Me
End Function

Public Function Mutate_ToSample_QNaNIndeterminate(Optional payload = 1) As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_QNaNIndeterminate(payload)
    On Error GoTo 0
    Set Mutate_ToSample_QNaNIndeterminate = Me
End Function

Public Function Mutate_ToSample_NegativeZero() As CDouble
    On Error Resume Next
      Me.Value = Me.Sample_NegativeZero()
    On Error GoTo 0
    Set Mutate_ToSample_NegativeZero = Me
End Function




Public Function Mutate_ToPrecisionComp(ByVal p&) As CDouble
    On Error Resume Next
      Me.Value = Me.Value_FindPrecisionComp(p)
    On Error GoTo 0
    Set Mutate_ToPrecisionComp = Me
End Function

Public Function Mutate_HexFloatingPointLiteral(hexstr$) As CDouble
    On Error Resume Next
      Value_HexFloatingPointLiteral = hexstr
    On Error GoTo 0
    Set Mutate_HexFloatingPointLiteral = Me
End Function

Public Function Mutate_Advance(by#) As CDouble
    On Error Resume Next
      Value = Value + by
    On Error GoTo 0
    Set Mutate_Advance = Me
End Function

Public Function Mutate_AdvanceULP(Optional steps = 1) As CDouble
    Dim newval, maxval
    On Error Resume Next
      If IsNumeric(steps) Then
        maxval = CDec(RAW_INTEGER_MAX_DECIMAL)
        newval = Me.Raw_Entire_Decimal + Int(steps)
        If newval < 0 Then newval = 0
        If newval > maxval Then newval = maxval
        Me.Raw_Entire_Decimal = newval
      End If
    On Error GoTo 0
    Set Mutate_AdvanceULP = Me
End Function


Public Property Get Value_Is_Identical(x#) As Boolean
    Me.Compare.Value = x
    Value_Is_Identical = Me.Raw_Entire_Decimal = Me.Compare.Raw_Entire_Decimal
End Property


Public Function Func_ieeeMean#(a#, b#)
    Dim x As Object
    Dim y As Object
    Dim z As Object
    
    Set x = Compare
    Set y = Compare.Compare
    Set z = Compare.Compare.Compare
    
    x.Value = a
    y.Value = b
    z.Value = 0
    z.Raw_Entire_Decimal = (x.Raw_Entire_Decimal + y.Raw_Entire_Decimal) / 2
    Func_ieeeMean = z.Value
End Function

Public Function Mutate_ieeeMean(x) As CDouble
    On Error Resume Next
      Me.Compare.Value = x
      Me.Raw_Entire_Decimal = (Me.Raw_Entire_Decimal + Me.Compare.Raw_Entire_Decimal) / 2
    On Error GoTo 0
    Set Mutate_ieeeMean = Me
End Function

Public Function Mutate_x2Power(pow&) As CDouble
    On Error Resume Next
      Me.Value = Me.Func_x2Power(Me.Value, pow)
    On Error GoTo 0
    Set Mutate_x2Power = Me
End Function



Public Function Mutate_AddProduct(a#, b#) As CDouble
    On Error Resume Next
      Value = Value + a * b
    On Error GoTo 0
    Set Mutate_AddProduct = Me
End Function

Public Function Mutate_AddSum(a#, b#) As CDouble
    On Error Resume Next
      Value = Value + a + b
    On Error GoTo 0
    Set Mutate_AddSum = Me
End Function

Public Function Mutate_AddQuotient(a#, b#) As CDouble
    On Error Resume Next
      Value = Value + a / b
    On Error GoTo 0
    Set Mutate_AddQuotient = Me
End Function

Public Function Mutate_AddDifference(a#, b#) As CDouble
    On Error Resume Next
      Value = Value + a - b
    On Error GoTo 0
    Set Mutate_AddDifference = Me
End Function



Public Function Mutate_Hypotenuse(a#, b#) As CDouble
    On Error Resume Next
      Value = Sqr((a ^ 2) + (b ^ 2))
    On Error GoTo 0
    Set Mutate_Hypotenuse = Me
End Function



Public Function Mutate_Binade(BinadeExponent&) As CDouble
    If BinadeExponent >= Spec_ExponentMinBiased And BinadeExponent <= Spec_ExponentMaxBiased Then
        Raw_Exponent_Biased_Decimal = CFloat.PowerOf02Double(BinadeExponent)
    End If
    Set Mutate_Binade = Me
End Function

Public Function Mutate_AdvanceBinade(Optional by& = 1) As CDouble
    Dim e&, n&
    e = Value_Exponent_Base10
    n = e + by
    If n > EXPONENT_MAX Then n = EXPONENT_MAX
    If n < EXPONENT_MIN Then n = EXPONENT_MIN
    Value_Exponent_Base10 = n
    Set Mutate_AdvanceBinade = Me
End Function






Public Function Mutate_Negate() As CDouble
    Raw_Bit(SIGN_BIT) = -(Raw_Bit(SIGN_BIT) <> 1)
    Synchronize
    Set Mutate_Negate = Me
End Function




Public Function Mutate_NextUp() As CDouble
    Value = Value_NextUp
    Set Mutate_NextUp = Me
End Function

Public Function Mutate_NextDown() As CDouble
    Value = Value_NextDown
    Set Mutate_NextDown = Me
End Function



Public Function Mutate_ToPower(Optional Power# = 2) As CDouble
    On Error Resume Next
      Value = CFloat.Power(Value, Power)
      If Err Then Value = Sample_QNaN
    On Error GoTo 0
    Set Mutate_ToPower = Me
End Function

Public Function Mutate_ToRoot(Optional Root# = 2) As CDouble
    On Error Resume Next
      Value = CFloat.Root(Value, Root)
      If Err Then Value = Sample_QNaN
    On Error GoTo 0
    Set Mutate_ToRoot = Me
End Function

Public Function Mutate_ToLogNatural() As CDouble
    If Value > 0 Then
        Value = Log(Value)
    Else
        Value = Sample_QNaN
    End If
    Set Mutate_ToLogNatural = Me
End Function

Public Function Mutate_ToLogBaseN(n#) As CDouble
    If Value > 0 And n > 0 Then
        Value = Log(Value) / Log(n)
    Else
        Value = Sample_QNaN
    End If
    Set Mutate_ToLogBaseN = Me
End Function

Public Function Mutate_ToCosine(Optional degrees As Boolean = False) As CDouble
    On Error Resume Next
      If degrees Then
        Value = Cos(CFloat.DegreesToRadians(Value))
      Else
        Value = Cos(Value)
      End If
    On Error GoTo 0
    Set Mutate_ToCosine = Me
End Function

Public Function Mutate_ToSine(Optional degrees As Boolean = False) As CDouble
    On Error Resume Next
      If degrees Then
        Value = Sin(CFloat.DegreesToRadians(Value))
      Else
        Value = Sin(Value)
      End If
    On Error GoTo 0
    Set Mutate_ToSine = Me
End Function

Public Function Mutate_ToTangent(Optional degrees As Boolean = False) As CDouble
    On Error Resume Next
      If degrees Then
        Value = Tan(CFloat.DegreesToRadians(Value))
      Else
        Value = Tan(Value)
      End If
    On Error GoTo 0
    Set Mutate_ToTangent = Me
End Function

Public Function Mutate_ToArctangent(Optional degrees As Boolean = False) As CDouble
    On Error Resume Next
      If degrees Then
        Value = Atn(CFloat.DegreesToRadians(Value))
      Else
        Value = Atn(Value)
      End If
    On Error GoTo 0
    Set Mutate_ToArctangent = Me
End Function

Public Function Mutate_ToRadians() As CDouble
    On Error Resume Next
      Value = CFloat.DegreesToRadians(Value)
    On Error GoTo 0
    Set Mutate_ToRadians = Me
End Function

Public Function Mutate_ToDegrees() As CDouble
    On Error Resume Next
      Value = CFloat.RadiansToDegrees(Value)
    On Error GoTo 0
    Set Mutate_ToDegrees = Me
End Function


Public Function Mutate_ToProduct(by#) As CDouble
    On Error Resume Next
      Value = Value * by
    On Error GoTo 0
    Set Mutate_ToProduct = Me
End Function

Public Function Mutate_ToQuotient(by#) As CDouble
    If by = 0 Then Value = Sample_Infinity: Exit Function
    Value = Value / by
    Set Mutate_ToQuotient = Me
End Function

Public Function Mutate_ToMod(by#) As CDouble
    If by = 0 Then Value = Sample_Infinity: Exit Function
    Value = Value Mod by
    Set Mutate_ToMod = Me
End Function

Public Function Mutate_ToExp() As CDouble
    On Error Resume Next
      Value = Exp(Value)
    On Error GoTo 0
    Set Mutate_ToExp = Me
End Function

Public Function Mutate_ToMax(n#) As CDouble
    On Error Resume Next
      If n > Value Then Value = n
    On Error GoTo 0
    Set Mutate_ToMax = Me
End Function

Public Function Mutate_ToMin(n#) As CDouble
    On Error Resume Next
      If n < Value Then Value = n
    On Error GoTo 0
    Set Mutate_ToMin = Me
End Function


Public Function Mutate_ToCeil() As CDouble
    On Error Resume Next
      Value = CFloat.Ceil(Value)
    On Error GoTo 0
    Set Mutate_ToCeil = Me
End Function


Public Function Mutate_ToFloor() As CDouble
    On Error Resume Next
      Value = CFloat.Floor(Value)
    On Error GoTo 0
    Set Mutate_ToFloor = Me
End Function


Public Function Mutate_Random(Optional bRandomize As Boolean = False) As CDouble
    If bRandomize Then Randomize
    Value = Rnd
    Set Mutate_Random = Me
End Function

Public Function Mutate_RandomIntegerBetween(lo#, hi#, Optional bRandomize As Boolean = False) As CDouble
    If bRandomize Then Randomize
    On Error Resume Next
      Value = CFloat.RandomIntegerBetween(lo, hi)
    On Error GoTo 0
    Set Mutate_RandomIntegerBetween = Me
End Function

Public Function Mutate_RandomInBinade(BinadeExponent&, Optional bIncludeNegatives As Boolean, Optional bIncludeZero As Boolean, Optional bRandomize As Boolean) As CDouble
    If BinadeExponent > Me.Spec_ExponentMax Then Me.Value = Me.Sample_Infinity: GoTo TheExit
    If BinadeExponent < EXPONENT_MIN_SUBNORMAL_EXPANSION Then Me.Value = 0:     GoTo TheExit
    Me.Value = 0
    
    If bRandomize Then Randomize
    Select Case True
        Case BinadeExponent > Spec_ExponentMin
            'Normal
            Me.Value_Exponent_Base10 = BinadeExponent
            Me.Raw_Mantissa_Decimal = CFloat.RandomIntegerBetween(0, CDec(RAW_MANTISSA_MAX_DECIMAL))
        Case Else
            'Subnormal
            Dim reduced_significand_size&
            Me.Value_Exponent_Base10 = Me.Spec_ExponentMin
            reduced_significand_size = -EXPONENT_MIN_SUBNORMAL_EXPANSION - -BinadeExponent
            Me.Raw_Mantissa_Decimal = CFloat.RandomIntegerBetween(-(Not bIncludeZero), CFloat.PowerOf02Double(reduced_significand_size) - 1)
            If Not bIncludeZero Then Me.Raw_Bit(reduced_significand_size + 1) = 1
    End Select
    
TheExit:
    If bIncludeNegatives Then If CFloat.RandomIntegerBetween(0, 1) Then Me.Mutate_Negate
    Set Mutate_RandomInBinade = Me
End Function

Public Function Mutate_RandomSubnormal(Optional bIncludeNegatives As Boolean, Optional bIncludeZero As Boolean, Optional bRandomize As Boolean) As CDouble
    Dim b&
    If bRandomize Then Randomize
    b = Abs(EXPONENT_MIN - EXPONENT_MIN_SUBNORMAL_EXPANSION)
    b = Abs(EXPONENT_MIN) + CFloat.RandomIntegerBetween(0, b)
    Mutate_RandomInBinade -b, bIncludeZero:=bIncludeZero
    If bIncludeNegatives Then If CFloat.RandomIntegerBetween(0, 1) Then Me.Mutate_Negate
    Set Mutate_RandomSubnormal = Me
End Function

Public Function Mutate_RandomNormal(Optional bIncludeInfNan As Boolean, Optional bIncludeNegatives As Boolean, Optional bRandomize As Boolean) As CDouble
    If bRandomize Then Randomize
    Me.Mutate_RandomInBinade CFloat.RandomIntegerBetween(Me.Spec_ExponentMin + 1, Me.Spec_ExponentMax + Not bIncludeInfNan)
    If bIncludeNegatives Then If CFloat.RandomIntegerBetween(0, 1) Then Me.Mutate_Negate
    Set Mutate_RandomNormal = Me
End Function

Public Function Mutate_RandomNormalSubnormal(Optional bIncludeNegatives As Boolean, Optional bIncludeZero As Boolean, Optional bIncludeInfNan As Boolean, Optional bRandomize As Boolean) As CDouble
    Dim b&
    If bRandomize Then Randomize
    b = CFloat.RandomIntegerBetween(Me.Spec_ExponentMin - Not bIncludeZero, Me.Spec_ExponentMax + Not bIncludeInfNan)
    If b > EXPONENT_MIN Then
        Me.Mutate_RandomInBinade b, bIncludeZero:=bIncludeZero
    Else
        Me.Mutate_RandomSubnormal bIncludeZero:=bIncludeZero
    End If
    If bIncludeNegatives Then If CFloat.RandomIntegerBetween(0, 1) Then Me.Mutate_Negate
    Set Mutate_RandomNormalSubnormal = Me
End Function





Public Property Get Value_DistanceTo#(target)
    On Error Resume Next
      Value_DistanceTo = target - Value
    On Error GoTo 0
End Property


Public Property Get Value_NextUpDelta#()
    Value_NextUpDelta# = Value_NextUp - Value
End Property

Public Property Get Value_NextDown#(Optional steps&)
    Dim i&, bits&()
        
    If steps Then
        Compare.Value = Value
        If steps > 0 Then
            For i = 1 To steps
                Compare.Mutate_NextDown
            Next
        Else
            For i = 1 To steps
                Compare.Mutate_NextUp
            Next
        End If
        Value_NextDown = Compare.Value
        Exit Property
    End If
    
    
    If Value_Is_ZeroPositive Then Value_NextDown = Sample_NegativeZero: Exit Property
    If Value_Is_ZeroNegative Then Value_NextDown = -Sample_SubnormalMin: Exit Property
    
    ReDim bits(1 To STORAGE_BITS)

    If Value_Sign < 0 Then
        Compare.Value = Value
        Compare.Mutate_Negate
        CFloat.VariantDecimalIntegerToBitArrayFill Compare.Raw_Entire_Decimal + 1, bits
        bits(SIGN_BIT) = -(bits(SIGN_BIT) <> 1)
    
    Else
        CFloat.VariantDecimalIntegerToBitArrayFill Raw_Entire_Decimal - 1, bits
    End If
    Value_NextDown = Reinterpret_ByteArrayAsFloat(CFloat.BitArrayToByteArray(bits))

End Property

Public Property Get Value_NextDownDelta#()
    Value_NextDownDelta = Value_NextDown - Value
End Property





Public Property Get Raw_Bit&(ByVal n&)
    If n >= 1 And n <= STORAGE_BITS Then
        Raw_Bit = theBitArray(n)
    Else
        Raw_Bit = -1
    End If
End Property
Public Property Let Raw_Bit(ByVal n&, ByVal Value&)
    If n >= 1 And n <= STORAGE_BITS Then
        theBitArray(n) = -(Value <> 0)
        Raw_BitsArray = theBitArray
    Else
        Stop
    End If
End Property



Public Property Get Raw_BitsArray() As Long()
    Raw_BitsArray = theBitArray
End Property
Public Property Let Raw_BitsArray(bits&())
    Raw_BytesArray = CFloat.BitArrayToByteArray(bits)
End Property



Public Property Get Raw_Byte&(ByVal n&)
    If n >= 0 And n <= (STORAGE_BYTES - 1) Then
        Raw_Byte = theByteArray.Elements(n)
    Else
        Raw_Byte = -1
    End If
End Property
Public Property Let Raw_Byte(ByVal n&, ByVal Value&)
    Dim b() As Byte
    
    If n >= 0 And n <= (STORAGE_BYTES - 1) Then
        theByteArray.Elements(n) = Value
        b = theByteArray.Elements
        Raw_BytesArray = b
    Else
        Stop
    End If
End Property



Public Property Get Raw_BytesArray() As Byte()
    Raw_BytesArray = theByteArray.Elements
End Property
Public Property Let Raw_BytesArray(b() As Byte)
    Static st As StringTemplate
    st.Value = b
    LSet theByteArray = st
    LSet theValue = theByteArray
    theBitArray = CFloat.ByteArrayToBitArray(b)
    cached = 0
End Property



Public Property Get Raw_Entire_Binary$(Optional ByVal FieldSpaces&, Optional ByVal ByteSpaces&, Optional ByVal SubnormalAlt&)
    Raw_Entire_Binary = CFloat.ByteArrayToBinaryString(theByteArray.Elements, FieldSpaces, ByteSpaces, SubnormalAlt)
End Property
Public Property Let Raw_Entire_Binary(Optional ByVal FieldSpaces&, Optional ByVal ByteSpaces&, Optional ByVal SubnormalAlt&, ByVal bin$)
    CFloat.BinaryStringToBitArrayFill Right$(CFloat.ForceStringToBinaryDigitsOnly(bin), STORAGE_BITS), theBitArray
    Synchronize
End Property


Public Property Get Raw_Entire_Binary_WithLegend$()
    Dim a$, b$, c$
    
    a = Space$(11) & Raw_Entire_Binary(4) & vbLf
    b = Space$(11) & "S" & Space$(4) & String$(EXPONENT_SIZE, "E") & Space$(4) & String$(MANTISSA_SIZE, "M") & vbLf
    c = "bits:" & Space$(11 - Len("bits:") - 1) & SIGN_BIT & Space$(3) & EXPONENT_BIT_HI & Space$(9) & EXPONENT_BIT_LO & Space$(2) & MANTISSA_BIT_HI & Space$(50) & MANTISSA_BIT_LO & vbLf

    Raw_Entire_Binary_WithLegend = vbLf & a & b & c
End Property





'        sign    exponent---    mantissa--------------------------------------------
'binary:    0    10000011001    0001001111010100001010000001001111010100000110000000
'           |    |         |    |                                                  |
'bits:     64    63       53    52                                                 1




'        sign
'           |
'           |  exponent    mantissa
'           ||----------||--------------------------------------------------------|
'           01000001 10010001 00111101 01000010 10000001 00111101 01000001 10000000
'           -------- -------- -------- -------- -------- -------- -------- --------
'                 41       91       3D       42       81       3D       41       80





'           1  *  2^(1024)     *  1.5000000000000002220446049250313080847263336181640625
'           |        |              |
'           |        |              |
'      (-1)^0       (2047-1023)     5000000000000002220446049250313080847263336181640625
'           |          |            |
'           |          |            |
'        sign       exponent...     mantissa............................................
'bits:      0       11111111111     1000000000000000000000000000000000000000000000000001
'           |       |         |     |                                                  |
'index:    64       63       53     52                                                 1





Public Property Get Raw_Entire_Decimal()
    Raw_Entire_Decimal = CFloat.BitArrayToVariantDecimalInteger(theBitArray)
End Property
Public Property Let Raw_Entire_Decimal(ByVal n)
    If Not IsNumeric(n) Then Exit Property
    n = Int(CDec(n))
    
    If n >= 0 And n <= CDec(RAW_INTEGER_MAX_DECIMAL) Then
        CFloat.BinaryStringToBitArrayFill CFloat.DecToBin(n), theBitArray
        Synchronize
    Else
        Stop
    End If
End Property



Public Property Get Raw_Entire_Hexadecimal$(Optional ByVal lowercase As Boolean = False, Optional ByVal prefix$)
    Raw_Entire_Hexadecimal = CFloat.ForceCase(CFloat.ByteArrayToHexString(theByteArray.Elements, prefix), --lowercase)
End Property
Public Property Let Raw_Entire_Hexadecimal(Optional ByVal lowercase As Boolean = False, Optional ByVal prefix$, ByVal h$)
    Raw_BytesArray = CFloat.HexStringToByteArray(Right$(CFloat.ForceStringToHexDigitsOnly(h), RAW_HEX_DIGITS))
End Property



Public Property Get Raw_Exponent_Biased_Binary$()
    Raw_Exponent_Biased_Binary = CFloat.BitArrayToBinaryString(theBitArray, EXPONENT_BIT_LO, EXPONENT_BIT_HI)
End Property
Public Property Let Raw_Exponent_Biased_Binary(ByVal bin$)
    
    bin = CFloat.ForceStringToBinaryDigitsOnly(bin)
    
    If Len(bin) > 1 And Len(bin) <= EXPONENT_SIZE Then
        CFloat.BinaryStringToBitArrayFill bin, theBitArray, EXPONENT_BIT_LO, EXPONENT_BIT_HI
        Synchronize
    Else
        Stop
    End If
End Property



Public Property Get Raw_Exponent_Biased_Decimal()
    Raw_Exponent_Biased_Decimal = CFloat.BitArrayToVariantDecimalInteger(theBitArray, EXPONENT_BIT_LO, EXPONENT_BIT_HI)
End Property
Public Property Let Raw_Exponent_Biased_Decimal(ByVal n)
    If Not IsNumeric(n) Then Exit Property
    If n >= 0 And n <= RAW_EXPONENT_MAX Then
        CFloat.VariantDecimalIntegerToBitArrayFill n, theBitArray, EXPONENT_BIT_LO, EXPONENT_BIT_HI
        Synchronize
    Else
        Err.Raise vbObjectError + 13, , vbLf & n & " is not a valid binary64 Biased exponent value." & vbLf & vbLf & "0 - " & RAW_EXPONENT_MAX & " is the Biased exponent range."
    End If
End Property



Public Property Get Raw_Exponent_Biased_Hexadecimal$(Optional ByVal lowercase As Boolean = False)
    Raw_Exponent_Biased_Hexadecimal = CFloat.ForceCase(CFloat.BitArrayToHexString(theBitArray, EXPONENT_BIT_LO, EXPONENT_BIT_HI), --lowercase)
End Property
Public Property Let Raw_Exponent_Biased_Hexadecimal(Optional ByVal lowercase As Boolean = False, ByVal h$)
    If Len(h) Then
        CFloat.HexStringToBitArrayFill Right$(CFloat.ForceStringToHexDigitsOnly(h), EXPONENT_HEX_DIGITS), theBitArray, EXPONENT_BIT_LO, EXPONENT_BIT_HI
    Else
        Stop
    End If
End Property






Public Property Get Raw_Mantissa_Binary$(Optional ByVal ByteSpaces&)
    Raw_Mantissa_Binary = Right$(Raw_Entire_Binary(, ByteSpaces), MANTISSA_SIZE + (MANTISSA_SIZE \ 8) * ByteSpaces)
End Property
Public Property Let Raw_Mantissa_Binary(Optional ByVal ByteSpaces&, ByVal bin$)
    bin = CFloat.ForceStringToBinaryDigitsOnly(bin)
    CFloat.BinaryStringToBitArrayFill Right$(bin, MANTISSA_SIZE), theBitArray, 1, MANTISSA_SIZE
    Synchronize
End Property



Public Property Get Raw_Mantissa_Decimal(Optional custom_hibit& = 0)
    If custom_hibit = 0 Then
        Raw_Mantissa_Decimal = CFloat.BitArrayToVariantDecimalInteger(theBitArray, 1, MANTISSA_SIZE)
    Else
        Raw_Mantissa_Decimal = CFloat.BitArrayToVariantDecimalInteger(theBitArray, 1, custom_hibit)
    End If
End Property
Public Property Let Raw_Mantissa_Decimal(Optional custom_hibit& = 0, ByVal n)
    If Not IsNumeric(n) Then Exit Property
    n = Int(CDec(n))
    
    If n >= 0 And n <= CDec(RAW_MANTISSA_MAX_DECIMAL) Then
        CFloat.VariantDecimalIntegerToBitArrayFill n, theBitArray, 1, MANTISSA_SIZE
        Synchronize
    Else
        Stop
    End If
End Property



Public Property Get Raw_Mantissa_Hexadecimal$(Optional ByVal lowercase As Boolean = False)
    Raw_Mantissa_Hexadecimal = CFloat.ForceCase(CFloat.TrimLeadingZeros(CFloat.ByteArrayToHexString(CFloat.BitArrayToByteArray(theBitArray, 1, MANTISSA_SIZE))), --lowercase)
    If Raw_Mantissa_Hexadecimal = vbNullString Then Raw_Mantissa_Hexadecimal = "0"
End Property
Public Property Let Raw_Mantissa_Hexadecimal(Optional ByVal lowercase As Boolean = False, ByVal h$)
    CFloat.HexStringToBitArrayFill Right$(CFloat.ForceStringToHexDigitsOnly(h), MANTISSA_HEX_DIGITS), theBitArray, 1, MANTISSA_SIZE
    Synchronize
End Property



Public Property Get Raw_SignBit&()
    Raw_SignBit = theBitArray(SIGN_BIT)
End Property
Public Property Let Raw_SignBit(n&)
    Raw_Bit(SIGN_BIT) = -(n <> 0)
End Property






'Default Property:
Public Property Get Value#(Optional ByVal bytecopy As Boolean = False)
Attribute Value.VB_UserMemId = 0
    If bytecopy Then
        CopyMemory ByVal VarPtr(Value), theByteArray.Elements(0), STORAGE_BYTES
    Else
        On Error Resume Next
        Value = theValue.Value
        On Error GoTo 0
    End If
End Property
Public Property Let Value(Optional ByVal bytecopy As Boolean, ByRef n#)
    Dim b() As Byte
    b = Reinterpret_FloatByRefAsByteArray(n)
    Raw_BytesArray = b
End Property


Public Property Get Value_Exact$(Optional ENotation As Boolean = True)
    Dim s$, e$, m$
    
    s = Value_Sign_Symbol(True)
    e = CFloat.PowerOf02Exact(Value_Exponent_Base10)
    m = Value_Mantissa_Base10(True)
    
    Value_Exact = s & CFloat.ExactProduct(e, m, ENotation)
End Property







Public Property Get Value_Exponent_Base02()
    Dim e10&, s$
    e10 = Value_Exponent_Base10
    If e10 < 0 Then s = "-"
    Value_Exponent_Base02 = s & CFloat.DecToBin(Abs(e10), Spec_ExponentBits)
End Property

Public Property Get Value_Exponent_Base10()
    Dim biased&
    biased = Raw_Exponent_Biased_Decimal
    If biased = 0 Then biased = 1
    Value_Exponent_Base10 = biased - EXPONENT_BIAS
End Property
Public Property Let Value_Exponent_Base10(ByVal n)
    If Not IsNumeric(n) Then Exit Property
    If n >= EXPONENT_MIN And n <= EXPONENT_MAX Then
        n = n + EXPONENT_BIAS
        CFloat.VariantDecimalIntegerToBitArrayFill n, theBitArray, EXPONENT_BIT_LO, EXPONENT_BIT_HI
        Synchronize
    Else
    
    End If
    
    Exit Property
    
End Property



Public Property Get Value_HexFloatingPointLiteral$(Optional ByVal lowercase As Boolean = True)
    Const DIGIT& = 1
    Static cache_$, lowercasemem As Boolean
    
    If Not cached And DIGIT Or lowercase <> lowercasemem Then
        cached = cached Or DIGIT
        cache_ = CFloat.FloatByRefToFloatingLiteral(theValue.Value, Me, lowercase)
        lowercasemem = lowercase
    End If
    Value_HexFloatingPointLiteral = cache_
End Property
Public Property Let Value_HexFloatingPointLiteral(Optional ByVal lowercase As Boolean = True, ByVal h$)
    
    If InStr(1, h, "0x", vbTextCompare) > 0 And InStr(1, h, "p", vbTextCompare) > 0 Then
        CFloat.FloatingLiteralToVariantByRef h, theValue.Value, Me
        Raw_BytesArray = Reinterpret_FloatByRefAsByteArray(theValue.Value)
    Else
        Stop
    End If

End Property



Public Property Get Value_ImplicitBit&()
    Value_ImplicitBit = -(0 <> Raw_Exponent_Biased_Decimal)
End Property
Public Property Let Value_ImplicitBit(n&)
    If n = 1 Or n = 0 Then
        Raw_Exponent_Biased_Decimal = n * RAW_EXPONENT_MAX
    Else
        Stop
    End If
End Property



Public Property Get Value_IntegerRatio$(Optional num$, Optional den$)
    Value_IntegerRatio = Func_floatIntegerRatio$(Me.Value, num, den)
End Property



Public Property Get Value_Mantissa_Base10$(Optional include_implicit_bit As Boolean = False, Optional custom_hibit& = -1)  'Read Only
    Value_Mantissa_Base10 = CFloat.BitArrayToMantissaBase10DecimalString(theBitArray, -include_implicit_bit * Value_ImplicitBit, custom_hibit)
End Property

Public Property Get Value_Mantissa_Base02$(Optional include_implicit_bit As Boolean = False)  'Read Only
    Dim intval$
    intval = "0"
    If include_implicit_bit Then intval = Value_ImplicitBit
    Value_Mantissa_Base02 = intval & "." & Raw_Mantissa_Binary
End Property



Public Property Get Value_NanPayload()
    If Not Value_Is_NaN Then Value_NanPayload = 0: Exit Property
    Value_NanPayload = Func_nanPayload(Value)
End Property




Public Property Get Value_Sign&()
    Value_Sign = (-1) ^ theBitArray(SIGN_BIT)
End Property
Public Property Let Value_Sign(n&)
    If n = 1 Or n = -1 Then
        Raw_Bit(SIGN_BIT) = CFloat.Ceil(n * -0.5)
    Else
        Stop
    End If
End Property



Public Property Get Sample_Infinity(Optional negate As Boolean) As CDouble
    Set Sample_Infinity = New CDouble
    Sample_Infinity = CFloat.HexStringToVariantFloat(HEX_INFINITY__POSITIVE, Me)
    If negate Then Sample_Infinity = -Sample_Infinity
End Property

Public Property Get Sample_NegativeZero() As CDouble
    Set Sample_NegativeZero = New CDouble
    Sample_NegativeZero = CFloat.HexStringToVariantFloat(HEX_ZERO______NEGATIVE, Me)
End Property



Public Property Get Sample_QNaN(Optional payload = 1, Optional negate As Boolean) As CDouble
    Dim n#
    Set Sample_QNaN = New CDouble
    n = CFloat.HexStringToVariantFloat(HEX_QNAN______POSITIVE, Me)
    Sample_QNaN = CFloat.FloatNanPayloadImplant(n, payload, Me, negate)
    
End Property

Public Property Get Sample_QNaNIndeterminate(Optional payload = 1) As CDouble
    Dim n#
    Set Sample_QNaNIndeterminate = New CDouble
    n = CFloat.HexStringToVariantFloat(HEX_QNAN_INDETERMINATE, Me)
    Sample_QNaNIndeterminate = CFloat.FloatNanPayloadImplant(n, payload, Me)
    
End Property

Public Property Get Sample_SNaN(Optional payload = 1, Optional negate As Boolean) As CDouble
    Dim n#
    Set Sample_SNaN = New CDouble
    n = CFloat.HexStringToVariantFloat(HEX_SNAN______POSITIVE, Me)
    Sample_SNaN = CFloat.FloatNanPayloadImplant(n, payload, Me, negate)
    
End Property


Public Property Get Sample_MaxConsecutiveInteger(Optional negate As Boolean) As CDouble
    Set Sample_MaxConsecutiveInteger = New CDouble
    Sample_MaxConsecutiveInteger = CFloat.HexStringToVariantFloat(HEX_MAX_CONSEC_INT_POS, Me)
    If negate Then Sample_MaxConsecutiveInteger = -Sample_MaxConsecutiveInteger
End Property

Public Property Get Sample_MaxSafeInteger(Optional negate As Boolean) As CDouble
    Set Sample_MaxSafeInteger = New CDouble
    Sample_MaxSafeInteger = CFloat.HexStringToVariantFloat(HEX_MAX_SAFE___INT_POS, Me)
    If negate Then Sample_MaxSafeInteger = -Sample_MaxSafeInteger
End Property


Public Property Get Sample_NormalMax(Optional negate As Boolean) As CDouble
    Set Sample_NormalMax = New CDouble
    Sample_NormalMax = CFloat.HexStringToVariantFloat(HEX_NORM_MAX__POSITIVE, Me)
    If negate Then Sample_NormalMax = -Sample_NormalMax
End Property

Public Property Get Sample_NormalMin(Optional negate As Boolean) As CDouble
    Set Sample_NormalMin = New CDouble
    Sample_NormalMin = CFloat.HexStringToVariantFloat(HEX_NORM_MIN__POSITIVE, Me)
    If negate Then Sample_NormalMin = -Sample_NormalMin
End Property

Public Property Get Sample_SubnormalMax(Optional negate As Boolean) As CDouble
    Set Sample_SubnormalMax = New CDouble
    Sample_SubnormalMax = CFloat.HexStringToVariantFloat(HEX_SUBNORMMAXPOSITIVE, Me)
    If negate Then Sample_SubnormalMax = -Sample_SubnormalMax
End Property

Public Property Get Sample_SubnormalMin(Optional negate As Boolean) As CDouble
    Set Sample_SubnormalMin = New CDouble
    Sample_SubnormalMin = CFloat.HexStringToVariantFloat(HEX_SUBNORMMINPOSITIVE$, Me)
    If negate Then Sample_SubnormalMin = -Sample_SubnormalMin
End Property

Public Property Get Sample_pi(Optional negate As Boolean) As CDouble
    Set Sample_pi = New CDouble
    Sample_pi = CFloat.HexStringToVariantFloat(HEX_PI________POSITIVE, Me)
    If negate Then Sample_pi = -Sample_pi
End Property

Public Property Get Sample_e(Optional negate As Boolean) As CDouble
    Set Sample_e = New CDouble
    Sample_e = CFloat.HexStringToVariantFloat(HEX_E_________POSITIVE$, Me)
    If negate Then Sample_e = -Sample_e
End Property

Public Property Get Sample_ULPofOne(Optional negate As Boolean) As CDouble
    Set Sample_ULPofOne = New CDouble
    Sample_ULPofOne = Func_ulp(1)
    If negate Then Sample_ULPofOne = -Sample_ULPofOne
End Property

Public Property Get Sample_MachineEpsilon(Optional negate As Boolean) As CDouble
    Set Sample_MachineEpsilon = New CDouble
    Sample_MachineEpsilon = Sample_ULPofOne
    If negate Then Sample_MachineEpsilon = -Sample_MachineEpsilon
End Property







'===================================================================================
'===================================================================================
'===================================================================================
'===================================================================================



Public Property Get Sample_SpecializedNaN#(Optional bNegative As Boolean, Optional exponent&, Optional MantissaHex$)
    Dim hibyte&, m$, se$, b() As Byte
    
    hibyte = Spec_BytesStorage
    ReDim b(0 To hibyte)
    b(hibyte) = &H7F: If bNegative Then b(hibyte) = &HFF
    If exponent > Spec_ExponentMaxBiased Then
        b(hibyte - 1) = &HF0
    ElseIf Len(MantissaHex) = 0 Then
        b(hibyte - 1) = &HF8
    Else
        m = String$(Spec_SignificandHexDigitCount, "0")
        Mid$(m, 1, Len(Left$(MantissaHex, Spec_SignificandHexDigitCount))) = MantissaHex
        se = "7FF"
        If b(hibyte) = &HFF Then se = "FFF"
        Sample_SpecializedNaN = CFloat.HexStringToVariantFloat(se & m, Me)
        Exit Property
    End If
    Sample_SpecializedNaN = Reinterpret_ByteArrayAsFloat(b)
End Property



Public Function Func_precisionComp&(a#, b#)

   Func_precisionComp = CFloat.PrecisionComp(a, b, Me)

End Function


Public Function Reinterpret_FloatByRefAsByteArray(ByRef n As Double) As Byte()
    Static bt As ByteArrayTemplate
    Static ft As FloatTemplate
    ft.Value = n
    LSet bt = ft
    Reinterpret_FloatByRefAsByteArray = bt.Elements
End Function


Public Function Reinterpret_ByteArrayAsFloat#(b() As Byte)
    Static ft As FloatTemplate
    Static st As StringTemplate
    st.Value = b
    LSet ft = st
    On Error Resume Next
    Reinterpret_ByteArrayAsFloat = ft.Value
    On Error GoTo 0
End Function


Private Sub Synchronize()
    Raw_BytesArray = CFloat.BitArrayToByteArray(theBitArray)
End Sub


Private Sub Class_Initialize()
    Value = 0   'Sample_QNaN
End Sub


