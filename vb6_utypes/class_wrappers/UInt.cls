VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Enum op
'    op_add = 0
'    op_sub = 1
'    op_div = 2
'    op_mul = 3
'    op_mod = 4
'    op_xor = 5
'    op_and = 6
'    op_or = 7
'    op_rsh = 8
'    op_lsh = 9
'    op_gt = 10
'    op_lt = 11
'    op_gteq = 12
'    op_lteq = 13
'    op_rol32 = 14
'End Enum
'
'Enum modes
'    mUnsigned = 0
'    mSigned = 1
'    mHex = 2
'End Enum

Public use0x As Boolean
Public rawInt As Integer
Public padLeft As Boolean

'unsigned math operations
Private Declare Function uiop Lib "utypes.dll" Alias "UInt" (ByVal v1 As Integer, ByVal v2 As Integer, ByVal operation As Long) As Integer

'convert an unsigned long (or int) to unsigned string (vb6 hex and signed displays are fine so ommited..)
Private Declare Function UShortStr Lib "utypes.dll" (ByVal v1 As Integer, ByVal buf As String, ByVal cBufferSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'create a 32 bit number from a string in specified base (16 for a hex string)
Private Declare Function Str2Long Lib "utypes.dll" (ByVal s As String, Optional ByVal base As Long = 10) As Long

Dim hLib As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Property Get Value()
    Value = rawInt
End Property

'this is the default method of the class
Property Let Value(v)
    Dim tmp As Integer

    If TypeName(v) = "String" Then
        fromString v
    ElseIf TypeName(v) = "UInt" Then
        rawInt = v.Value
    ElseIf TypeName(v) = "Integer" Or TypeName(v) = "Long" Or TypeName(v) = "Currency" Then
        CopyMemory tmp, VarPtr(v), 2
        rawInt = tmp
    ElseIf TypeName(v) = "Byte" Then
        rawInt = v
    Else
        Err.Raise 1, "Cint.Value", "Unsupported conversion type toint: " & TypeName(v)
    End If

End Property

'this way you can .setval(x).add(y)
Function setVal(v) As UInt
    Set setVal = Me
    Value = v
End Function

Property Get MAX_SIGNED() As String
     MAX_SIGNED = "7FFF"
     If use0x Then MAX_SIGNED = "0x" & MAX_SIGNED
End Property

Property Get MIN_SIGNED() As String
    MIN_SIGNED = "8000"
    If use0x Then MIN_SIGNED = "0x" & MIN_SIGNED
End Property

Property Get MAX_UNSIGNED() As String
    MAX_UNSIGNED = "FFFF"
    If use0x Then MAX_UNSIGNED = "0x" & MAX_UNSIGNED
End Property

Property Get isNegBitSet() As Boolean
    isNegBitSet = (Left(CStr(rawInt), 1) = "-")
End Property

''as a property, the IDE can show it results as a tooltip when debugging
Function toString(Optional m As modes = mHex) As String
    Dim tmp As String

    If m = mUnsigned Then
        toString = GetUnsigned(rawInt)
    ElseIf m = mSigned Then
        toString = rawInt
    Else
        tmp = Hex(rawInt)

        If padLeft Then
            tmp = Right("0000" & tmp, 4)
        End If

        If use0x Then tmp = "0x" & tmp

        toString = tmp
    End If

End Function

Function fromString(ByVal s As String, Optional forceHex As Boolean = True) As Boolean
    Dim tmp As Long, tint As Integer

    On Error Resume Next
    s = Trim(s)

    If Left(s, 1) = "-" Then forceHex = False

    If LCase(Left(s, 2)) = "0x" Or LCase(Left(s, 2)) = "&h" Or LCase(Left(s, 2)) = "\x" Then
        forceHex = True
        s = Mid(s, 3)
    End If

    tmp = Str2Long(s, IIf(forceHex, 16, 10))
    CopyMemory tint, VarPtr(tmp), 2
    rawInt = tint

    fromString = True
    Exit Function
hell:
End Function

Private Function GetUnsigned(v As Integer) As String
    Dim tmp As String, i As Long
    tmp = Space(64)
    i = UShortStr(v, tmp, 64)
    If i > 0 Then GetUnsigned = Mid(tmp, 1, i)
End Function

'for signed math..you can use native vb6 operations..so mode is omitted..
Function raw_op(operand2, operation As op) As UInt
    Dim ret As New UInt
    Dim v2 As Long
    'Dim tmp As Integer

    If TypeName(operand2) = "UInt" Then
        v2 = operand2.Value
    Else
        v2 = CInt(operand2)
    End If

    'tmp = rawInt
    ret.use0x = Me.use0x
    ret.padLeft = Me.padLeft
    ret.Value = uiop(rawInt, v2, operation)

    Set raw_op = ret
'
End Function

Private Sub Class_Initialize()
    If ensureUTypes() Then hLib = modGeneral.hUtypes
    padLeft = True
End Sub

'these stubs are for readability/ease of use
Function add(operand2) As UInt
    Set add = raw_op(operand2, op_add)
End Function

Function subtract(operand2) As UInt
    Set subtract = raw_op(operand2, op_sub)
End Function

Function divide(operand2) As UInt
    Set divide = raw_op(operand2, op_div)
End Function

Function multiply(operand2) As UInt
    Set multiply = raw_op(operand2, op_mul)
End Function

Function modulus(operand2) As UInt
    Set modulus = raw_op(operand2, op_mod)
End Function

Function xor_(operand2) As UInt
    Set xor_ = raw_op(operand2, op_xor)
End Function

Function and_(operand2) As UInt
    Set and_ = raw_op(operand2, op_and)
End Function

Function or_(operand2) As UInt
    Set or_ = raw_op(operand2, op_or)
End Function

Function rshift(operand2) As UInt
    Set rshift = raw_op(operand2, op_rsh)
End Function

Function lshift(operand2) As UInt
    Set lshift = raw_op(operand2, op_lsh)
End Function

Function greaterThan(operand2) As Boolean
    Dim a As UInt
    Set a = raw_op(operand2, op_gt)
    greaterThan = (a.Value = 1)
End Function

Function lessThan(operand2) As Boolean
    Dim a As UInt
    Set a = raw_op(operand2, op_lt)
    lessThan = (a.Value = 1)
End Function

Function greaterThanEqual(operand2) As Boolean
    Dim a As UInt
    Set a = raw_op(operand2, op_gteq)
    greaterThanEqual = (a.Value = 1)
End Function

Function lessThanEqual(operand2) As Boolean
    Dim a As UInt
    Set a = raw_op(operand2, op_lteq)
    lessThanEqual = (a.Value = 1)
End Function

Function rol32(operand2) As UInt
    Set rol32 = raw_op(operand2, op_rol32)
End Function

''bit operations
'Function setBit(operand2) As UInt
'    Set setBit = raw_op(operand2, op_setBit)
'End Function
'
'Function clearBit(operand2) As UInt
'    Set clearBit = raw_op(operand2, op_clearBit)
'End Function
'
'Function toggleBit(operand2) As UInt
'    Set toggleBit = raw_op(operand2, op_toggleBit)
'End Function
'
'Function testBit(operand2) As UInt
'    Set testBit = raw_op(operand2, op_testBit)
'End Function

