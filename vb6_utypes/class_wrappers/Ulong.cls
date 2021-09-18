VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ULong"
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
Public rawValue As Long           'this is the default method of the class
Public padLeft As Boolean

'unsigned math operations
Private Declare Function ULong Lib "utypes.dll" (ByVal v1 As Long, ByVal v2 As Long, ByVal operation As Long) As Long

'convert an unsigned long (or int) to unsigned string (vb6 hex and signed displays are fine so ommited..)
Private Declare Function U2Str Lib "utypes.dll" (ByVal v1 As Long, ByVal buf As String, ByVal cBufferSize As Long) As Long

'create a 32 bit number from a string in specified base (16 for a hex string)
Private Declare Function Str2Long Lib "utypes.dll" (ByVal s As String, Optional ByVal base As Long = 10) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Dim hLib As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Property Get Value()
Attribute Value.VB_UserMemId = 0
    Value = rawValue
End Property

'this is the default method of the class
Property Let Value(v)
    
    If TypeName(v) = "String" Then
        fromString v
    ElseIf TypeName(v) = "ULong" Or TypeName(v) = "UInt" Or TypeName(v) = "UByte" Then
        rawValue = v.Value
    ElseIf TypeName(v) = "Long" Or TypeName(v) = "Byte" Or TypeName(v) = "Integer" Then
        rawValue = v
    Else
        Err.Raise 1, "ULong.Value", "Unsupported conversion type to long: " & TypeName(v)
    End If

End Property

'this way you can .setval(x).add(y)
Function setVal(v) As ULong
    Set setVal = Me
    Value = v
End Function

Property Get MAX_SIGNED() As String
     MAX_SIGNED = "7FFFFFFF"
     If use0x Then MAX_SIGNED = "0x" & MAX_SIGNED
End Property

Property Get MIN_SIGNED() As String
    MIN_SIGNED = "80000000"
    If use0x Then MIN_SIGNED = "0x" & MIN_SIGNED
End Property

Property Get MAX_UNSIGNED() As String
    MAX_UNSIGNED = "FFFFFFFF"
    If use0x Then MAX_UNSIGNED = "0x" & MAX_UNSIGNED
End Property

Property Get isNegBitSet() As Boolean
    isNegBitSet = (Left(CStr(rawValue), 1) = "-")
End Property

'as a property, the IDE can show it results as a tooltip when debugging
Function toString(Optional m As modes = mHex) As String
    Dim tmp As String
    
    If m = mUnsigned Then
        toString = GetUnsigned(rawValue)
    ElseIf m = mSigned Then
        toString = rawValue
    Else
        tmp = Hex(rawValue)
         
        If padLeft Then
            tmp = Right("00000000" & tmp, 8)
        End If
             
        If use0x Then tmp = "0x" & tmp
        
        toString = tmp
    End If
    
End Function

Function fromString(ByVal s As String, Optional forceHex As Boolean = True) As Boolean
    On Error GoTo hell
    Dim tmp As Long
    
    s = Trim(s)
    If Left(s, 1) = "-" Then forceHex = False
    
    If LCase(Left(s, 2)) = "0x" Or LCase(Left(s, 2)) = "&h" Or LCase(Left(s, 2)) = "\x" Then
        forceHex = True
        s = Mid(s, 3)
    End If
     
    rawValue = Str2Long(s, IIf(forceHex, 16, 10))
    fromString = True
    Exit Function
hell:
End Function

Private Function GetUnsigned(v As Long) As String
    Dim tmp As String, i As Long
    tmp = Space(64)
    i = U2Str(v, tmp, 64)
    If i > 0 Then GetUnsigned = Mid(tmp, 1, i)
End Function

'for signed math..you can use native vb6 operations..so mode is omitted..
Function raw_op(operand2, operation As op) As ULong
    Dim ret As New ULong
    Dim v2 As Long
    
    If TypeName(operand2) = "ULong" Then
        v2 = operand2.Value
    Else
        v2 = CLng(operand2)
    End If
    
    ret.use0x = Me.use0x
    ret.padLeft = Me.padLeft
    ret.rawValue = ULong(rawValue, v2, operation)
    
    Set raw_op = ret
    
End Function

Private Sub Class_Initialize()
    If ensureUTypes() Then hLib = modGeneral.hUtypes
    padLeft = True
End Sub

'these stubs are for readability/ease of use
Function add(operand2) As ULong
    Set add = raw_op(operand2, op_add)
End Function

Function subtract(operand2) As ULong
    Set subtract = raw_op(operand2, op_sub)
End Function

Function divide(operand2) As ULong
    Set divide = raw_op(operand2, op_div)
End Function

Function multiply(operand2) As ULong
    Set multiply = raw_op(operand2, op_mul)
End Function

Function modulus(operand2) As ULong
    Set modulus = raw_op(operand2, op_mod)
End Function

Function xor_(operand2) As ULong
    Set xor_ = raw_op(operand2, op_xor)
End Function

Function and_(operand2) As ULong
    Set and_ = raw_op(operand2, op_and)
End Function

Function or_(operand2) As ULong
    Set or_ = raw_op(operand2, op_or)
End Function

Function rshift(operand2) As ULong
    Set rshift = raw_op(operand2, op_rsh)
End Function

Function lshift(operand2) As ULong
    Set lshift = raw_op(operand2, op_lsh)
End Function

Function greaterThan(operand2) As Boolean
    Dim a As ULong
    Set a = raw_op(operand2, op_gt)
    greaterThan = (a.rawValue = 1)
End Function

Function lessThan(operand2) As Boolean
    Dim a As ULong
    Set a = raw_op(operand2, op_lt)
    lessThan = (a.rawValue = 1)
End Function

Function greaterThanEqual(operand2) As Boolean
    Dim a As ULong
    Set a = raw_op(operand2, op_gteq)
    greaterThanEqual = (a.rawValue = 1)
End Function

Function lessThanEqual(operand2) As Boolean
    Dim a As ULong
    Set a = raw_op(operand2, op_lteq)
    lessThanEqual = (a.rawValue = 1)
End Function

Function rol32(operand2) As ULong
    Set rol32 = raw_op(operand2, op_rol32)
End Function

''bit operations
'Function setBit(operand2) As ULong
'    Set setBit = raw_op(operand2, op_setBit)
'End Function
'
'Function clearBit(operand2) As ULong
'    Set clearBit = raw_op(operand2, op_clearBit)
'End Function
'
'Function toggleBit(operand2) As ULong
'    Set toggleBit = raw_op(operand2, op_toggleBit)
'End Function
'
'Function testBit(operand2) As ULong
'    Set testBit = raw_op(operand2, op_testBit)
'End Function

