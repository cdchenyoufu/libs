VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UByte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'this class is so you can do byte operations which would overflow 255 and go back to 0 like in C

Enum op
    op_add = 0
    op_sub = 1
    op_div = 2
    op_mul = 3
    op_mod = 4
    op_xor = 5
    op_and = 6
    op_or = 7
    op_rsh = 8
    op_lsh = 9
    op_gt = 10
    op_lt = 11
    op_gteq = 12
    op_lteq = 13
    op_setBit = 15
    op_clearBit = 16
    op_toggleBit = 17
    op_testBit = 18
End Enum

Enum modes
    mUnsigned = 0
    mSigned = 1
    mHex = 2
End Enum

Public use0x As Boolean
Public Value As Byte          'this is the default method of the class
Attribute Value.VB_VarUserMemId = 0
Public padLeft As Boolean

'unsigned math operations
Private Declare Function ubop Lib "utypes.dll" Alias "UByte" (ByVal v1 As Byte, ByVal v2 As Byte, ByVal operation As Long) As Byte

Dim hLib As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'this way you can .setval(x).add(y)
Function setVal(v) As UByte
    Set setVal = Me
    Value = v
End Function

'as a property, the IDE can show it results as a tooltip when debugging
Function toString(Optional asHex As Boolean = True) As String
    Dim tmp As String
    
    If Not asHex Then
        toString = m_Value
    Else
        tmp = Hex(m_Value)
         
        If padLeft Then
            tmp = Right("00" & tmp, 2)
        End If
             
        If use0x Then tmp = "0x" & tmp
        
        toString = tmp
    End If
    
End Function

Function fromString(ByVal s As String, Optional forceHex As Boolean = False) As Boolean
    On Error GoTo hell
    
    s = trim(s)
    
    If LCase(Left(s, 2)) = "0x" Or LCase(Left(s, 2)) = "&h" Or LCase(Left(s, 2)) = "\x" Then
        forceHex = True
        s = Mid(s, 3)
    End If
    
    If forceHex Then
        m_Value = CByte(CInt("&h" & s))
    Else
        m_Value = CByte(s)
    End If
    
    fromString = True
    Exit Function
hell:
End Function

'for signed math..you can use native vb6 operations..so mode is omitted..
Function raw_op(operand2, operation As op) As UByte
    Dim ret As New UByte
    Dim v2 As Byte
    
    If TypeName(operand2) = "UByte" Then
        v2 = operand2.Value
    Else
        v2 = CByte(operand2)
    End If
    
    ret.padLeft = Me.padLeft
    ret.use0x = Me.use0x
    ret.Value = ubop(Value, v2, operation)
    
    Set raw_op = ret
    
End Function

Private Sub Class_Initialize()
'    hLib = LoadLibrary("UTypes.dll")
'    If hLib = 0 Then hLib = LoadLibrary(App.path & "\UTypes.dll")
'    If hLib = 0 Then hLib = LoadLibrary(App.path & "\..\UTypes.dll")
'    If hLib = 0 Then hLib = LoadLibrary(App.path & "\..\..\UTypes.dll")
    If ensureUTypes() Then hLib = modGeneral.hUTypes
    padLeft = True
End Sub


'these stubs are for readability/ease of use
Function add(operand2) As UByte
    Set add = raw_op(operand2, op_add)
End Function

Function subtract(operand2) As UByte
    Set subtract = raw_op(operand2, op_sub)
End Function

Function divide(operand2) As UByte
    Set divide = raw_op(operand2, op_div)
End Function

Function multiply(operand2) As UByte
    Set multiply = raw_op(operand2, op_mul)
End Function

Function modulus(operand2) As UByte
    Set modulus = raw_op(operand2, op_mod)
End Function

Function xor_(operand2) As UByte
    Set xor_ = raw_op(operand2, op_xor)
End Function

Function and_(operand2) As UByte
    Set and_ = raw_op(operand2, op_and)
End Function

Function or_(operand2) As UByte
    Set or_ = raw_op(operand2, op_or)
End Function

Function rshift(operand2) As UByte
    Set rshift = raw_op(operand2, op_rsh)
End Function

Function lshift(operand2) As UByte
    Set lshift = raw_op(operand2, op_lsh)
End Function


Function greaterThan(operand2) As Boolean
    Dim a As ULong
    Set a = raw_op(operand2, op_gt)
    greaterThan = (a.Value = 1)
End Function

Function lessThan(operand2) As Boolean
    Dim a As ULong
    Set a = raw_op(operand2, op_lt)
    lessThan = (a.Value = 1)
End Function

Function greaterThanEqual(operand2) As Boolean
    Dim a As ULong
    Set a = raw_op(operand2, op_gteq)
    greaterThanEqual = (a.Value = 1)
End Function

Function lessThanEqual(operand2) As Boolean
    Dim a As ULong
    Set a = raw_op(operand2, op_lteq)
    lessThanEqual = (a.Value = 1)
End Function

''bit operations
Function setBit(operand2) As ULong64
    Set setBit = raw_op(operand2, op_setBit)
End Function

Function clearBit(operand2) As ULong64
    Set clearBit = raw_op(operand2, op_clearBit)
End Function

Function toggleBit(operand2) As ULong64
    Set toggleBit = raw_op(operand2, op_toggleBit)
End Function

Function testBit(operand2) As ULong64
    Set testBit = raw_op(operand2, op_testBit)
End Function
