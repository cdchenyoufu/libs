VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBitTest 
      Caption         =   "Bit test"
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   3420
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   300
      TabIndex        =   3
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "vc rand"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Native Tests"
      Height          =   330
      Left            =   3195
      TabIndex        =   1
      Top             =   3375
      Width           =   1140
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    op_rol32 = 14
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


Private hLib As Long 'for IDE release
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'unsigned math operations
Private Declare Function ULong Lib "utypes.dll" (ByVal v1 As Long, ByVal v2 As Long, ByVal operation As Long) As Long
Private Declare Function UInt Lib "utypes.dll" (ByVal v1 As Integer, ByVal v2 As Integer, ByVal operation As Long) As Integer
Private Declare Function U64 Lib "utypes.dll" (ByVal v1 As Currency, ByVal v2 As Currency, ByVal operation As op) As Currency
Private Declare Function UByte Lib "utypes.dll" (ByVal v1 As Byte, ByVal v2 As Byte, ByVal operation As op) As Byte

'signed math for 64 bit numbers (necessary?)
Private Declare Function S64 Lib "utypes.dll" (ByVal v1 As Currency, ByVal v2 As Currency, ByVal operation As op) As Currency

'create 64 bit number from hi and lo longs
Private Declare Function toU64 Lib "utypes.dll" (ByVal v1 As Long, ByVal v2 As Long) As Currency

'create a 64 bit number from a string in specified base (16 for a hex string)
Private Declare Function Str264 Lib "utypes.dll" (ByVal s As String, Optional ByVal base As Long = 10) As Currency

'convert a 64 bit number to string in specified format
Private Declare Function U642Str Lib "utypes.dll" (ByVal v1 As Currency, ByVal buf As String, ByVal cBufferSize As Long, ByVal mode As modes) As Long

'get hi and lo longs from 64 numbers
Private Declare Sub fromU64 Lib "utypes.dll" (ByVal v0 As Currency, ByRef v1 As Long, ByRef v2 As Long)

'convert an unsigned long (or int) to unsigned string (vb6 hex and signed displays are fine so ommited..)
Private Declare Function U2Str Lib "utypes.dll" (ByVal v1 As Long, ByVal buf As String, ByVal cBufferSize As Long) As Long

Private Declare Sub srand Lib "utypes.dll" Alias "vc_srand" (ByVal v1 As Long)
Private Declare Function rand Lib "utypes.dll" Alias "vc_rand" () As Long

Private Declare Function crc64s Lib "utypes.dll" (ByVal wStrPtr As Long, Optional asciiOnly As Long = 1) As Currency
Private Declare Function crc64 Lib "utypes.dll" (ByRef stream As Byte, ByVal sz As Long) As Currency

Private Declare Function entropy Lib "utypes.dll" (ByRef stream As Byte, ByVal sz As Long) As Double

Sub showBits(v As Variant)
    
    Dim max As Long, i As Long
    Dim tmp As String
    
    If TypeName(v) = "Byte" Then max = 7
    If TypeName(v) = "Integer" Then max = 15
    If TypeName(v) = "Long" Then max = 32
    If TypeName(v) = "Currency" Then max = 63
    
    If max = 0 Then
        List1.AddItem "showBits Unsupported type " & TypeName(v)
        Exit Sub
    End If
    
    tmp = "Bits set in " & TypeName(v) & " " & Hex(v) & ": "
    
    For i = 0 To max
        If max = 7 Then tmp = tmp & UByte(CByte(v), i, op_testBit) & " "
        If max = 15 Then tmp = tmp & UInt(CInt(v), i, op_testBit) & " "
        If max = 32 Then tmp = tmp & ULong(CLng(v), i, op_testBit) & " "
        If max = 63 Then tmp = tmp & ULong(CCur(v), i, op_testBit) & " "
    Next
    
    List1.AddItem tmp
    
End Sub

Private Sub cmdBitTest_Click()
    
    Dim b As Byte, b2 As Byte, tmp As String, i As Long
    List1.Clear
    
    b = 0
    For i = 0 To 7
        b = UByte(b, i, op_setBit)
        tmp = tmp & b & " "
    Next
    
    List1.AddItem tmp
    tmp = ""
    
    b = 255
    For i = 7 To 0 Step -1
        b = UByte(b, i, op_clearBit)
        tmp = tmp & b & " "
    Next
    
    List1.AddItem tmp
    showBits CByte(127)
    showBits CInt(&HFF00)
    showBits CLng(&HF0F0F0)
    
    List1.AddItem "1.ToggleBit(0) = " & UByte(1, 0, op_toggleBit)
    List1.AddItem "0.ToggleBit(0) = " & UByte(0, 0, op_toggleBit)
    
    
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim a As Long
    Dim b As Integer
    Dim c As Byte
    
    a = 2147483647 + 1
    MsgBox Hex(2147483647) & " + 1 = " & Hex(a) & " Error: " & Err.Description
    Err.Clear
    
    a = -2147483648# - 1
    MsgBox Hex(-2147483648#) & " - 1 = " & Hex(a) & " Error: " & Err.Description
    Err.Clear
     
    b = 32767 + 1
    MsgBox Hex(32767) & " + 1 = " & Hex(b) & " Error: " & Err.Description
    Err.Clear
    
    b = -32768 - 1
    MsgBox Hex(-32768) & " - 1 = " & Hex(b) & " Error: " & Err.Description
    Err.Clear
     
    c = 0 - 1
    MsgBox "Byte: 0 - 1 = " & c & " Error: " & Err.Description
    Err.Clear
    
    c = &HFF + 2
    MsgBox "Byte: &HFF + 2 = " & c & " Error: " & Err.Description
    Err.Clear
   

End Sub

Private Sub Command2_Click()
    Dim i As Long
    List1.Clear
    List1.AddItem "seed: &h4b4"
    srand &H4B4
    For i = 0 To 25
        List1.AddItem Hex(rand())
    Next
End Sub

Private Sub Command3_Click()
    Dim i(), x, j, k As Long
    
    List1.Clear
    i = Array(1, 2, 4, &HA, &HFFFFFFFF)
    For j = 1 To 2
        For Each x In i
            k = ULong(x, j, op_rol32)
            List1.AddItem "ULong(" & x & "," & j & ", op_rol32) = " & Hex(k)
        Next
    Next
    
End Sub

Private Sub Form_Load()
    
    Dim pth As String
    pth = App.Path & "\UTypes.dll"
    hLib = LoadLibrary(pth)
    List1.AddItem Hex(hLib) & " = " & pth
    
    testLong 2147483647, 1, op_add
    testLong -2147483648#, 1, op_sub
    
    testInt 32767, 1, op_add
    testInt -32768, 1, op_sub
    
    Dim d As Currency
    d = toU64(&HAAAAAAAA, &HBBBBBBBB)
    List1.AddItem Get64(d, mHex)
    List1.AddItem Get64(d, mUnsigned)
    
    d = U64(toU64(&HCCCCCCCC, 0), toU64(0, &HDDDDDDDD), op_add)
    List1.AddItem Get64(d, mHex)
    
    Dim l As Long, hi As Long, lo As Long
    l = ULong(2147483647, 1, op_add)
    List1.AddItem "Unsigned: " & GetUnsigned(l) & " hex:" & Hex(l) & " signed:" & l
    
    d = Str264("1122334455667788", 16)
    List1.AddItem Get64(d, mHex)
    
    fromU64 d, hi, lo
    List1.AddItem Hex(hi) & " " & Hex(lo)
    
    d = Str264("2147483648") 'max signed long + 1
    fromU64 d, hi, lo
    List1.AddItem "hi: " & Hex(hi) & " lo: " & Hex(lo)
    
    d = S64(Str264("-1"), Str264("1"), op_sub)
    List1.AddItem Get64(d, mSigned) & " unsigned: " & Get64(d, mUnsigned)
    
    d = U64(Str264("-1"), Str264("1"), op_sub)
    List1.AddItem Get64(d, mUnsigned)
    
    List1.AddItem "Byte &HFF + 2: " & UByte(&HFF, 2, op_add)
    List1.AddItem "Byte 0 - 2: " & UByte(0, 2, op_sub)
    
    Dim s As String, b() As Byte
    s = "IHATEMATH" '"99eb96dd94c88e975b585d2f28785e36"
    'printf("taking CRC64 of \"99eb96dd94c88e975b585d2f28785e36\" (should be DB7AC38F63413C4E)\n");
    'assert CRC64digest("IHATEMATH") == "E3DCADD69B01ADD1"
    d = crc64s(StrPtr(s))
    List1.AddItem "crc64s(" & s & ") = " & Get64(d)
    
    b = StrConv(s, vbFromUnicode, &H409)
    d = crc64(b(0), UBound(b) + 1)
    List1.AddItem "crc64(" & s & ") = " & Get64(d)
    
    Dim ee As Double, bb() As Byte
    bb = StrConv("AAAAA", vbFromUnicode) 'String(100, "A") 'these match a pure C test
    ee = entropy(bb(0), UBound(bb) + 1)
    List1.AddItem "entropy(AAAAA) = " & ee
    
    bb = StrConv("abcdefghijklmnop", vbFromUnicode)
    ee = entropy(bb(0), UBound(bb) + 1)
    List1.AddItem "entropy(abcdefghijklmnop) = " & ee
    
    
End Sub

Function Get64(v As Currency, Optional m As modes = mHex) As String
    Dim tmp As String, i As Long
    tmp = Space(64)
    i = U642Str(v, tmp, 64, m)
    If i > 0 Then Get64 = Mid(tmp, 1, i)
End Function

Function GetUnsigned(v As Long) As String
    Dim tmp As String, i As Long
    tmp = Space(64)
    i = U2Str(v, tmp, 64)
    If i > 0 Then GetUnsigned = Mid(tmp, 1, i)
End Function


Sub testLong(a As Long, b As Long, opp As op)
    
    Dim ret As Long, o As Variant, msg As String
    o = Array("+", "-", "/", "*", "mod", "xor", "and", "or")
    
    ret = ULong(a, b, opp)
    msg = Hex(a) & " " & o(opp) & " " & Hex(b) & " = " & Hex(ret)
    
    List1.AddItem msg
    Debug.Print msg
End Sub

Sub testInt(a As Integer, b As Integer, opp As op)
    
    Dim ret As Integer, o As Variant, msg As String
    o = Array("+", "-", "/", "*", "mod", "xor", "and", "or")
    
    ret = UInt(a, b, opp)
    msg = Hex(a) & " " & o(opp) & " " & Hex(b) & " = " & Hex(ret)
    
    List1.AddItem msg
    Debug.Print msg
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FreeLibrary hLib
End Sub
