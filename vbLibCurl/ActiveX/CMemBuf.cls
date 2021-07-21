VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'dzzie

Public growthCycles As Long

Private binaryData() As Byte
Private curSize As Long
Private Const OneMB As Long = 1000000
'Private Const OneMB As Long = 10 'grow test

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Friend Property Get rawBufSize() As Long
   rawBufSize = UBound(b) + 1
End Property

Property Get size() As Long
    size = curSize
End Property

Property Get binData() As Byte()
    Dim b() As Byte
    If curSize = 0 Then Exit Property
    ReDim b(curSize - 1)
    CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(binaryData(0)), curSize
    binData = b()
End Property

Property Get asString() As String
    If curSize = 0 Then Exit Property
    asString = StrConv(binData, vbUnicode, &H409)
End Property

Sub memAppendBuf(rawBytes As Long, totalBytes As Long)

    If curSize + totalBytes > UBound(binaryData) Then
        ReDim Preserve binaryData(curSize + totalBytes + OneMB)
        growthCycles = growthCycles + 1
    End If
    
    CopyMemory ByVal VarPtr(binaryData(curSize)), ByVal rawBytes, totalBytes
    curSize = curSize + totalBytes
    
End Sub

Sub appendBuf(b() As Byte)
    
    Dim ptr As Long
    Dim sz As Long
    
    sz = UBound(b) + 1
    ptr = curSize
    
    If ptr + sz > UBound(binaryData) Then
        ReDim Preserve binaryData(curSize + sz + OneMB)
        growthCycles = growthCycles + 1
    End If
    
    CopyMemory ByVal VarPtr(binaryData(ptr)), ByVal VarPtr(b(0)), sz
    curSize = curSize + sz
    
End Sub

Private Sub Class_Initialize()
    ReDim binaryData(OneMB)
End Sub
