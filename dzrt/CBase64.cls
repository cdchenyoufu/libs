VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'base for this was taken from PSC
'http://www.pscode.com/xq/ASP/txtCodeId.3600/lngWId.1/qx/vb/scripts/ShowCode.htm
'and encode routine from
'http://pscode.com/xq/ASP/txtCodeId.4174/lngWId.1/qx/vb/scripts/ShowCode.htm
'which in turn was based on fx from Sebastian...true power of open source :)
Option Explicit

Private Alphabet() As Byte
Private decodeTable(255) As Byte
Private Const LANG_US = &H409

Private Sub Class_Initialize()
    SetCustomAlphabet "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
End Sub

Sub DefaultAlphabet()
    Class_Initialize
End Sub

Function SetCustomAlphabet(ByVal string65 As String) As Boolean
    Dim i As Long
    
    If Len(string65) <> 65 Then
        If Len(string65) = 64 And Right(string65, 1) <> "=" Then
            string65 = string65 & "="
        Else
            MsgBox "Base64 Custom alphabet must be 65 characters long", vbInformation
            Exit Function
        End If
    End If
    
    Alphabet = StrConv(string65, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(decodeTable)
        decodeTable(i) = &HFF
    Next
    
    Dim b As Byte
    For i = 0 To UBound(Alphabet)
        b = Alphabet(i)
        decodeTable(b) = i
    Next
    
    SetCustomAlphabet = True
    
End Function

Function Encode(bArrayStringOrFile, Optional retAsString As Boolean = True) As Variant
    
    Dim b() As Byte
    Dim b2() As Byte
    
    On Error Resume Next
    
    If TypeName(bArrayStringOrFile) = "Byte()" Then
        b() = bArrayStringOrFile
    ElseIf FileExists(bArrayStringOrFile) Then
        LoadFile bArrayStringOrFile, b()
    Else
        If Len(bArrayStringOrFile) = 0 Then Exit Function
        b() = StrConv(CStr(bArrayStringOrFile), vbFromUnicode, LANG_US)
    End If
    
    b2 = EncodeBytes(b)
    
    If retAsString Then
        Encode = StrConv(b2, vbUnicode, LANG_US)
    Else
        Encode = b2
    End If
    
    
End Function

Function Decode(bArrayStringOrFile, ByRef dataOut As Variant, Optional retAsString As Boolean = True) As Boolean
    
    Dim b() As Byte
    Dim b2() As Byte
    Dim ok As Boolean
    
    On Error GoTo hell
    
    If TypeName(bArrayStringOrFile) = "Byte()" Then
        b() = bArrayStringOrFile
    ElseIf FileExists(bArrayStringOrFile) Then
        LoadFile bArrayStringOrFile, b()
    Else
        If Len(bArrayStringOrFile) = 0 Then Exit Function
        b() = StrConv(CStr(bArrayStringOrFile), vbFromUnicode, LANG_US)
    End If
    
    ok = DecodeBytes(b, b2)
    
    If retAsString Then
        dataOut = StrConv(b2, vbUnicode, LANG_US)
    Else
        dataOut = b2
    End If
    
hell:
    Decode = ok
    
End Function

Public Function EncodeBytes(ByRef DataIn() As Byte, Optional autoSpan As Boolean = True) As Byte()

  Dim DataTemp(2) As Byte
  Dim iTemp As Long
  Dim iLoop As Long
  Dim BytesDataIn As Long
  Dim BytesDataOut As Long
  Dim ExtraBytes As Integer
  Dim dataOut() As Byte
    
  BytesDataIn = UBound(DataIn) + 1 'length of the string
  ExtraBytes = (BytesDataIn Mod 3)
  If ExtraBytes = 0 Then
      BytesDataOut = ((BytesDataIn / 3) * 4)  ' how many bytes will the encoded string have
  Else
      BytesDataOut = (((BytesDataIn \ 3) + 1) * 4) ' how many bytes will the encoded string have
  End If
  
  ReDim dataOut(BytesDataOut - 1)
  
  For iLoop = 0 To BytesDataIn - ExtraBytes - 1 Step 3
      dataOut(iTemp) = Alphabet((DataIn(iLoop) \ 4) And &H3F)
      dataOut(iTemp + 1) = Alphabet((DataIn(iLoop) And &H3) * 16 Or (DataIn(iLoop + 1) \ 16) And &HF)
      dataOut(iTemp + 2) = Alphabet((DataIn(iLoop + 1) And &HF) * 4 Or (DataIn(iLoop + 2) \ 64) And &H3)
      dataOut(iTemp + 3) = Alphabet(DataIn(iLoop + 2) And &H3F)
      iTemp = iTemp + 4
  Next
  If ExtraBytes = 1 Then ' if there is 1 byte remaining
      'read 1 byte, the second in 0
      DataTemp(0) = DataIn(UBound(DataIn))
      DataTemp(1) = 0
      DataTemp(2) = 0
      dataOut(UBound(dataOut) - 3) = Alphabet((DataTemp(0) \ 4) And &H3F)
      dataOut(UBound(dataOut) - 2) = Alphabet((DataTemp(0) And &H3) * 16 Or (DataTemp(1) \ 16) And &HF)
      dataOut(UBound(dataOut) - 1) = 61
      dataOut(UBound(dataOut)) = 61
  ElseIf ExtraBytes = 2 Then 'if there are 2 bytes remaining
      'read 2 bytes, the third is 0
      DataTemp(0) = DataIn(UBound(DataIn) - 1)
      DataTemp(1) = DataIn(UBound(DataIn))
      DataTemp(2) = 0
      dataOut(UBound(dataOut) - 3) = Alphabet((DataTemp(0) \ 4) And &H3F)
      dataOut(UBound(dataOut) - 2) = Alphabet((DataTemp(0) And &H3) * 16 Or (DataTemp(1) \ 16) And &HF)
      dataOut(UBound(dataOut) - 1) = Alphabet((DataTemp(1) And &HF) * 4 Or (DataTemp(2) \ 64) And &H3)
      dataOut(UBound(dataOut)) = 61
  End If
  
  If autoSpan Then
        EncodeBytes = Span(dataOut())
  Else
        EncodeBytes = dataOut()
  End If

End Function


Public Function DecodeBytes(ByRef data() As Byte, ByRef out() As Byte) As Boolean
    Dim inp(3) As Byte
    Dim fileIn() As Byte
    Dim pad As Long, Lenght As Long, BytesOut As Long, i As Long, iTemp As Long
    
    
    On Error Resume Next
    
    fileIn() = Unspan(data)
    
    While (UBound(fileIn) + 1) Mod 4 <> 0 'some clients don't ad = pads!!
        ReDim Preserve fileIn(UBound(fileIn) + 1)
        fileIn(UBound(fileIn)) = CByte(Asc("="))
        'MsgBox "pad forced up"
        'this loop may occur a maxium of twice
    Wend
        
    If fileIn(UBound(fileIn) - 1) = Asc("=") Then
        pad = 2
    ElseIf fileIn(UBound(fileIn)) = Asc("=") Then
        pad = 1
    Else
        pad = 0
    End If

    'MsgBox pad
    
    Lenght = UBound(fileIn) ' + 1    'lenght of the string
    BytesOut = ((Lenght / 4) * 3) - pad ' how many bytes will the b64Decoded string have
    ReDim out(BytesOut)

    For i = 0 To Lenght Step 4
        inp(0) = decodeTable(fileIn(i))
        inp(1) = decodeTable(fileIn(i + 1))
        inp(2) = decodeTable(fileIn(i + 2))
        inp(3) = decodeTable(fileIn(i + 3))
        If inp(3) = 64 Or inp(2) = 64 Then
            If inp(3) = 64 And Not (inp(2) = 64) Then
                inp(0) = decodeTable(fileIn(i))
                inp(1) = decodeTable(fileIn(i + 1))
                inp(2) = decodeTable(fileIn(i + 2))
                '2 bytes out
                out(iTemp) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
                out(iTemp + 1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
                GoTo retNow
            ElseIf inp(2) = 64 Then
                inp(0) = decodeTable(fileIn(i))
                inp(1) = decodeTable(fileIn(i + 1))
                '1 byte out
                out(iTemp) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
                GoTo retNow
            End If
        End If
        '3 bytes out
        out(iTemp) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
        out(iTemp + 1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
        out(iTemp + 2) = ((inp(2) And &H3) * 64) Or inp(3)
        iTemp = iTemp + 3
    Next
    
retNow:
    DecodeBytes = (Err.Number = 0)
    
End Function

Public Function Span(bAryOrString, Optional CharsPerLine As Long = 72) As Variant
    
    Dim b() As Byte, i As Long, cnt As Long, J As Long
    Dim tmp() As Byte
    Dim isString As Boolean
    
    If TypeName(bAryOrString) = "Byte()" Then
        If AryIsEmpty(bAryOrString) Then Exit Function
        b() = bAryOrString
    Else
        If Len(bAryOrString) = 0 Then Exit Function
        isString = True
        b() = StrConv(CStr(bAryOrString), vbFromUnicode, LANG_US)
    End If
    
    cnt = UBound(b) / CharsPerLine
    If cnt > 0 Then
        ReDim tmp(UBound(b) + (cnt * 2))
        For i = 1 To UBound(b) + 1
            tmp(J) = b(i - 1)
            J = J + 1
            If i Mod CharsPerLine = 0 Then
                tmp(J) = &HD
                tmp(J + 1) = &HA
                J = J + 2
            End If
        Next
        b() = tmp()
    End If
    
    If isString Then
        Span = StrConv(b, vbUnicode, LANG_US)
    Else
        Span = b()
    End If
            
    
End Function

'Public Function Span(inAry() As Byte, Optional CharsPerLine As Long = 72) As Byte()
'
'
'
'
'    If UBound(inAry) > CharsPerLine Then
'        Dim ret() As Byte: x = 0: cpl = 0
'        While x <= UBound(inAry)
'            If cpl < CharsPerLine Then
'                If x = 0 Then ReDim ret(0) _
'                Else ReDim Preserve ret(UBound(ret) + 1)
'                ret(UBound(ret)) = inAry(x)
'                cpl = cpl + 1
'                x = x + 1
'            ElseIf cpl = CharsPerLine Then
'                ReDim Preserve ret(UBound(ret) + 2)
'                ret(UBound(ret) - 1) = CByte(13)
'                ret(UBound(ret)) = CByte(10)
'                cpl = 0
'            End If
'        Wend
'
'        Span = ret()
'    Else
'        Span = inAry()
'    End If
'
'End Function

'if arg is a string, it returns a string,
'if arg is byte array it returns a byte array
Public Function Unspan(bAryOrString) As Variant
    Dim b() As Byte, i As Long, cnt As Long, J As Long
    Dim tmp() As Byte
    Dim isString As Boolean
    
    If TypeName(bAryOrString) = "Byte()" Then
        b() = bAryOrString
    Else
        isString = True
        b() = StrConv(CStr(bAryOrString), vbFromUnicode, LANG_US)
    End If
    
    For i = 0 To UBound(b)
        If b(i) = 10 Or b(i) = 13 Then cnt = cnt + 1
    Next
    
    If cnt > 0 Then
        If cnt = UBound(b) Then Exit Function
        
        ReDim tmp(UBound(b) - cnt)
        For i = 0 To UBound(b)
            If b(i) <> 10 And b(i) <> 13 Then
                tmp(J) = b(i)
                J = J + 1
            End If
        Next
    
        b() = tmp()
    End If
    
    If isString Then
        Unspan = StrConv(b, vbUnicode, LANG_US)
    Else
        Unspan = b()
    End If
    
End Function

Function EncodeString(ByVal it As String, Optional autoSpan As Boolean = True) As String
   Dim b() As Byte, b2() As Byte
   b() = StrConv(it, vbFromUnicode, LANG_US)
   b2() = EncodeBytes(b, autoSpan)
   EncodeString = ByteArrayToString(Span(b2))
End Function
   
Function EncodeFileToString(ByVal fpath As String, Optional autoSpan As Boolean = True) As String
   Dim b() As Byte, b2() As Byte
   
   If Not FileExists(fpath) Then Exit Function
   
   LoadFile CStr(fpath), b()
   b2() = EncodeBytes(b, autoSpan)
   EncodeFileToString = ByteArrayToString(Span(b2))
   
End Function

Function EncodeFileToFile(ByVal fpath As String, ByVal saveAsPath As String, Optional autoSpan As Boolean = True) As Boolean
   Dim b() As Byte, b2() As Byte
   On Error Resume Next
   LoadFile CStr(fpath), b()
   b2() = EncodeBytes(b, autoSpan)
   WriteFile saveAsPath, Span(b2)
   EncodeFileToFile = (Err.Number = 0)
End Function

Function DecodeString(ByVal mimedata As String, ByRef dataOut) As Boolean
   Dim b() As Byte, b2() As Byte, tmp As String, ok As Boolean
   b() = StrConv(mimedata, vbFromUnicode, LANG_US)
   ok = DecodeBytes(b(), b2())
   dataOut = ByteArrayToString(b2())
   DecodeString = ok
End Function

Function DecodeStringToFile(ByVal fpath As String, ByVal mimedata As String) As Boolean
    Dim b() As Byte, b2() As Byte, ok As Boolean
    On Error Resume Next
    
    If Not FileExists(fpath) Then Exit Function

    b() = StrConv(mimedata, vbFromUnicode, LANG_US)
    ok = DecodeBytes(b(), b2())
    WriteFile CStr(fpath), b2()
    
    DecodeStringToFile = ok
    If Err.Number <> 0 Then DecodeStringToFile = False
    
End Function

Function DecodeFileToFile(ByVal inFile As String, ByVal outFile As String) As Boolean
    Dim b() As Byte, b2() As Byte, tmp As String, ok As Boolean
    
    On Error Resume Next
    
    If Not FileExists(inFile) Then Exit Function
    If FileExists(outFile) Then Kill outFile
    
    LoadFile inFile, b()
    ok = DecodeBytes(b(), b2())
    WriteFile outFile, b2()
    
    DecodeFileToFile = ok
    If Err.Number <> 0 Then DecodeFileToFile = False

End Function

Private Function countOccurances(it, find) As Integer
    Dim tmp() As String
    If InStr(1, it, find, vbTextCompare) < 1 Then countOccurances = 0: Exit Function
    tmp = split(it, find, , vbTextCompare)
    countOccurances = UBound(tmp)
End Function

Private Function ByteArrayToString(ByRef pbArrayInput() As Byte) As String
   ByteArrayToString = StrConv(pbArrayInput, vbUnicode, LANG_US)
End Function

Private Sub LoadFile(fpath, ByRef out() As Byte)
   Dim fsize As Long, f As Long
   
   fsize = FileLen(fpath) - 1
   f = FreeFile
   ReDim out(fsize)
   
   Open fpath For Binary As f
   Get f, , out()
   Close f

End Sub

Private Sub WriteFile(fpath, bAry() As Byte)
    Dim f As Long
    f = FreeFile
    Open fpath For Binary As f
    Put f, , bAry()
    Close f
End Sub

Private Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
     If Err.Number <> 0 Then Exit Function
     FileExists = True
  End If
End Function

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim X
  
    X = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
