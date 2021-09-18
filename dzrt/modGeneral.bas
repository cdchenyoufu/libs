Attribute VB_Name = "modGeneral"
Option Explicit

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'just a useful bonus method to add into the C dll might as well...
Declare Function ut_crc32 Lib "utypes.dll" Alias "crc32" (ByRef b As Byte, ByVal sz As Long) As Long
Declare Function ut_crc32w Lib "utypes.dll" Alias "crc32w" (ByVal b As Long, ByVal sz As Long) As Long

'Private Declare Function RtlGetCompressionWorkSpaceSize Lib "NTDLL" (ByVal flags As Integer, WorkSpaceSize As Long, UNKNOWN_PARAMETER As Long) As Long
'Private Declare Function NtAllocateVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, ByVal NumBits As Long, regionsize As Long, ByVal flags As Long, ByVal ProtectMode As Long) As Long
'Private Declare Function RtlCompressBuffer Lib "NTDLL" (ByVal flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, ByVal UNKNOWN_PARAMETER As Long, OutputSize As Long, ByVal WorkSpace As Long) As Long
'Private Declare Function RtlDecompressBuffer Lib "NTDLL" (ByVal flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, OutputSize As Long) As Long
'Private Declare Function NtFreeVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, regionsize As Long, ByVal flags As Long) As Long
 
Private Declare Function CryptBinaryToString Lib "Crypt32" Alias "CryptBinaryToStringW" (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
Private Declare Function CryptStringToBinary Lib "Crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Private Declare Function RtlGetCompressionWorkSpaceSize Lib "NTDLL" (ByVal Flags As Integer, WorkSpaceSize As Long, UNKNOWN_PARAMETER As Long) As Long
Private Declare Function NtAllocateVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, ByVal NumBits As Long, regionsize As Long, ByVal Flags As Long, ByVal ProtectMode As Long) As Long
Private Declare Function RtlCompressBuffer Lib "NTDLL" (ByVal Flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, ByVal UNKNOWN_PARAMETER As Long, OutputSize As Long, ByVal WorkSpace As Long) As Long
Private Declare Function RtlDecompressBuffer Lib "NTDLL" (ByVal Flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, OutputSize As Long) As Long
Private Declare Function NtFreeVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, regionsize As Long, ByVal Flags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
 
Public hUTypes As Long
 
Global Const LANG_US = &H409
Const STATUS_SUCCESS = 0
Const STATUS_BUFFER_ALL_ZEROS = &H117
Const STATUS_INVALID_PARAMETER = &HC000000D
Const STATUS_UNSUPPORTED_COMPRESSION = &HC000025F
Const STATUS_NOT_SUPPORTED_ON_SBS = &HC0000300
Const STATUS_BUFFER_TOO_SMALL = &HC0000023
Const STATUS_BAD_COMPRESSION_BUFFER = &HC0000242

Const COMPRESSION_FORMAT_LZNT1 = &H2
Const COMPRESSION_ENGINE_STANDARD = &H0   '// Standard compression
Const COMPRESSION_ENGINE_MAXIMUM = &H100  '// Maximum compression
 

Function ensureDll(dllName) As Boolean
    
    On Error Resume Next
    Dim hDll As Long, pth As String, b() As Byte, f As Long, a As Long, basename As String
    
    hDll = GetModuleHandle(dllName)
    
    If hDll <> 0 Then
        ensureDll = True
        Exit Function
    End If
    
    a = InStrRev(dllName, ".")
    If a < 1 Then
        basename = dllName
        dllName = dllName & ".dll"
    Else
        basename = Mid(dllName, 1, a)
    End If
    
    pth = App.path & "\" & dllName
    If Not FileExists(pth) Then pth = App.path & "\..\" & dllName
    If Not FileExists(pth) Then pth = App.path & "\..\..\" & dllName
    If Not FileExists(pth) Then pth = App.path & "\..\..\..\" & dllName
    
    If Not FileExists(pth) Then
        pth = App.path & "\" & dllName
        b() = LoadResData(basename, "DLLS")
        
        If AryIsEmpty(b) Then
            MsgBox "Failed to find " & dllName & " in resource?"
            Exit Function
        End If
        
        f = FreeFile
        Open pth For Binary As f
        Put f, , b()
        Close f
        
        'MsgBox "Dropped zlib.dll to: " & pth & " - Err: " & Err.Number
    End If
    
    If Not FileExists(pth) Then
        MsgBox "Failed to write " & dllName & " to disk from resource?"
        Exit Function
    End If
        
    hDll = LoadLibrary(pth)
        
    If hDll = 0 Then
        MsgBox "Failed to load " & dllName & " library?"
        Exit Function
    End If
    
    ensureDll = True
    
End Function
 
 
Function ensureUTypes() As Boolean
    
    On Error Resume Next
    
    If hUTypes <> 0 Then
        ensureUTypes = True
        Exit Function
    End If
    
    Dim pth As String, b() As Byte, f As Long
    
    pth = App.path & "\UTypes.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\UTypes.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\..\UTypes.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\..\..\UTypes.dll"
    
    If Not FileExists(pth) Then

        pth = App.path & "\UTypes.dll"
        b() = LoadResData("UTYPES", "DLLS")
        If AryIsEmpty(b) Then
            MsgBox "Failed to find UTypes.dll in resource?"
            Exit Function
        End If
        
        f = FreeFile
        Open pth For Binary As f
        Put f, , b()
        Close f
        
        'MsgBox "Dropped utypes.dll to: " & pth & " - Err: " & Err.Number
    End If
    
    If Not FileExists(pth) Then
        MsgBox "Failed to write UTypes.dll to disk from resource?"
        Exit Function
    End If
        
    hUTypes = LoadLibrary(pth)
        
    If hUTypes = 0 Then
        MsgBox "Failed to load UTypes.dll library?"
        Exit Function
    End If
    
    ensureUTypes = True
    
End Function

Function FileExists(path) As Boolean
  On Error GoTo hell
    
  '.(0), ..(0) etc cause dir to read it as cwd!
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If InStr(path, Chr(0)) > 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function


Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim X
  
    X = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

'Public Function RTLCompress(data() As Byte, Out() As Byte) As Long
'   Dim WorkSpaceSize As Long
'   Dim WorkSpace As Long
'   Dim lOutputSize As Long
'
'   ReDim Out(UBound(data) * 1.13 + 4)
'   RtlGetCompressionWorkSpaceSize 2, WorkSpaceSize, 0
'   NtAllocateVirtualMemory -1, WorkSpace, 0, WorkSpaceSize, 4096, 64
'   RtlCompressBuffer 2, VarPtr(data(0)), UBound(data) + 1, VarPtr(Out(0)), (UBound(data) * 1.13 + 4), 0, lOutputSize, WorkSpace
'   NtFreeVirtualMemory -1, WorkSpace, 0, 16384
'   ReDim Preserve Out(lOutputSize)
'   RTLCompress = lOutputSize
'
'End Function
'
'Public Function RTLDeCompress(data() As Byte, dest() As Byte) As Long
'   If UBound(data) Then
'       Dim lBufferSize As Long
'       ReDim dest(UBound(data) * 12.5)
'       RtlDecompressBuffer 2, VarPtr(dest(0)), (UBound(data) * 12.5), VarPtr(data(0)), UBound(data), lBufferSize
'       If lBufferSize Then
'           ReDim Preserve dest(lBufferSize - 1)
'           RTLDeCompress = lBufferSize - 1
'       End If
'   End If
'End Function

Public Function RTLCompress(data() As Byte, bOut() As Byte, Optional max As Boolean = False) As Boolean
    Dim WorkSpaceSize As Long
    Dim WorkSpace As Long
    Dim lCompress As Long
    Dim Flags As Long
    Dim ret As Long
    
    If AryIsEmpty(data) Then Exit Function

    Flags = COMPRESSION_FORMAT_LZNT1
    If max Then Flags = Flags Or COMPRESSION_ENGINE_MAXIMUM
    
    ReDim bOut(UBound(data) * 1.13 + 4)
    RtlGetCompressionWorkSpaceSize 2, WorkSpaceSize, 0
    NtAllocateVirtualMemory -1, WorkSpace, 0, WorkSpaceSize, 4096, 64
    ret = RtlCompressBuffer(Flags, VarPtr(data(0)), UBound(data) + 1, VarPtr(bOut(0)), (UBound(data) * 1.13 + 4), 0, lCompress, WorkSpace)
    NtFreeVirtualMemory -1, WorkSpace, 0, 16384
    
    If ret = STATUS_SUCCESS Then
        RTLCompress = True
        ReDim Preserve bOut(lCompress + 1)
    Else
        Erase bOut
    End If
    
    
'    If CryptBinaryToString(Out(0), lCompress + 1, &H4, 0&, WorkSpaceSize) <> 0 Then
'        strHexOut = String(WorkSpaceSize - 1, 0)
'        If CryptBinaryToString(Out(0), lCompress + 1, &H4, StrPtr(strHexOut), WorkSpaceSize) <> 0 Then
'            strHexOut = Replace$(Replace$(Replace$(strHexOut, " ", vbNullString), vbNewLine, vbNullString), vbTab, vbNullString)
'            CompressToHex = True
'        End If
'    End If

End Function

Public Function RTLDeCompress(ByRef bytBuf() As Byte, BitOut() As Byte, Optional max As Boolean = False) As Boolean
    Dim lBufferSize As Long
    Dim WorkSpaceSize As Long
    Dim dwActualUsed As Long
    Dim ret As Long
    Dim Flags As Long
    
    If AryIsEmpty(bytBuf) Then Exit Function
    
    Flags = COMPRESSION_FORMAT_LZNT1
    If max Then Flags = Flags Or COMPRESSION_ENGINE_MAXIMUM
    
    ReDim BitOut(UBound(bytBuf) * 12.5)
    ret = RtlDecompressBuffer(Flags, VarPtr(BitOut(0)), (UBound(bytBuf) * 12.5), VarPtr(bytBuf(0)), UBound(bytBuf), lBufferSize)
    
    If ret = STATUS_SUCCESS And lBufferSize Then
        ReDim Preserve BitOut(lBufferSize - 1)
        RTLDeCompress = True
    Else
        Erase BitOut
        Exit Function
    End If
    
End Function

' Sub CopyMemory(Destination As Long, Source As Long, ByVal Length As Long)
'     CopyMemory_ ByVal Destination, ByVal Source, Length
' End Sub
 
'this function will convert any of the following to a byte array:
'   read a file if path supplied and allowFilePaths = true
'   byte(), integer() or long() arrays
'   all other data types it will attempt to convert them to string, then to byte array
'   if the data type you pass can not be converted with cstr() it will throw an error.
'   no other types make sense to support explicitly
'   this assumes all arrays are 0 based..
Function LoadData(fileStringOrByte, Optional allowFilePaths As Boolean = True) As Byte()
    
    Dim f As Long
    Dim Size As Long
    Dim b() As Byte
    Dim L() As Long    ' must cast to specific array type or
    Dim i() As Integer ' else you are reading part of the variant structure..
    
    If allowFilePaths And FileExists(fileStringOrByte) Then
         f = FreeFile
         Open fileStringOrByte For Binary As f
         ReDim b(lof(f) - 1)
         Get f, , b()
         Close f
    ElseIf TypeName(fileStringOrByte) = "Byte()" Then
        b() = fileStringOrByte
    ElseIf TypeName(fileStringOrByte) = "Integer()" Then
        i() = fileStringOrByte
        ReDim b((UBound(i) * 2) - 1)
        CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(i(0)), UBound(b) + 1
    ElseIf TypeName(fileStringOrByte) = "Long()" Then
        L() = fileStringOrByte
        ReDim b((UBound(L) * 4) - 1)
        CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(L(0)), UBound(b) + 1
    Else
        b() = StrConv(CStr(fileStringOrByte), vbFromUnicode, LANG_US)
    End If
    
    LoadData = b()
    
End Function


'ported from Detect It Easy - Binary::calculateEntropy
'   https://github.com/horsicq/DIE-engine/blob/master/binary.cpp#L2319
Function fileEntropy(pth As String, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    Dim sz As Long
    Dim fEntropy As Single
    Dim bytes(255) As Single
    Dim temp As Single
    Dim nSize As Long
    Dim nTemp As Long
    Const BUFFER_SIZE = &H1000
    Dim Buf() As Byte
    Dim f As Long
    Dim i As Long
    
    On Error Resume Next
    
    f = FreeFile
    Open pth For Binary Access Read As f
    If Err.Number <> 0 Then GoTo ret0
    
    sz = lof(f) - 1
    
    If leng = 0 Then GoTo ret0
    
    If leng = -1 Then
        leng = sz - offset
        If leng = 0 Then GoTo ret0
    End If
    
    If offset >= sz Then GoTo ret0
    If offset + leng > sz Then GoTo ret0
    
    Seek f, offset
    nSize = leng
    fEntropy = 1.44269504088896
    ReDim Buf(BUFFER_SIZE)
    
    'read the file in chunks and count how many times each byte value occurs
    While (nSize > 0)
        nTemp = IIf(nSize < BUFFER_SIZE, nSize, BUFFER_SIZE)
        If nTemp <> BUFFER_SIZE Then ReDim Buf(nTemp) 'last chunk, partial buffer
        Get f, , Buf()
        For i = 0 To UBound(Buf)
            bytes(Buf(i)) = bytes(Buf(i)) + 1
        Next
        nSize = nSize - nTemp
    Wend
    
    For i = 0 To UBound(bytes)
        temp = bytes(i) / CSng(leng)
        If temp <> 0 Then
            fEntropy = fEntropy + (-Log(temp) / Log(2)) * bytes(i)
        End If
    Next
    
    Close f
    fileEntropy = Round(fEntropy / CSng(leng), 3)
    
Exit Function
ret0:
    Close f
End Function


Function memEntropy(Buf() As Byte, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    Dim sz As Long
    Dim fEntropy As Single
    Dim bytes(255) As Single
    Dim temp As Single
    Const BUFFER_SIZE = &H1000
    Dim i As Long
    
    sz = UBound(Buf)
    
    If leng = 0 Then GoTo ret0
    If leng = -1 Then
        leng = sz - offset
        If leng = 0 Then GoTo ret0
    End If
    
    If offset >= sz Then GoTo ret0
    If offset + leng > sz Then GoTo ret0
    
    fEntropy = 1.44269504088896
    
    While (offset < sz)
        'count each byte value occurance
        bytes(Buf(offset)) = bytes(Buf(offset)) + 1
        offset = offset + 1
    Wend
    
    For i = 0 To UBound(bytes)
        temp = bytes(i) / CSng(leng)
        If temp <> 0 Then
            fEntropy = fEntropy + (-Log(temp) / Log(2)) * bytes(i)
        End If
    Next
    
    memEntropy = Round(fEntropy / CSng(leng), 3)
    
Exit Function
ret0:
End Function


'supports %x, %c, %s, %d, %10d \t \n %% %-20s
Function printf(ByVal msg As String, vars() As Variant) As String

    Dim t
    Dim ret As String
    Dim i As Long, base, marker
    
    msg = Replace(msg, Chr(0), Empty)
    msg = Replace(msg, "\t", vbTab)
    msg = Replace(msg, "\n", vbCrLf) 'simplified
    msg = Replace(msg, "%%", Chr(0))
    
    t = split(msg, "%")
    If UBound(t) <> UBound(vars) + 1 Then
        MsgBox "Format string mismatch.."
        Exit Function
    End If
    
    ret = t(0)
    For i = 1 To UBound(t)
        base = t(i)
        marker = ExtractSpecifier(base)
        If Len(marker) > 0 Then
            ret = ret & HandleMarker(base, marker, vars(i - 1))
        Else
            ret = ret & base
        End If
    Next
    
    ret = Replace(ret, Chr(0), "%")
    printf = ret
    
End Function

Private Function HandleMarker(base, ByVal marker, var) As String
    Dim newBase As String
    Dim mType As Integer
    Dim nVal As String
    Dim spacer As String
    Dim prefix As String
    Dim Count As Long
    Dim leftJustify As Boolean
    
    If Len(base) > Len(marker) Then
        newBase = Mid(base, Len(marker) + 1) 'remove the marker..
    End If
    
    mType = Asc(Mid(marker, Len(marker), 1))  'last character
    
    Select Case mType
        Case Asc("x"): nVal = Hex(var)
        Case Asc("X"): nVal = UCase(Hex(var))
        Case Asc("s"): nVal = var
        Case Asc("S"): nVal = UCase(var)
        Case Asc("c"): nVal = Chr(var)
        Case Asc("d"): nVal = var
        
        Case Else: nVal = var
    End Select
    
    If Len(marker) > 1 Then 'it has some more formatting involved..
        marker = Mid(marker, 1, Len(marker) - 1) 'trim off type
        If Left(marker, 1) = "-" Then
            leftJustify = True
            marker = Mid(marker, 2)  'trim off left justify marker
        End If
        If Left(marker, 1) = "0" Then
            spacer = "0"
            marker = Mid(marker, 2)
        Else
            spacer = " "
        End If
        Count = CLng(marker) - Len(nVal)
        If Count > 0 Then prefix = String(Count, spacer)
    End If
    
    If leftJustify Then
        HandleMarker = nVal & prefix & newBase
    Else
        HandleMarker = prefix & nVal & newBase
    End If
    
End Function

Private Function ExtractSpecifier(v)
    
    Dim ret As String
    Dim b() As Byte
    Dim i As Long
    If Len(v) = 0 Then Exit Function
    
    b() = StrConv(v, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(b)
        ret = ret & Chr(b(i))
        If b(i) = Asc("x") Then Exit For
        If b(i) = Asc("X") Then Exit For
        If b(i) = Asc("c") Then Exit For
        If b(i) = Asc("s") Then Exit For
        If b(i) = Asc("S") Then Exit For
        If b(i) = Asc("d") Then Exit For
    Next
    
    ExtractSpecifier = ret
    
End Function

'Public Function ado_ConnectionString(dbServer As dbServers, dbName As String, Optional server As String, Optional Port = 3306, Optional User As String, Optional pass As String) As String
'    Dim dbPath As String, baseString As String, blnInlineAuth As Boolean
'
'    Select Case dbServer
'        Case Access
'            baseString = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=____;"
'        Case FileDsn
'            baseString = "FILEDSN=____;"
'        Case DSN
'            baseString = "DSN=____;"
'        Case dBase
'            baseString = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=____;"
'        Case mysql
'            baseString = "Driver={mySQL};Server=" & server & ";Port=" & Port & ";Stmt=;Option=16834;Database=____;"
'        Case MsSql2k
'            baseString = "Driver={SQL Server};Server=" & server & ";Database=____;"
'        Case JetAccess2k
'            baseString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=____;" & _
'                         "User Id=" & User & ";" & _
'                         "Password=" & pass & ";"
'                         blnInlineAuth = True
'    End Select
'
'
'    If Not blnInlineAuth Then
'        If User <> Empty Then baseString = baseString & "Uid:" & User & ";"
'        If pass <> Empty Then baseString = baseString & "Pwd:" & User & ";"
'    End If
'
'    '%AP% is like enviromental variable for app.path i am lazy :P
'    dbPath = Replace(dbName, "%AP%", App.path)
'
'    ado_ConnectionString = Replace(baseString, "____", dbPath)
'
'End Function

Function GetParentFolder(path) As String
    Dim tmp() As String
    Dim my_path
    Dim ub As String
    
    On Error GoTo hell
    If Len(path) = 0 Then Exit Function
    
    my_path = path
    While Len(my_path) > 0 And Right(my_path, 1) = "\"
        my_path = Mid(my_path, 1, Len(my_path) - 1)
    Wend
    
    tmp = split(my_path, "\")
    tmp(UBound(tmp)) = Empty
    my_path = Replace(Join(tmp, "\"), "\\", "\")
    If VBA.Right(my_path, 1) = "\" Then my_path = Mid(my_path, 1, Len(my_path) - 1)
    
    GetParentFolder = my_path
    Exit Function
    
hell:
    GetParentFolder = Empty
    
End Function
