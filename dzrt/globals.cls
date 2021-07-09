VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "globals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Const VK_SHIFT = &H10
Const VK_CONTROL As Integer = &H11 'Ctrl
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const LB_GETCOUNT = &H18B
Private Const LB_SETTOPINDEX = &H197
Private startTime As Long

Private Declare Function GetTickCount_ Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function LoadLibrary_ Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary_ Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SendMessage_ Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub Sleep_ Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
 
Private Declare Sub vc_srand Lib "utypes.dll" (ByVal v1 As Long)
Private Declare Function vc_rand Lib "utypes.dll" () As Long
'Private Declare Function crc64s Lib "utypes.dll" (ByVal wStrPtr As Long, Optional asciiOnly As Long = 1) As Currency
Private Declare Function ut_crc64 Lib "utypes.dll" Alias "crc64" (ByRef stream As Byte, ByVal sz As Long) As Currency
Private Declare Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

'all but jet are odbc jet is ole
Public Enum dbServers
    db_Access
    db_JetAccess2k
    db_MsSql2k
    db_mysql
    db_DSN
    db_FileDsn
    db_dBase
    db_mysql5
End Enum

Public Enum CompressTypes
    ct_Zlib = 0
    ct_LZNT1 = 1      'so the compress seems to work..but is 1 byte short?
    ct_LZNT1_MAX = 2  'the decompress works with other known samples, but not with one compressed with our compress..havent been able to fix the error yet!
End Enum

Public Enum InputFormats
    ifHex = 1
    ifDecimal = 2
    ifString = 3
    ifHexDump = 4
End Enum

Public Enum hexOutFormats
    hoDump
    hoSpaced
    hoHexOnly
End Enum

Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = 0
Private Const INTERNET_MAX_URL_LENGTH As Long = 2048
Private Const URL_ESCAPE_PERCENT As Long = &H1000&

Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeA" ( _
    ByVal pszURL As String, _
    ByVal pszEscaped As String, _
    ByRef pcchEscaped As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function UrlUnescape Lib "shlwapi" Alias "UrlUnescapeA" ( _
    ByVal pszURL As String, _
    ByVal pszUnescaped As String, _
    ByRef pcchUnescaped As Long, _
    ByVal dwFlags As Long) As Long



Private OnBits(0 To 31) As Long
Private crc_table() As Long
Private bitPatterns() 'used in testbit as optimization init on first use..

Public base64 As New CBase64
Public hash As New CWinHash
Public fso As New CFileSystem3
Public cmdline As New CCmdLine
Public zlib As New CZlib
Public extCtl As New CExtControls

Public Sub dbgPrint(ParamArray values() As Variant)

    Dim v       As Variant
    Dim sMsg    As String
    Dim bNext   As Boolean
    
    For Each v In values
        If bNext Then
            sMsg = sMsg & Space$(8&)
            sMsg = Left$(sMsg, (Len(sMsg) \ 8&) * 8&)
        End If
        bNext = True
        sMsg = sMsg & CStr(v)
    Next

    modDebugPrint.DebugPrint sMsg
    
End Sub

Public Sub dbgPrintf(ByVal msg As String, ParamArray values() As Variant)

    Dim i As Long, tmp(), printf As String
    
    For i = 0 To UBound(values)
        If IsNull(values(i)) Then
            push tmp, "[Null]"
        ElseIf IsObject(values(i)) Then
            push tmp, "[Object:" & TypeName(values(i)) & "]"
        Else
            push tmp, values(i)
        End If
    Next
    
    printf = modGeneral.printf(msg, tmp)
    modDebugPrint.DebugPrint printf
    
End Sub

Public Sub dbgClear()
     modDebugPrint.DebugPrint "<cls>"
End Sub

Public Sub dbgDiv()
    modDebugPrint.DebugPrint "<div>"
End Sub

Function ExpandEnvStr(str) As String

    Dim Buf As String
    Dim sz As Long
    
    If InStr(str, "%") < 1 Then GoTo errHandler
    
    sz = ExpandEnvironmentStrings(str, ByVal Buf, 0)
    If sz = 0 Then GoTo errHandler
    
    Buf = Space(sz + 100)
    sz = ExpandEnvironmentStrings(str, ByVal Buf, Len(Buf))
    If sz = 0 Then GoTo errHandler
    
    Buf = Left(Buf, sz - 1)
    ExpandEnvStr = Buf
    
Exit Function
errHandler: ExpandEnvStr = str
End Function

Sub srand(seed As Long)
    If ensureUTypes() Then vc_srand seed
End Sub

Function rand() As Long
    If ensureUTypes() Then rand = vc_rand()
End Function

Function crc64(fileByteOrString) As String
    Dim b() As Byte
    Dim v As Currency
    Dim hi&, lo&
    
    If Not ensureUTypes() Then Exit Function 'use the C dll version its faster
    
    b() = LoadData(fileByteOrString)
    If AryIsEmpty(b) Then Exit Function
    
    v = ut_crc64(b(0), UBound(b) + 1)
    
    GetMem4 ByVal VarPtr(v), lo
    GetMem4 ByVal VarPtr(v) + 4, hi
    crc64 = Hex$(hi) & Right$("0000000" & Hex$(lo), 8)
    
End Function

'this function will convert any of the following to a byte array:
'   read a file if path supplied and allowFilePaths = true
'   byte(), integer() or long() arrays
'   all other data types it will attempt to convert them to string, then to byte array
'   if the data type you pass can not be converted with cstr() it will throw an error.
'   no other types make sense to support explicitly
'   this assumes all arrays are 0 based..
Function LoadData(fileStringOrByte, Optional allowFilePaths As Boolean = True) As Byte()
    LoadData = modGeneral.LoadData(fileStringOrByte, allowFilePaths)
End Function

Function tocsv(ParamArray X()) As String
    Dim i As Long
    Dim tmp()
    For i = 0 To UBound(X)
        If IsNull(X(i)) Then
            push tmp, "[Null]"
        ElseIf IsObject(X(i)) Then
            push tmp, "[Obj:" & TypeName(X(i)) & "]"
        Else
            push tmp, X(i)
        End If
    Next
    tocsv = Join(tmp, ",") & vbCrLf
End Function


Function GetTickCount() As Long
    GetTickCount = GetTickCount_()
End Function

Function LoadLibrary(ByVal lpLibFileName As String) As Long
    LoadLibrary = LoadLibrary_(lpLibFileName)
End Function

 Function FreeLibrary(ByVal hLibModule As Long) As Long
    FreeLibrary = FreeLibrary_(hLibModule)
 End Function
 
 Sub Sleep(ByVal dwMilliseconds As Long)
    Sleep_ dwMilliseconds
 End Sub
 
 Sub memcpy(Destination As Long, Source As Long, ByVal Length As Long)
     CopyMemory ByVal Destination, ByVal Source, Length
 End Sub
 
 Function SendMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    SendMessage = SendMessage_(hWnd, msg, wParam, lParam)
 End Function

Property Get MAX_LONG() As Long
    MAX_LONG = 2147483647
End Property

Property Get MIN_LONG() As Long
    MIN_LONG = -2147483648#
End Property

Property Get MAX_INT() As Integer
    MAX_INT = 32767
End Property

Property Get MIN_INT() As Integer
    MIN_INT = -32768
End Property

Property Get LANG_US() As Long
    LANG_US = &H409
End Property

Sub StartBenchMark(Optional ByRef t As Long = -111)
    If t = -111 Then
        startTime = GetTickCount()
    Else
        t = GetTickCount()
    End If
End Sub

Function EndBenchMark(Optional ByRef t As Long = -111) As String
    Dim endTime As Long, loadTime As Long
    endTime = GetTickCount()
    If t = -111 Then
        loadTime = endTime - startTime
    Else
        loadTime = endTime - t
    End If
    EndBenchMark = loadTime / 1000 & " seconds"
End Function

Public Function FileSize(fpath As String) As String
    Dim fsize As Long
    Dim szName As String
    On Error GoTo hell
    
    fsize = FileLen(fpath)
    
    szName = " bytes"
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Kb"
    End If
    
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Mb"
    End If
    
    FileSize = fsize & szName
    
    Exit Function
hell:
    
End Function

Function HexStrToLng(ByVal X As String, Optional ByRef hadErr As Boolean) As Long
    On Error Resume Next
    X = trim(X)
    If InStr(1, X, "0x", vbTextCompare) > 0 Then X = Replace(X, "0x", Empty, , , vbTextCompare)
    If Left(X, 1) <> "&h" Then X = "&h" & X
    HexStrToLng = CLng(X)
    hadErr = (Err.Number <> 0)
End Function

Function Extract(Value, marker1, marker2, ByRef outVar, _
                Optional includeMarkers As Boolean = False, _
                Optional start As Long = 1, _
                Optional ByRef lastPos As Long, _
                Optional method As VbCompareMethod = vbBinaryCompare _
) As Long

    Dim a As Long, b As Long
    
    lastPos = 0
    outVar = Empty
    
    If Len(marker1) = 0 Then
        a = 1
    Else
        a = InStr(start, Value, marker1, method)
        If a < 1 Then Exit Function
    End If
    
    a = a + Len(marker1)
    If Len(marker2) = 0 Then
        outVar = Mid(Value, a)
    Else
        b = InStr(a, Value, marker2, method)
        If b < 1 Then Exit Function
        
        lastPos = b + Len(marker2)
        outVar = Mid(Value, a, b - a)
    End If
    
    Extract = Len(outVar)
    If includeMarkers And Len(outVar) > 0 Then outVar = marker1 & outVar & marker2
    
End Function

Public Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIde = False
    Exit Function
hell:
    isIde = True
End Function


Function TopMost(frm As Object, Optional ontop As Boolean = True)
    On Error Resume Next
    Dim s
    s = IIf(ontop, HWND_TOPMOST, HWND_NOTOPMOST)
    SetWindowPos frm.hWnd, s, frm.Left / 15, frm.Top / 15, frm.Width / 15, frm.Height / 15, 0
End Function

Public Function GetShortName(ByVal sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long
    Dim iCreated As Boolean
    
    'the path must actually exist to get the short path name !!
    If Not FileExists(sFile) Then
        iCreated = True
        WriteFile sFile, ""
    End If
    
    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)
    
    If Len(GetShortName) = 0 Then GetShortName = sFile
    
    If iCreated Then Kill sFile
    
End Function

Sub FormPos(fform As Object, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz, i, ff, def
    f = split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.Name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.Name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Sub SaveMySetting(key, Value)
    SaveSetting App.EXEName, "Settings", key, Value
End Sub

Function GetMySetting(key, Optional defaultval = "")
    GetMySetting = GetSetting(App.EXEName, "Settings", key, defaultval)
End Function

Function FolderExists(path) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Asc(Left(tmp, 1)) = 0 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
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

Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error Resume Next

    Do While 1
        Err.Clear
        Randomize
        tmp = Round(Timer * Now * Rnd(), 0)
        RandomNum = tmp
        If Err.Number = 0 Then Exit Function
        If tries < 100 Then
            tries = tries + 1
        Else
            Exit Do
        End If
    Loop
    
    RandomNum = GetTickCount
    
End Function

Public Function randInRange(ByVal MaxValue As Long, Optional ByVal MinValue As Long = 0)
  On Error Resume Next
  Randomize Timer
  randInRange = Int((MaxValue - MinValue + 1) * Rnd) + MinValue
End Function

Function ctrlDown() As Boolean
    ctrlDown = (GetKeyState(VK_CONTROL) < 0)
End Function

Function shiftDown() As Boolean
    shiftDown = (GetKeyState(VK_SHIFT) < 0)
End Function


'my traditional simple prototype for general use
Function ReadFile(filename) As String
    Dim ret As String
    If Not ReadFile2(filename, ret) Then Exit Function
    ReadFile = ret
End Function

'can load a byte array, string, or string array of lines split at either crlf or lf
Function ReadFile2(filename, ByRef data) As Boolean 'this one should be binary safe...
  On Error GoTo hell
  Dim f As Long
  Dim b() As Byte
  Dim tmp As String
  Dim t2() As String
  
  f = FreeFile
  Open filename For Binary As #f
  ReDim b(lof(f) - 1)
  Get f, , b()
  Close #f
  
  If TypeName(data) = "Byte()" Then 'Or TypeName(data) = "Variant"
        data = b()
  ElseIf TypeName(data) = "String()" Then
        tmp = CStr(StrConv(b(), vbUnicode, LANG_US))
        If InStr(tmp, vbCrLf) > 0 Then
            t2 = split(tmp, vbCrLf)
        Else
            t2 = split(tmp, vbLf)
        End If
        data = t2
  Else
        data = CStr(StrConv(b(), vbUnicode, LANG_US))
  End If
  
  ReadFile2 = True
  
  Exit Function
hell:   ReadFile2 = False
End Function

Function WriteFile(path, it) As Boolean 'this one should be binary safe...
    On Error GoTo hell
    Dim b() As Byte, f As Long
    
    If FileExists(path) Then Kill path
    f = FreeFile
    
    If TypeName(it) = "Byte()" Then
        b() = it
    Else
        b() = StrConv(it, vbFromUnicode, LANG_US)
    End If
    
    Open path For Binary As #f
    Put f, , b()
    Close f
    
    WriteFile = True
    Exit Function
    
hell: WriteFile = False
End Function

Function AppendFile(path, it) As Boolean  'not binary safe
    Dim f As Long
    On Error Resume Next
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
    AppendFile = (Err.Number = 0)
End Function

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo Init
    Dim X
       
    X = UBound(ary)
    ReDim Preserve ary(X + 1)
    
    If IsObject(Value) Then
        Set ary(X + 1) = Value
    Else
        ary(X + 1) = Value
    End If
    
    Exit Sub
Init:
    ReDim ary(0)
    If IsObject(Value) Then
        Set ary(0) = Value
    Else
        ary(0) = Value
    End If
End Sub

Function pop(ary, Optional count As Long = 1) As Variant
    On Error GoTo isEmpty
    Dim X
    
    X = UBound(ary)
    
    If count > X Then
        Erase ary
        Exit Function
    End If
    
    ReDim Preserve ary(X - count + 1)
        
    If IsObject(ary(X)) Then
        Set pop = ary(X)
    Else
        pop = ary(X)
    End If
    
    ReDim Preserve ary(X - count)
    
    Exit Function
isEmpty: Erase ary
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim X
    If IsObject(UBound(ary)) Then AryIsEmpty = False
    'x = UBound(ary)
  Exit Function
oops: AryIsEmpty = True
End Function

Function AryReverse(ary)
    Dim tmp, i, al() As Long, ai() As Long, ass() As String, av() As Variant, ab() As Byte
    
    If Not IsArray(ary) Then Exit Function
    
    If TypeName(ary) = "Long()" Then
        tmp = al
    ElseIf TypeName(ary) = "Integer()" Then
        tmp = ai
    ElseIf TypeName(ary) = "String()" Then
        tmp = ass
    ElseIf TypeName(ary) = "Variant()" Then
        tmp = av
    ElseIf TypeName(ary) = "Byte()" Then
        tmp = ab
    Else
        MsgBox "AryReverse Add support for: " & TypeName(ary)
    End If
        
    If AryIsEmpty(ary) Then
        AryReverse = tmp
        Exit Function
    End If
    
    For i = UBound(ary) To LBound(ary) Step -1
        'Debug.Print i & " " & Hex(ary(i))
        push tmp, ary(i)
    Next
    
    AryReverse = tmp
End Function

'python compatability
Function in_range(v, min As Long, max As Long) As Boolean
    If v >= min And v <= max Then in_range = True
End Function


Function keyExists(ByVal t As String, c As Collection) As Boolean
    On Error GoTo hell
    Dim X
    
    If IsObject(c(t)) Then
        Set X = c(t)
    Else
        X = c(t)
    End If
    
    keyExists = True
Exit Function
hell:
End Function

Function anyOfTheseInstr(data, match, Optional compare As VbCompareMethod = vbTextCompare) As Boolean
    Dim tmp() As String
    Dim X
    tmp = split(match, ",")
    For Each X In tmp
        If InStr(1, data, X, compare) > 0 Then
            anyOfTheseInstr = True
            Exit Function
        End If
    Next
End Function

Public Function isHexNum(v, Optional ByRef dataOut) As Boolean
    On Error Resume Next
    dataOut = CLng("&h" & v)
    If Err.Number = 0 Then isHexNum = True
End Function

Function Rpad(v, Optional L As Long = 8, Optional char As String = " ")
    On Error GoTo hell
    Dim X As Long
    X = Len(v)
    If X < L Then
        Rpad = v & String(L - X, char)
    Else
hell:
        Rpad = v
    End If
End Function

Function lpad(v, Optional L As Long = 8, Optional char As String = " ")
    On Error GoTo hell
    Dim X As Long
    X = Len(v)
    If X < L Then
        lpad = String(L - X, char) & v
    Else
hell:
        lpad = v
    End If
End Function

Function pad(v, Optional L As Long = 8, Optional char As String = " ", Optional padRight As Boolean = True)
    On Error GoTo hell
    Dim X As Long
    X = Len(v)
    If X < L Then
        If padRight Then
             pad = v & String(L - X, char)
        Else
             pad = String(L - X, char) & v
        End If
    Else
hell:
        pad = v
    End If
End Function

Function printf(ByVal msg As String, ParamArray values()) As String
   
    Dim i As Long, tmp()
    
    For i = 0 To UBound(values)
        If IsNull(values(i)) Then
            push tmp, "[Null]"
        ElseIf IsObject(values(i)) Then
            push tmp, "[Object:" & TypeName(values(i)) & "]"
        Else
            push tmp, values(i)
        End If
    Next
    
    printf = modGeneral.printf(msg, tmp)

End Function

Function ado_BuildConnectionString(dbServer As dbServers, dbName As String, Optional server As String, _
                          Optional Port = 3306, Optional User As String, Optional pass As String) As String
    
    Dim dbPath As String
    Dim baseString As String
    Dim blnInlineAuth As Boolean
    
    Select Case dbServer
        Case db_Access
            baseString = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=____;"
        Case db_FileDsn
            baseString = "FILEDSN=____;"
        Case db_DSN
            baseString = "DSN=____;"
        Case db_dBase
            baseString = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=____;"
        Case db_mysql
            baseString = "Driver={mySQL};Server=" & server & ";Port=" & Port & ";Stmt=;Option=16834;Database=____;User=" & User & ";Password=" & pass & ";"
        Case db_mysql5
            baseString = "Driver={MySQL ODBC 5.1 Driver};Server=" & server & ";Port=" & Port & ";Database=____;User=" & User & ";Password=" & pass & ";Option=3;"
        Case db_MsSql2k
            baseString = "Driver={SQL Server};Server=" & server & ";Database=____;"
        Case db_JetAccess2k
            baseString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=____;" & _
                         "User Id=" & User & ";" & _
                         "Password=" & pass & ";"
                         blnInlineAuth = True
    End Select
                         
        
    If Not blnInlineAuth Then
        If User <> Empty Then baseString = baseString & "Uid:" & User & ";"
        If pass <> Empty Then baseString = baseString & "Pwd:" & User & ";"
    End If
       
    '%AP% is like enviromental variable for app.path i am lazy :P
    dbPath = Replace(dbName, "%AP%", App.path)
    
    ado_BuildConnectionString = Replace(baseString, "____", dbPath)
    
End Function

Function ado_buildInsert(tblName, fields, ParamArray params()) As String
    Dim values()
    values() = params() 'force byval
    ado_buildInsert = ado_buildInsert2(tblName, fields, values)
End Function

Function ado_buildInsert2(tblName, fields, params()) As String

    Dim sSql As String, sTmp As String, i As Integer, X
    Dim values()
    
    values() = params() 'force byval
    sSql = "Insert into " & tblName & " (" & fields & ") VALUES(____)"
    
    Dim tn As String
    For i = 0 To UBound(values)
        tn = LCase(TypeName(values(i)))
        If tn = "string" Or tn = "textbox" Or tn = "date" Then
            values(i) = "'" & Replace(values(i), "'", "''") & "'"
        End If
    Next
    
    sTmp = Join(values, ",")
    ado_buildInsert2 = Replace(sSql, "____", sTmp)
     
    
End Function


Function ado_buildUpdate(tblName, criteria, fields, ParamArray params()) As String
    Dim values()
    values() = params() 'force byval
    ado_buildUpdate = ado_buildUpdate2(tblName, criteria, fields, values)
End Function

Function ado_buildUpdate2(tblName, criteria, fields, params()) As String

    Dim sSql As String, sTmp As String, i As Integer
    Dim values()
    
    values() = params() 'force byval
    sSql = "Update " & tblName & " SET ____" & " " & criteria
    
    Dim tn As String
    For i = 0 To UBound(values)
        tn = LCase(TypeName(values(i)))
        If tn = "string" Or tn = "textbox" Or tn = "date" Then
            values(i) = "'" & Replace(values(i), "'", "''") & "'"
        End If
    Next
    
    Dim tmp() As String
    tmp() = split(fields, ",")
    
    If UBound(tmp()) <> UBound(values()) Then MsgBox "Err in update fields/values"
    
    For i = 0 To UBound(tmp)
        tmp(i) = tmp(i) & "=" & values(i)
    Next
    
    sTmp = Join(tmp, ",")
    ado_buildUpdate2 = Replace(sSql, "____", sTmp)
    
End Function


''endian swap and return spaced out hex string
'Private Function lng2Hex(x As Long) As String
'    Dim b(1 To 4) As Byte
'    CopyMemory b(1), x, 4
'    lng2Hex = Hex(b(1)) & " " & Hex(b(2)) & " " & Hex(b(3)) & " " & Hex(b(4))
'End Function


Function isPrintable(b As Byte) As Boolean
     
    isPrintable = True
    If b = 9 Or b = 10 Or b = 13 Then
        'tab, cr, lf are ok..
    ElseIf b >= &H20 And b <= &H7E Then
        'its a printable character and is ok..
    Else
        isPrintable = False
    End If
 
End Function

Function countOccurances(it, find) As Integer
    Dim tmp() As String
    
    If InStr(1, it, find, vbTextCompare) < 1 Then Exit Function
    tmp = split(it, find, , vbTextCompare)
    countOccurances = UBound(tmp)
    
End Function

Function h2l(ByVal X, ByRef outVal As Long) As Boolean
    On Error Resume Next
    X = trim(Replace(X, "0x", Empty, , , vbTextCompare))
    X = trim(Replace(X, "&h", Empty, , , vbTextCompare))
    outVal = CLng("&h" & X)
    h2l = (Err.Number = 0)
End Function

Function HexDump(bAryOrStrData, Optional ByVal Length As Long = -1, Optional ByVal startAt As Long = 1, Optional hexFormat As hexOutFormats = hoDump) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim i As Long, tt, h, X
    Dim hexOnly As Long
    
    offset = 0
    If hexFormat <> hoDump Then hexOnly = 1
    
    If TypeName(bAryOrStrData) = "Byte()" Then
        If AryIsEmpty(bAryOrStrData) Then Exit Function
        ary() = bAryOrStrData
    Else
        If Len(CStr(bAryOrStrData)) = 0 Then Exit Function
        ary = StrConv(CStr(bAryOrStrData), vbFromUnicode, LANG_US)
    End If
    
    If startAt < 1 Then startAt = 1
    If Length < 1 Then Length = -1
    
    While startAt Mod 16 <> 0
        startAt = startAt - 1
    Wend
    
    startAt = startAt + 1
    
    chars = "   "
    For i = startAt To UBound(ary) + 1
        tt = Hex(ary(i - 1))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        X = ary(i - 1)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((X > 32 And X < 127), Chr(X), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
        If Length <> -1 Then
            Length = Length - 1
            If Length = 0 Then Exit For
        End If
    Next
    
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        If hexFormat = hoHexOnly Then HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function

Public Function toBytes(ByVal hexstr, ByRef outVar, Optional ByVal inputformat As InputFormats = ifHex) As Boolean

'supports:
'11 22 33 44   spaced hex chars
'11223344      run together hex strings
'11,22,33,44   csv hex
'1,2,3,4       csv hex with no lead 0
'121,99,44,255 decimal csv or spaced values
'%xx%yy
'%uxxxx\u7766
'%u6162%63
'isDecimal flag requires csv or spaced values..
'ignores common C source prefixes and characters

    Dim ret As String, X As String, str As String
    Dim R() As Byte, b As Byte, b1 As Byte
    Dim foundDecimal As Boolean, tmp, i, a, a2
    Dim pos As Long, marker As String
    
    On Error GoTo hell
    
    If inputformat = ifHexDump Then
        X = ExtractHexFromDump(hexstr)      'returns just the hex string
        If Not toBytes(X, R) Then GoTo hell 'now we convert it to actual bytes..
        GoTo retNow
    End If
    
    If inputformat = ifString Then
        R() = StrConv(hexstr, vbFromUnicode, LANG_US)
        GoTo retNow
    End If
    
    str = Replace(hexstr, vbCr, Empty)
    str = Replace(str, vbLf, Empty)
    str = Replace(str, vbTab, Empty)
    str = Replace(str, Chr(0), Empty)
    str = Replace(str, "{", Empty)
    str = Replace(str, "}", Empty)
    str = Replace(str, ";", Empty)
    str = Replace(str, "+", Empty)
    str = Replace(str, """""", Empty)
    str = Replace(str, "'", Empty)
    If inputformat = ifDecimal Then str = Replace(str, " ", Empty)
    hexstr = str
    
    If InStr(hexstr, "\u") > 0 Then hexstr = Replace(hexstr, "\u", "%u")
    
    If InStr(hexstr, "%u") > 0 Then
        tmp = split(hexstr, "%u")
        For i = 1 To UBound(tmp)
            a = InStr(tmp(i), "%")
            X = ""
            If a > 1 Then
                X = Mid(tmp(i), a)
                tmp(i) = Mid(tmp(i), 1, a - 1)
            End If
            If Len(tmp(i)) = 3 Then tmp(i) = "0" & tmp(i)
            If Len(tmp(i)) = 4 Then
                a = Mid(tmp(i), 1, 2)
                a2 = Mid(tmp(i), 3, 2)
                tmp(i) = a2 & a
            End If
            If Len(X) > 0 Then tmp(i) = tmp(i) & X
        Next
        hexstr = Join(tmp, "")
    End If
    
    If InStr(hexstr, "%") > 0 Then
        tmp = split(hexstr, "%")
        For i = 1 To UBound(tmp)
            If Len(tmp(i)) < 2 Then
                tmp(i) = 0 & tmp(i)
            End If
        Next
        hexstr = Join(tmp, "")
    End If
    
    If Len(hexstr) > 4 Then
        pos = FirstOccurance(hexstr, " ,", marker)
        If pos > 0 And pos < 5 Then   'make sure all are double digit hex chars...(also account for decimal 1,11,111,
            tmp = split(hexstr, marker)
            
            If inputformat = ifDecimal Then
                For i = 0 To UBound(tmp)
                    tmp(i) = Hex(CLng(tmp(i)))
                Next
            End If
            
            For i = 0 To UBound(tmp)
                If Len(tmp(i)) = 1 Then tmp(i) = "0" & tmp(i)
            Next
            
            hexstr = Join(tmp, "")
        End If
    End If
        
    str = Replace(hexstr, " ", Empty)
    str = Replace(str, "0x", Empty)
    str = Replace(str, ",", Empty)
    
    For i = 1 To Len(str) Step 2
        X = Mid(str, i, 2)
        If Not isHexChar(X, b) Then Exit Function
        bpush R(), b
    Next
    
retNow:
    If TypeName(outVar) = "Byte()" Then
        outVar = R
    Else
        outVar = StrConv(R, vbUnicode, LANG_US)
    End If
    
    toBytes = True
    Exit Function
    
hell:
    toBytes = False
    
End Function

Function FirstOccurance(it, ByVal csvFind As String, ByRef outFoundVal) As Long
    If Len(csvFind) = 0 Then Exit Function
    
    Dim find() As String, X, lowestOffset As Long, lowestIndex As Long, i As Long, a As Long
    
    outFoundVal = Empty
    lowestOffset = MAX_LONG
    find = split(csvFind, ",")
    
    For i = 0 To UBound(find)
        If Len(find(i)) = 0 Then find(i) = ","
        a = InStr(1, it, find(i), vbTextCompare)
        If a > 0 And a < lowestOffset Then
            lowestOffset = a
            lowestIndex = i
        End If
    Next
    
    If lowestOffset = MAX_LONG Then Exit Function
    
    outFoundVal = find(lowestIndex)
    FirstOccurance = lowestOffset
    
End Function

Private Sub bpush(bAry() As Byte, b As Byte) 'this modifies parent ary object
    On Error GoTo Init
    Dim X As Long
    
    X = UBound(bAry) '<-throws Error If Not initalized
    ReDim Preserve bAry(UBound(bAry) + 1)
    bAry(UBound(bAry)) = b
    
    Exit Sub

Init:
    ReDim bAry(0)
    bAry(0) = b
    
End Sub

Public Function isHexChar(hexValue As String, Optional b As Byte) As Boolean
    On Error Resume Next
    Dim v As Long
    
    If Len(hexValue) = 0 Then GoTo nope
    If Len(hexValue) > 2 Then GoTo nope 'expecting hex char code like FF or 90
    
    v = CLng("&h" & hexValue)
    If Err.Number <> 0 Then GoTo nope 'invalid hex code
    
    b = CByte(v)
    If Err.Number <> 0 Then GoTo nope  'shouldnt happen.. > 255 cant be with len() <=2 ?

    isHexChar = True
    
    Exit Function
nope:
    Err.Clear
    isHexChar = False
End Function

Public Function xorDecode(fileByteOrString As Variant, ByVal password As Variant, Optional strRet As Boolean = True)
    Dim key() As Byte, ByteArray() As Byte, ko As Long, i As Long
    
    Dim plen As Long
    
    ByteArray() = LoadData(fileByteOrString)
    
    If TypeName(password) = "Byte()" Then
        key() = password
    ElseIf TypeName(password) = "Byte" Then
        ReDim key(0)
        key(0) = password
    Else
        If Len(password) = 0 Then Exit Function
         key() = StrConv(CStr(password), vbFromUnicode, LANG_US)
    End If
    
    For i = 0 To UBound(ByteArray)
        ko = i Mod (UBound(key) + 1)
        ByteArray(i) = ByteArray(i) Xor key(ko)
    Next
    
    If strRet Then
        xorDecode = StrConv(ByteArray, vbUnicode, LANG_US)
    Else
        xorDecode = ByteArray
    End If
    
End Function


Public Function rc4(fileByteOrString As Variant, ByVal password As Variant, Optional strRet As Boolean = True)
On Error Resume Next
    Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, key() As Byte, ByteArray() As Byte, temp As Byte
    
    Dim plen As Long
    
    ByteArray() = LoadData(fileByteOrString)
    
    If TypeName(password) = "Byte()" Then
        key() = password
        If UBound(key) > 255 Then ReDim Preserve key(255)
    Else
        If Len(password) = 0 Then
            Exit Function
        End If

        If Len(password) > 256 Then
            key() = StrConv(Left$(CStr(password), 256), vbFromUnicode, LANG_US)
        Else
            key() = StrConv(CStr(password), vbFromUnicode, LANG_US)
        End If
    End If
    
    plen = UBound(key) + 1
 
    'Debug.Print "key=" & HexDump(Key)
    'Debug.Print "data=" & HexDump(ByteArray)
    
    For X = 0 To 255
        RB(X) = X
    Next X
    
    X = 0
    Y = 0
    Z = 0
    For X = 0 To 255
        Y = (Y + RB(X) + key(X Mod plen)) Mod 256
        temp = RB(X)
        RB(X) = RB(Y)
        RB(Y) = temp
    Next X
    
    X = 0
    Y = 0
    Z = 0
    For X = 0 To UBound(ByteArray)
        Y = (Y + 1) Mod 256
        Z = (Z + RB(Y)) Mod 256
        temp = RB(Y)
        RB(Y) = RB(Z)
        RB(Z) = temp
        ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
    Next X
    
    If strRet Then
        rc4 = StrConv(ByteArray, vbUnicode, LANG_US)
    Else
        rc4 = ByteArray
    End If
    
End Function

Function b64Encode(bArrayStringOrFile, Optional retAsString As Boolean = True) As Variant
    b64Encode = base64.Encode(bArrayStringOrFile, retAsString)
End Function

Function b64Decode(bArrayStringOrFile, ByRef dataOut As Variant, Optional retAsString As Boolean = True) As Boolean
     b64Decode = base64.Decode(bArrayStringOrFile, dataOut, retAsString)
End Function

Function entropy(fileByteOrString, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    If FileExists(fileByteOrString) Then
        entropy = fileEntropy(CStr(fileByteOrString), offset, leng)
        Exit Function
    End If
    
    Dim b() As Byte
    b() = LoadData(fileByteOrString, False)
    entropy = memEntropy(b, offset, leng)
        
End Function

Function crc32(fileByteOrString) As String

    Dim b() As Byte
    Dim v As Long
    
    If Not ensureUTypes() Then Exit Function 'use the C dll version its faster
    
    b() = LoadData(fileByteOrString)
    

    If AryIsEmpty(b) Then Exit Function
    
    v = ut_crc32(b(0), UBound(b) + 1)
    crc32 = Hex(v)
    
'    vb native implementation...
'    Dim c As Long, n As Long, x As Long
'    Dim b() As Byte
'
'    c = -1
'    If AryIsEmpty(crc_table) Then make_crc_table
'
'    If TypeName(bAryOrString) = "Byte()" Then
'        b() = bAryOrString
'    Else
'        b() = StrConv(CStr(bAryOrString), vbFromUnicode, LANG_US)
'    End If
'
'    For n = 0 To UBound(b)
'        c = crc_table((c Xor b(n)) And &HFF) Xor rshift(c, 8)
'    Next
'
'    crc32 = Hex(c Xor &HFFFFFFFF)
    
End Function

'Private Sub make_crc_table()
'    Dim c As Long, n As Long, k As Long
'
'    ReDim crc_table(256)
'
'    For n = 0 To 255
'          c = n
'          For k = 0 To 7
'                If c And 1 Then
'                     c = &HEDB88320 Xor rshift(c)
'                Else
'                    c = rshift(c)
'                End If
'           Next
'          crc_table(n) = c
'    Next
'
'End Sub

'Public Function lshift(ByVal Value As Long, Optional ByVal Shift As Integer = 1) As Long
'    MakeOnBits
'    If (Value And (2 ^ (31 - Shift))) Then 'GoTo OverFlow
'        lshift = ((Value And OnBits(31 - (Shift + 1))) * (2 ^ (Shift))) Or &H80000000
'    Else
'        lshift = ((Value And OnBits(31 - Shift)) * (2 ^ Shift))
'    End If
'End Function
'
'Public Function rshift(ByVal Value As Long, Optional ByVal Shift As Integer = 1) As Long
'    Dim hi As Long
'    MakeOnBits
'    If (Value And &H80000000) Then hi = &H40000000
'    rshift = (Value And &H7FFFFFFE) \ (2 ^ Shift)
'    rshift = (rshift Or (hi \ (2 ^ (Shift - 1))))
'End Function
'
'Private Sub MakeOnBits()
'    Dim j As Integer, v As Long
'
'    For j = 0 To 30
'        v = v + (2 ^ j)
'        OnBits(j) = v
'    Next j
'
'    OnBits(j) = v + &H80000000
'
'End Sub

Public Function inc(ByRef X, Optional Increment As Long = 1)
    X = X + Increment
End Function

Public Function dec(ByRef X, Optional Increment As Long = 1)
    X = X - Increment
End Function

Function paramsToArray(ParamArray values()) As Variant()
    'paramsToArray = values 'creates multidim array?
    
    Dim i As Long, tmp()
    
    For i = 0 To UBound(values)
        If IsNull(values(i)) Then
            push tmp, "[Null]"
        ElseIf IsObject(values(i)) Then
            push tmp, "[Object:" & TypeName(values(i)) & "]"
        Else
            push tmp, values(i)
        End If
    Next
    
    paramsToArray = tmp
    
End Function

Function max(values() As Variant)
   On Error Resume Next
   Dim Item
   For Each Item In values
      max = IIf(max < Item, Item, max)
   Next
End Function

Function min(values() As Variant)
   Dim Item
   On Error Resume Next
   min = &H7FFFFFFF
   For Each Item In values
      min = IIf(min > Item, Item, min)
   Next
End Function

Function md5(fileByteOrString) As String
    Dim b() As Byte
    b() = LoadData(fileByteOrString)
    md5 = hash.HashBytes(b)
End Function

Function ExtractHexFromDump(dump) As String
    
    On Error Resume Next
    Dim Y() As String, tmp() As String, i As Long
    Dim dat As String
    Dim a As Long
    Dim X
    
    dat = trim(dump)
    If InStr(dat, " ") > 0 Then
        tmp = split(dat, vbCrLf)
        For i = 0 To UBound(tmp)
            X = trim(tmp(i))
            If InStr(1, X, "Offset", vbTextCompare) = 1 Then GoTo nextone
            If Len(X) = 0 Then GoTo nextone
            a = InStr(X, " ") '8
            If a > 0 Then 'remove offset
                X = trim(Mid(X, a))
            End If
            
            a = Len(X)
            If a >= 48 Then
                X = trim(Mid(X, 1, 48))
            Else
                'they didnt copy the ascii area for last line? it will be copied anyway...
                'If i = UBound(tmp) Then
            End If
            
            
            push Y, Replace(X, " ", Empty)
nextone:
        Next
        dat = Join(Y, "")
    End If
    
    ExtractHexFromDump = dat
     
End Function

Function Compress(fileStringOrByte, ByRef outVal, Optional method As CompressTypes = ct_Zlib) As Boolean
    
    Dim b() As Byte
    Dim bOut() As Byte
    Dim success As Boolean
    Dim f As Long
    
    b() = LoadData(fileStringOrByte)
    
    Select Case method
        Case ct_Zlib:      success = zlib.CompressData(b, bOut)
        Case ct_LZNT1:     success = RTLCompress(b, bOut)
        Case ct_LZNT1_MAX: success = RTLCompress(b, bOut, True)
    End Select
    
    If success Then
        If TypeName(outVal) = "String" Or TypeName(outVal) = "Variant" Then
           outVal = StrConv(bOut, vbUnicode, LANG_US)
        Else
           outVal = bOut
        End If
    End If
         
    Compress = success
    
End Function

'if outval is a string or variant this will return a string. if its a bytearray() it will return a byte array
Function DeCompress(fileStringOrByte, ByRef outVal, Optional method As CompressTypes = ct_Zlib) As Boolean
    
    Dim b() As Byte
    Dim bOut() As Byte
    Dim f As Long
    Dim success As Boolean
    
    On Error Resume Next
    b() = LoadData(fileStringOrByte)
    If Err.Number <> 0 Then Exit Function
    
    Select Case method
        Case ct_Zlib:  success = zlib.UncompressData(b, bOut)
        Case ct_LZNT1:   success = RTLDeCompress(b, bOut)
        Case ct_LZNT1_MAX:   success = RTLDeCompress(b, bOut, True)
    End Select
    
    If success Then
        If TypeName(outVal) = "String" Or TypeName(outVal) = "Variant" Then
           outVal = StrConv(bOut, vbUnicode, LANG_US)
        Else
           outVal = bOut
        End If
    End If
    
    DeCompress = success
         
End Function

Function DownloadFile(url, localPath) As Boolean
    DownloadFile = (URLDownloadToFile(0&, url, localPath, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS)
End Function

Function a2c(ary) As Collection
    Dim a As New Collection, X
    For Each X In ary
        a.add X
    Next
    Set a2c = a
End Function

Function c2a(c As Collection) As String()
    Dim t() As String, f
    For Each f In c
        push t, f
    Next
    c2a = t
End Function

Function c2s(c As Collection, Optional delimiter = vbCrLf) As String
    c2s = Join(c2a(c), delimiter)
End Function

Function cjoin(a As Collection, b As Collection) As Collection
    Dim ret As New Collection, f
    For Each f In a
        ret.add f
    Next
    For Each f In b
        ret.add b
    Next
    Set cjoin = ret
End Function

Function isIn(needle, haystack, Optional ByRef index) As Boolean
    Dim t As String, J As String, X
    Dim i As Long
    
    On Error GoTo hell
    'string or variant are treated as strings and index = instr
    'arrays will walk all elements, index will be found array index, search as above
    'collections as above index col index, also supports object types..
    'other datatypes
    index = -1
    t = TypeName(haystack)
    
    If t = "String" Or t = "Variant" Then
        index = InStr(1, haystack, needle, vbTextCompare)
        If index > 0 Then isIn = True
    ElseIf InStr(t, "()") > 0 Then 'its an array
        For i = LBound(haystack) To UBound(haystack)
            J = TypeName(haystack(i))
            If J = "String" Or t = "Variant" Then
                 index = InStr(1, haystack(i), needle, vbTextCompare)
                 If index > 0 Then
                    isIn = True
                    index = i
                    Exit For
                 End If
            Else
                If needle = haystack(i) Then
                    isIn = True
                    index = i
                    Exit For
                End If
            End If
        Next
    ElseIf t = "Collection" Then
        i = 0
        For Each X In haystack
            i = i + 1
            J = TypeName(X)
            If IsObject(needle) And IsObject(X) Then
                 If ObjPtr(needle) = ObjPtr(X) Then
                    isIn = True
                    index = i
                    Exit For
                 End If
            ElseIf J = "String" Or t = "Variant" Then
                 index = InStr(1, haystack(i), needle, vbTextCompare)
                 If index > 0 Then
                    isIn = True
                    index = i
                    Exit For
                 End If
            Else
                If needle = X Then
                    isIn = True
                    index = i
                    Exit For
                End If
            End If
        Next
    Else
        If needle = haystack Then
            isIn = True
            index = 0
        End If
    End If
                    
    Exit Function
    
hell:
    isIn = False
    index = -2
    
End Function

Function concat(divider, ParamArray elements()) As String
    Dim i As Long
    For i = 0 To UBound(elements)
        concat = concat & elements(i) & IIf(i <> UBound(elements), divider, Empty)
    Next
End Function

Public Function url_decode(ByVal url As String, ByRef outVal As String, Optional ByVal PlusSpace As Boolean = True) As Boolean
    
    Dim cchUnescaped As Long
    Dim HRESULT As Long
    Dim tmp As String
    
    outVal = Empty
    
    If PlusSpace Then url = Replace$(url, "+", " ")
    cchUnescaped = Len(url)
    tmp = String$(cchUnescaped, 0)
    HRESULT = UrlUnescape(url, tmp, cchUnescaped, 0)
    
    If HRESULT = E_POINTER Then
        tmp = String$(cchUnescaped, 0)
        HRESULT = UrlUnescape(url, tmp, cchUnescaped, 0)
    End If
    
    If HRESULT <> S_OK Then Exit Function
    
    outVal = Left$(tmp, cchUnescaped)
    url_decode = True
    
End Function

Public Function url_encode(ByVal url As String, ByRef outVal As String, Optional ByVal SpacePlus As Boolean = True) As Boolean
    
    Dim cchEscaped As Long
    Dim HRESULT As Long
    Dim tmp As String
    
    outVal = Empty
    
    If Len(url) > INTERNET_MAX_URL_LENGTH Then
         outVal = "URL parameter too long"
         Exit Function
    End If
    
    cchEscaped = Len(url) * 1.5
    tmp = String$(cchEscaped, 0)
    
    HRESULT = UrlEscape(url, tmp, cchEscaped, URL_ESCAPE_PERCENT)
    
    If HRESULT = E_POINTER Then
        tmp = String$(cchEscaped, 0)
        HRESULT = UrlEscape(url, tmp, cchEscaped, URL_ESCAPE_PERCENT)
    End If

    If HRESULT <> S_OK Then Exit Function
    
    tmp = Left$(tmp, cchEscaped)
    If SpacePlus Then
        tmp = Replace$(tmp, "+", "%2B")
        tmp = Replace$(tmp, " ", "+")
    End If
    
    outVal = tmp
    url_encode = True
    
End Function

Public Function html_encode(ByVal iString As String, ByRef outVal) As Boolean
    Dim o As Object
    
    outVal = Empty
    Set o = CreateObject("MSXML2.DOMDocument")
    
    If o Is Nothing Then Exit Function

    With o.createTextNode(iString)
        outVal = .xml
        If outVal <> iString Then html_encode = True
    End With
    
End Function

Public Function html_decode(ByVal iString As String, ByRef outVal) As Boolean
    Dim o As Object
    
    outVal = Empty
    Set o = CreateObject("MSXML2.DOMDocument")
    
    If o Is Nothing Then Exit Function
    
    With o
        html_decode = .loadXML("<p>" & iString & "</p>")
        If html_decode Then
            outVal = .selectSingleNode("p").nodeTypedValue
        End If
    End With
    
End Function

Function cast(varIn, varOut, Optional allowFilePaths As Boolean = False) As Boolean
    On Error Resume Next
    Dim b() As Byte
    Dim i() As Integer
    Dim L() As Long
    Dim c() As Currency
    Dim d() As Double
    
    Dim sz As Long
    Dim Bytes As Long
    
    On Error GoTo hell
    
    If TypeName(varIn) = "Object" Or TypeName(varOut) = "Object" Then
        Err.Raise "Invalid use of cast object type not supported"
    End If
    
    If Not allowFilePaths Then
        If TypeName(varIn) = TypeName(varOut) Then
            varOut = varIn
            Exit Function
        End If
    End If
        
    b() = LoadData(varIn, allowFilePaths)
    Bytes = UBound(b) + 1
    
    If TypeName(varOut) = "String" Or TypeName(varOut) = "Variant" Then
        varOut = StrConv(b, vbUnicode, LANG_US)
    ElseIf TypeName(varOut) = "Byte()" Then
        varOut = b()
    ElseIf TypeName(varOut) = "Integer()" Then
        If Bytes Mod 2 <> 0 Then
            sz = Bytes + 2 - (Bytes Mod 2)
            ReDim Preserve b(sz)
            Bytes = sz
        End If
        ReDim i((Bytes / 2) - 1)
        CopyMemory ByVal VarPtr(i(0)), ByVal VarPtr(b(0)), Bytes
        varOut = i()
    ElseIf TypeName(varOut) = "Long()" Then
        If Bytes Mod 4 <> 0 Then
            sz = Bytes + 4 - (Bytes Mod 4)
            ReDim Preserve b(sz)
            Bytes = sz
        End If
        ReDim L((Bytes / 4) - 1)
        CopyMemory ByVal VarPtr(L(0)), ByVal VarPtr(b(0)), Bytes
        varOut = L()
    ElseIf TypeName(varOut) = "Double()" Then
        If Bytes Mod 8 <> 0 Then
            sz = Bytes + 8 - (Bytes Mod 8)
            ReDim Preserve b(sz)
            Bytes = sz
        End If
        ReDim d((Bytes / 8) - 1)
        CopyMemory ByVal VarPtr(d(0)), ByVal VarPtr(b(0)), Bytes
        varOut = d()
    ElseIf TypeName(varOut) = "Currency()" Then
        If Bytes Mod 8 <> 0 Then
            sz = Bytes + 8 - (Bytes Mod 8)
            ReDim Preserve b(sz)
            Bytes = sz
        End If
        ReDim c((Bytes / 8) - 1)
        CopyMemory ByVal VarPtr(c(0)), ByVal VarPtr(b(0)), Bytes
        varOut = c()
    Else
        varOut = Empty
        cast = False
        Exit Function
    End If
    
    cast = True
    
Exit Function
hell:
    cast = False

End Function


Function LngToBytes(s As String, ByRef out_b() As Byte, Optional intOnly As Boolean) As Boolean
        On Error Resume Next
        
        Dim tmp, L As Long, b() As Byte, hadErr As Boolean, i As Integer
        tmp = s

        If Right(LCase(tmp), 1) = "d" And Len(tmp) > 1 Then
            tmp = Mid(tmp, 1, Len(tmp) - 1)
            L = CLng(tmp)
            If intOnly Then i = CInt(L)
            If Err.Number > 0 Then Exit Function
        Else
            L = HexStrToLng(tmp, hadErr)
            If intOnly Then i = CInt(L)
            If hadErr Or Err.Number > 0 Then Exit Function
        End If
        
        If intOnly Then
            ReDim b(1)
            CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(i), 2
        Else
            ReDim b(3)
            CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(L), 4
        End If
  
        out_b = b
        LngToBytes = True
        
End Function

Function SngToBytes(tmp As String, ByRef out_b() As Byte) As Boolean
       On Error Resume Next
        
       Dim s As Single, b() As Byte
       
       s = tmp
       If Err.Number <> 0 Then Exit Function
    
       ReDim b(LenB(s) - 1)
       CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(s), UBound(b) + 1
         
       out_b = b
       SngToBytes = True
        
End Function

Function CurToBytes(tmp As String, ByRef out_b() As Byte) As Boolean
       On Error Resume Next
        
       Dim s As Currency, b() As Byte
       
       s = tmp
       If Err.Number <> 0 Then Exit Function
    
       ReDim b(LenB(s) - 1)
       CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(s), UBound(b) + 1
         
       out_b = b
       CurToBytes = True
        
End Function

Function DblToBytes(tmp As String, ByRef out_b() As Byte) As Boolean
       On Error Resume Next
        
       Dim s As Double, b() As Byte
       
       s = tmp
       If Err.Number <> 0 Then Exit Function
    
       ReDim b(LenB(s) - 1)
       CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(s), UBound(b) + 1
         
       out_b = b
       DblToBytes = True
        
End Function

Function DateToBytes(tmp As String, ByRef out_b() As Byte) As Boolean
       On Error Resume Next
        
       Dim s As Date, b() As Byte
       
       s = tmp
       If Err.Number <> 0 Then Exit Function
    
       ReDim b(LenB(s) - 1)
       CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(s), UBound(b) + 1
         
       out_b = b
       DateToBytes = True
        
End Function


Function bitTest(ByVal val As Long, bit) As Boolean
    If AryIsEmpty(bitPatterns) Then
        bitPatterns = Array(0, &H1, &H2, &H4, &H8, &H10, &H20, &H40, &H80, &H100, &H200, &H400, &H800, &H1000, &H2000, &H4000, &H8000, &H10000, &H20000, &H40000, &H80000, &H100000, &H200000, &H400000, &H800000, &H1000000, &H2000000, &H4000000, &H8000000, &H10000000, &H20000000, &H40000000, &H80000000)
    End If
    If bit > 32 Or bit < 1 Then Err.Raise "invalid bit number 1-32 only"
    bitTest = ((val And bitPatterns(bit)) > 0)
End Function

Function toBinary(ByVal val As Long, Optional maxBits As Long = 32) As String
    
    Dim ret As String, J, i
    
    If maxBits < 1 Or maxBits > 32 Then maxBits = 32
    
    For i = maxBits To 1 Step -1
        J = J + 1
        ret = ret & IIf(bitTest(val, i), "1", "0")
        If J Mod 8 = 0 Then ret = ret & " "
        If J Mod 16 = 0 Then ret = ret & " "
    Next
    
    toBinary = ret
    
End Function
