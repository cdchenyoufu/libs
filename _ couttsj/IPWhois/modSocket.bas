Attribute VB_Name = "modSocket"
Option Explicit

Private Const WM_APP            As Long = 32768 '0x8000
Public Const RESOLVE_MESSAGE    As Long = WM_APP
Public Const SOCKET_MESSAGE     As Long = WM_APP + 1
Private Const PATCH_09          As Long = 119
Private Const PATCH_0C          As Long = 150
Private Const MEM_RELEASE       As Long = &H8000&   'Release allocated memory flag
Private Const MEM_COMMIT        As Long = &H1000&   'Commit allocated memory
Private Const PAGE_RWX          As Long = &H40&     'Allocate executable memory
Public Const SOCKET_ERROR       As Long = -1

Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Private Const GWL_WNDPROC As Long = (-4)
Private Const TIMER_TIMEOUT As Long = 200   'control timer time out, in milliseconds

Private lMsgCntA        As Long     'TableA entry count
Private lMsgCntB        As Long     'TableB entry count
Private lTableA1()      As Long     'TableA1: list of async handles
Private lTableA2()      As Long     'TableA2: list of async handles owners
Private lTableB1()      As Long     'TableB1: list of sockets
Private lTableB2()      As Long     'TableB2: list of sockets owners
Private hWndSub         As Long     'window handle subclassed
Private nAddrSubclass   As Long     'address of our WndProc
Private nAddrOriginal   As Long     'address of original WndProc
Private hTimer          As Long     'control timer handle
Public DbgFlg           As Boolean  'Determines if debug messages are recorded

'==============================================================================
'MEMBER VARIABLES
'==============================================================================
Private m_bInit                 As Boolean      'SubClass initialized
Private m_lSocksQuantity        As Long         'number of instances created
Private m_colSocketsInst        As Collection   'sockets list and instance owner
Private m_colAcceptList         As Collection   'sockets in queue that need to be accepted
Private m_hWindow               As Long         'message window handle

'WINSOCK CONTROL ERROR CODES
Public Const sckOutOfMemory = 7

' To initialize Winsock.
Private Type WSADATA
   wVersion                               As Integer
   wHighVersion                           As Integer
   szDescription(256 + 1)                 As Byte
   szSystemstatus(128 + 1)                As Byte
   iMaxSockets                            As Integer
   iMaxUpdDg                              As Integer
   lpVendorInfo                           As Long
End Type

' Messages send with WSAAsyncSelect().
Public Const FD_READ       As Long = &H1
Public Const FD_WRITE      As Long = &H2
Public Const FD_OOB        As Long = &H4
Public Const FD_ACCEPT     As Long = &H8
Public Const FD_CONNECT    As Long = &H10
Public Const FD_CLOSE      As Long = &H20

'==============================================================================
'SUBCLASSING DECLARATIONS
'by Paul Caton
'==============================================================================
Private Declare Function API_IsWindow Lib "user32" Alias "IsWindow" (ByVal hWnd As Long) As Long
Private Declare Function API_GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function API_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function API_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function API_GetProcAddress Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function API_DestroyWindow Lib "user32" Alias "DestroyWindow" (ByVal hWnd As Long) As Long

' DLL handling functions.
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

' Winsock error constants.
Private Const WSABASEERR          As Long = 10000
Private Const WSAEINTR            As Long = WSABASEERR + 4
Private Const WSAEBADF            As Long = WSABASEERR + 9
Private Const WSAEACCES           As Long = WSABASEERR + 13
Private Const WSAEFAULT           As Long = WSABASEERR + 14
Private Const WSAEINVAL           As Long = WSABASEERR + 22
Private Const WSAEMFILE           As Long = WSABASEERR + 24
Public Const WSAEWOULDBLOCK       As Long = WSABASEERR + 35
Private Const WSAEINPROGRESS      As Long = WSABASEERR + 36
Private Const WSAEALREADY         As Long = WSABASEERR + 37
Private Const WSAENOTSOCK         As Long = WSABASEERR + 38
Private Const WSAEDESTADDRREQ     As Long = WSABASEERR + 39
Public Const WSAEMSGSIZE          As Long = WSABASEERR + 40
Private Const WSAEPROTOTYPE       As Long = WSABASEERR + 41
Private Const WSAENOPROTOOPT      As Long = WSABASEERR + 42
Private Const WSAEPROTONOSUPPORT  As Long = WSABASEERR + 43
Private Const WSAESOCKTNOSUPPORT  As Long = WSABASEERR + 44
Private Const WSAEOPNOTSUPP       As Long = WSABASEERR + 45
Private Const WSAEPFNOSUPPORT     As Long = WSABASEERR + 46
Private Const WSAEAFNOSUPPORT     As Long = WSABASEERR + 47
Private Const WSAEADDRINUSE       As Long = WSABASEERR + 48
Public Const WSAEADDRNOTAVAIL     As Long = WSABASEERR + 49
Private Const WSAENETDOWN         As Long = WSABASEERR + 50
Private Const WSAENETUNREACH      As Long = WSABASEERR + 51
Private Const WSAENETRESET        As Long = WSABASEERR + 52
Private Const WSAECONNABORTED     As Long = WSABASEERR + 53
Private Const WSAECONNRESET       As Long = WSABASEERR + 54
Private Const WSAENOBUFS          As Long = WSABASEERR + 55
Private Const WSAEISCONN          As Long = WSABASEERR + 56
Private Const WSAENOTCONN         As Long = WSABASEERR + 57
Private Const WSAESHUTDOWN        As Long = WSABASEERR + 58
Private Const WSAETOOMANYREFS     As Long = WSABASEERR + 59
Private Const WSAETIMEDOUT        As Long = WSABASEERR + 60
Private Const WSAECONNREFUSED     As Long = WSABASEERR + 61
Private Const WSAELOOP            As Long = WSABASEERR + 62
Private Const WSAENAMETOOLONG     As Long = WSABASEERR + 63
Private Const WSAEHOSTDOWN        As Long = WSABASEERR + 64
Private Const WSAEHOSTUNREACH     As Long = WSABASEERR + 65
Private Const WSAENOTEMPTY        As Long = WSABASEERR + 66
Private Const WSAEPROCLIM         As Long = WSABASEERR + 67
Private Const WSAEUSERS           As Long = WSABASEERR + 68
Private Const WSAEDQUOT           As Long = WSABASEERR + 69
Private Const WSAESTALE           As Long = WSABASEERR + 70
Private Const WSAEREMOTE          As Long = WSABASEERR + 71
Private Const WSASYSNOTREADY      As Long = WSABASEERR + 91
Private Const WSAVERNOTSUPPORTED  As Long = WSABASEERR + 92
Private Const WSANOTINITIALISED   As Long = WSABASEERR + 93
Private Const WSAHOST_NOT_FOUND   As Long = WSABASEERR + 1001
Private Const WSATRY_AGAIN        As Long = WSABASEERR + 1002
Private Const WSANO_RECOVERY      As Long = WSABASEERR + 1003
Private Const WSANO_DATA          As Long = WSABASEERR + 1004

' Other general Win32 APIs.
Public Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (pDestination As Any, ByVal lByteCount As Long)
Private Declare Function API_LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function API_SetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function API_KillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Function AddByte(bArray1() As Byte, bArray2() As Byte, Optional lLen As Long) As Boolean
    Dim lLen1 As Long
    Dim lLen2 As Long
    lLen1 = GetbSize(bArray1)
    lLen2 = GetbSize(bArray2)
    If lLen2 = 0 Then GoTo Done
    If lLen > 0 Then
        If lLen2 > lLen Then
            lLen2 = lLen
        End If
    End If
    ReDim Preserve bArray1(lLen1 + lLen2 - 1)
    CopyMemory bArray1(lLen1), bArray2(0), lLen2
Done:
    AddByte = True
End Function

Public Function ByteToStr(bArray() As Byte) As String
    Dim lPntr As Long
    Dim bTmp() As Byte
    On Error GoTo ByteErr
    ReDim bTmp(UBound(bArray) * 2 + 1)
    For lPntr = 0 To UBound(bArray)
        bTmp(lPntr * 2) = bArray(lPntr)
    Next lPntr
    Let ByteToStr = bTmp
    Exit Function
ByteErr:
    ByteToStr = ""
End Function


Private Function CreateWinsockMessageWindow() As Long
    'Create a window that is used to capture sockets messages.
    'Returns 0 if it has success.
    m_hWindow = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    If m_hWindow = 0 Then
        CreateWinsockMessageWindow = sckOutOfMemory
        Exit Function
    Else
        CreateWinsockMessageWindow = 0
        Call PrintDebug("OK Created winsock message window " & CStr(m_hWindow))
    End If
End Function

Public Function DeleteByte(bArray1() As Byte, ByVal lLen As Long) As Boolean
    Dim lLen1 As Long
    Dim bTmp() As Byte
    lLen1 = GetbSize(bArray1)
    If lLen1 > lLen Then
        ReDim bTmp(lLen1 - lLen - 1)
        CopyMemory bTmp(0), bArray1(lLen - 1), lLen1 - lLen
    End If
    bArray1 = bTmp
End Function

Private Function DestroyWinsockMessageWindow() As Long
    'Destroy the window that is used to capture sockets messages.
    'Returns 0 if it has success.
    Dim lRet As Long
    DestroyWinsockMessageWindow = 0
    If m_hWindow = 0 Then
        Call PrintDebug("WARNING hWindow is ZERO")
        Exit Function
    End If
    lRet = API_DestroyWindow(m_hWindow)
    If lRet = 0 Then
        DestroyWinsockMessageWindow = sckOutOfMemory
        Err.Raise sckOutOfMemory, "modSimple.DestroyWinsockMessageWindow", "Out of memory"
    Else
        Call PrintDebug("OK Destroyed winsock message window " & CStr(m_hWindow))
        m_hWindow = 0
    End If
End Function

Private Function FileErrors(errVal As Integer) As Integer
'Return Value 0=Resume,              1=Resume Next,
'             2=Unrecoverable Error, 3=Unrecognized Error
Dim msgType%
Dim Msg$
Dim Response%
msgType% = 48
Select Case errVal
    Case 68
      Msg$ = "That device appears Unavailable."
      msgType% = msgType% + 4
    Case 71
      Msg$ = "Insert a Disk in the Drive"
    Case 53
      Msg$ = "Cannot Find File"
      msgType% = msgType% + 5
   Case 57
      Msg$ = "Internal Disk Error."
      msgType% = msgType% + 4
    Case 61
      Msg$ = "Disk is Full.  Continue?"
      msgType% = 35
    Case 64, 52
      Msg$ = "That Filename is Illegal!"
    Case 70
      Msg$ = "File in use by another user!"
      msgType% = msgType% + 5
    Case 76
      Msg$ = "Path does not Exist!"
      msgType% = msgType% + 2
    Case 54
      Msg$ = "Bad File Mode!"
    Case 55
      Msg$ = "File is Already Open."
    Case 62
      Msg$ = "Read Attempt Past End of File."
    Case Else
      FileErrors = 3
      Exit Function
  End Select
  Response% = MsgBox(Msg$, msgType%, "Disk Error")
  Select Case Response%
    Case 1, 4
      FileErrors = 0
    Case 5
      FileErrors = 1
    Case 2, 3
      FileErrors = 2
    Case Else
      FileErrors = 3
  End Select
End Function

Public Function Finalize() As Boolean
    'Once we are done with the class instance we call this
    'function to discount it and finish winsock service if
    'it was the last one.
    m_lSocksQuantity = m_lSocksQuantity - 1
    'if the service was initiated and there's no more instances
    'of the class then we finish the service
    If m_bInit And m_lSocksQuantity = 0 Then
        Call WSACleanup
        Debug.Print "OK Winsock Service Terminated!"
        Subclass_Terminate
        Debug.Print "OK SubClass Finalized!"
        m_bInit = False
    End If
End Function

Public Function GetbSize(bArray() As Byte) As Long
    On Error GoTo GetSizeErr
    GetbSize = UBound(bArray) + 1
    Exit Function
GetSizeErr:
    GetbSize = 0
End Function

Public Function GetErrorDescription(ByVal lErrorCode As Long) As String
    'This function receives a number that represents an error
    'and returns the corresponding description string.
    Select Case lErrorCode
        Case WSAEACCES
            GetErrorDescription = "Permission denied."
        Case WSAEADDRINUSE
            GetErrorDescription = "Address already in use."
        Case WSAEADDRNOTAVAIL
            GetErrorDescription = "Cannot assign requested address."
        Case WSAEAFNOSUPPORT
            GetErrorDescription = "Address family not supported by protocol family."
        Case WSAEALREADY
            GetErrorDescription = "Operation already in progress."
        Case WSAECONNABORTED
            GetErrorDescription = "Software caused connection abort."
        Case WSAECONNREFUSED
            GetErrorDescription = "Connection refused."
        Case WSAECONNRESET
            GetErrorDescription = "Connection reset by peer."
        Case WSAEDESTADDRREQ
            GetErrorDescription = "Destination address required."
        Case WSAEFAULT
            GetErrorDescription = "Bad address."
        Case WSAEHOSTUNREACH
            GetErrorDescription = "No route to host."
        Case WSAEINPROGRESS
            GetErrorDescription = "Operation now in progress."
        Case WSAEINTR
            GetErrorDescription = "Interrupted function call."
        Case WSAEINVAL
            GetErrorDescription = "Invalid argument."
        Case WSAEISCONN
            GetErrorDescription = "Socket is already connected."
        Case WSAEMFILE
            GetErrorDescription = "Too many open files."
        Case WSAEMSGSIZE
            GetErrorDescription = "Message too long."
        Case WSAENETDOWN
            GetErrorDescription = "Network is down."
        Case WSAENETRESET
            GetErrorDescription = "Network dropped connection on reset."
        Case WSAENETUNREACH
            GetErrorDescription = "Network is unreachable."
        Case WSAENOBUFS
            GetErrorDescription = "No buffer space available."
        Case WSAENOPROTOOPT
            GetErrorDescription = "Bad protocol option."
        Case WSAENOTCONN
            GetErrorDescription = "Socket is not connected."
        Case WSAENOTSOCK
            GetErrorDescription = "Socket operation on nonsocket."
        Case WSAEOPNOTSUPP
            GetErrorDescription = "Operation not supported."
        Case WSAEPFNOSUPPORT
            GetErrorDescription = "Protocol family not supported."
        Case WSAEPROCLIM
            GetErrorDescription = "Too many processes."
        Case WSAEPROTONOSUPPORT
            GetErrorDescription = "Protocol not supported."
        Case WSAEPROTOTYPE
            GetErrorDescription = "Protocol wrong type for socket."
        Case WSAESHUTDOWN
            GetErrorDescription = "Cannot send after socket shutdown."
        Case WSAESOCKTNOSUPPORT
            GetErrorDescription = "Socket type not supported."
        Case WSAETIMEDOUT
            GetErrorDescription = "Connection timed out."
        Case WSAEWOULDBLOCK
            GetErrorDescription = "Resource temporarily unavailable."
        Case WSAHOST_NOT_FOUND
            GetErrorDescription = "Host not found."
        Case WSANOTINITIALISED
            GetErrorDescription = "Successful WSAStartup not yet performed."
        Case WSANO_DATA
            GetErrorDescription = "Valid name, no data record of requested type."
        Case WSANO_RECOVERY
            GetErrorDescription = "This is a nonrecoverable error."
        Case WSASYSNOTREADY
            GetErrorDescription = "Network subsystem is unavailable."
        Case WSATRY_AGAIN
            GetErrorDescription = "Nonauthoritative host not found."
        Case WSAVERNOTSUPPORTED
            GetErrorDescription = "Winsock.dll version out of range."
        Case Else
            GetErrorDescription = "Unknown error."
    End Select
End Function


Public Function GetSocketCount() As Long
    GetSocketCount = m_colSocketsInst.Count
End Function


Public Function HexToByte(HexStr As String, bHex() As Byte) As Boolean
    Dim lLen As Long
    Dim lPntr As Long
    If Len(HexStr) > 1 Then
        lLen = Len(HexStr) / 2
        ReDim bHex(lLen - 1)
        For lPntr = 0 To UBound(bHex)
            bHex(lPntr) = Val("&H" & Mid$(HexStr, lPntr * 2 + 1, 2))
        Next lPntr
        HexToByte = True
    End If
End Function

Public Function Initialize() As Boolean
    Dim udtWSAData As WSADATA
    Dim lRet As Long
    m_lSocksQuantity = m_lSocksQuantity + 1
    'if the service wasn't initiated yet we do it now
    Initialize = True
    If Not m_bInit Then
        If Subclass_Initialize Then
            'Start the winsock service
            lRet = WSAStartup(&H202, udtWSAData)
            If lRet > 0 Then
                Err.Raise lRet, "modSimple.Initialize", GetErrorDescription(lRet)
            Else
                m_bInit = True
                Debug.Print ("OK SubClass Initialized!")
            End If
        Else
            Initialize = False
        End If
    End If
End Function

Public Function IntegerToUnsigned(Value As Integer) As Long
    'The function takes a Long containing a value in the range 
    'of an unsigned Integer and returns an Integer that you 
    'can pass to an API that requires an unsigned Integer
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
End Function

Public Function IsSocketRegistered(ByVal lSocket As Long) As Boolean
    'Returns TRUE si the socket that is passed is registered
    'in the colSocketsInst collection.
    On Error GoTo Error_Handler
    Call m_colSocketsInst.Item("S" & lSocket)
    IsSocketRegistered = True
    Exit Function
Error_Handler:
    IsSocketRegistered = False
End Function

Public Sub LogError(Log$)
    Dim LogFile%
    LogFile% = OpenFile(App.Path + "\Socket.Log", 3, 0, 80)
    If LogFile% = 0 Then
        MsgBox "File Error with LogFile", 16, "ABORT PROCEDURE"
        Exit Sub
    End If
    Print #LogFile%, CStr(Now) + ": " + Log$
    Close LogFile%
End Sub

Public Function OpenFile(FileName$, Mode%, RLock%, RecordLen%) As Integer
  Const REPLACEFILE = 1, READAFILE = 2, ADDTOFILE = 3
  Const RANDOMFILE = 4, BINARYFILE = 5
  Const NOLOCK = 0, RDLOCK = 1, WRLOCK = 2, RWLOCK = 3
  Dim FileNum%
  Dim Action%
  FileNum% = FreeFile
  On Error GoTo OpenErrors
  Select Case Mode
    Case REPLACEFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Output Shared As FileNum%
            Case RDLOCK
                Open FileName For Output Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Output Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Output Lock Read Write As FileNum%
        End Select
    Case READAFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Input Shared As FileNum%
            Case RDLOCK
                Open FileName For Input Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Input Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Input Lock Read Write As FileNum%
        End Select
    Case ADDTOFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Append Shared As FileNum%
            Case RDLOCK
                Open FileName For Append Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Append Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Append Lock Read Write As FileNum%
        End Select
    Case RANDOMFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Random Shared As FileNum% Len = RecordLen%
            Case RDLOCK
                Open FileName For Random Lock Read As FileNum% Len = RecordLen%
            Case WRLOCK
                Open FileName For Random Lock Write As FileNum% Len = RecordLen%
            Case RWLOCK
                Open FileName For Random Lock Read Write As FileNum% Len = RecordLen%
        End Select
    Case BINARYFILE
        Select Case RLock%
            Case NOLOCK
                Open FileName For Binary Shared As FileNum%
            Case RDLOCK
                Open FileName For Binary Lock Read As FileNum%
            Case WRLOCK
                Open FileName For Binary Lock Write As FileNum%
            Case RWLOCK
                Open FileName For Binary Lock Read Write As FileNum%
        End Select
    Case Else
      Exit Function
  End Select
  OpenFile = FileNum%
Exit Function
OpenErrors:
  Action% = FileErrors(Err)
  Select Case Action%
    Case 0
      Resume            'Resumes at line where ERROR occured
    Case 1
        Resume Next     'Resumes at line after ERROR
    Case 2
        OpenFile = 0     'Unrecoverable ERROR-reports error, exits function with error code
        Exit Function
    Case Else
        MsgBox Error$(Err) + vbCrLf + "After line " + CStr(Erl) + vbCrLf + "Program will TERMINATE!"
        'Unrecognized ERROR-reports error and terminates.
        'End
  End Select
End Function

Public Function PeekB(ByVal lpdwData As Long) As Byte
    CopyMemory PeekB, ByVal lpdwData, 1
End Function

Public Sub PrintDebug(Msg As String)
    Debug.Print Msg
    If DbgFlg Then Call LogError(Msg)
End Sub

Public Function RegisterSocket(ByVal lSocket As Long, ByVal lObjectPointer As Long, ByVal blnEvents As Boolean) As Boolean
    'Adds the socket to the m_colSocketsInst collection, and
    'registers that socket with WSAAsyncSelect Winsock API
    'function to receive network events for the socket.
    'If this socket is the first one to be registered, the
    'window and collection will be created in this function as well.
    Dim lEvents As Long
    Dim lRet As Long
    Dim lErrorCode As Long
    If m_colSocketsInst Is Nothing Then
        Set m_colSocketsInst = New Collection
        Call PrintDebug("OK Created socket collection")
        If CreateWinsockMessageWindow <> 0 Then
            Err.Raise sckOutOfMemory, "modSimple.RegisterSocket", "Out of memory"
        End If
        Subclass_Subclass (m_hWindow)
    End If
    Subclass_AddSocketMessage lSocket, lObjectPointer
    'Do we need to register socket events?
    If blnEvents Then
        lEvents = FD_READ Or FD_WRITE Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE
        lRet = WSAAsyncSelect(lSocket, m_hWindow, SOCKET_MESSAGE, lEvents)
        If lRet = SOCKET_ERROR Then
            Call PrintDebug("ERROR trying to register events from socket " & CStr(lSocket))
            lErrorCode = Err.LastDllError
            Err.Raise lErrorCode, "modSimple.RegisterSocket", GetErrorDescription(lErrorCode)
        Else
            Call PrintDebug("OK Registered events from socket " & CStr(lSocket))
        End If
    End If
    Call m_colSocketsInst.Add(lObjectPointer, "S" & lSocket)
    RegisterSocket = True
End Function

Public Function StringFromPointer(ByVal lPointer As Long) As String
    'Receives a string pointer and it turns it into a regular string.
    Dim sTemp As String
    Dim lRetVal As Long
    sTemp = String$(lstrlenA(ByVal lPointer), 0)
    lRetVal = lstrcpyA(ByVal sTemp, ByVal lPointer)
    If lRetVal Then StringFromPointer = sTemp
End Function

Public Function StrToByte(strInput As String) As Byte()
    Dim lPntr As Long
    Dim bTmp() As Byte
    Dim bArray() As Byte
    If Len(strInput) = 0 Then Exit Function
    ReDim bTmp(LenB(strInput) - 1) 'Memory length
    ReDim bArray(Len(strInput) - 1) 'String length
    CopyMemory bTmp(0), ByVal StrPtr(strInput), LenB(strInput)
    'Examine every second byte
    For lPntr = 0 To UBound(bArray)
        If bTmp(lPntr * 2 + 1) > 0 Then
            'bArray(lPntr) = Asc(Mid$(strInput, lPntr + 1, 1))
            StrToByte = bTmp
            Exit Function
        Else
            bArray(lPntr) = bTmp(lPntr * 2)
        End If
    Next lPntr
    StrToByte = bArray
End Function

Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    'Return the address of the passed function in the passed dll
    Subclass_AddrFunc = API_GetProcAddress(API_GetModuleHandle(sDLL), sProc)
End Function

Private Function Subclass_AddrMsgTbl(ByRef aMsgTbl() As Long) As Long
    'Return the address of the low bound of the passed table array
    On Error Resume Next 'The table may not be dimensioned yet so we need protection
    Subclass_AddrMsgTbl = VarPtr(aMsgTbl(1)) 'Get the address of the first element of the passed message table
    On Error GoTo 0                              'Switch off error protection
End Function

Private Sub Subclass_AddSocketMessage(ByVal lSocket As Long, ByVal lObjPntr As Long)
    Dim Count As Long
    For Count = 1 To lMsgCntB
        Select Case lTableB1(Count)
            Case -1
                lTableB1(Count) = lSocket
                lTableB2(Count) = lObjPntr
                Exit Sub
            Case lSocket
                Call PrintDebug("WARNING: Socket already registered!")
                Exit Sub
        End Select
    Next Count
    lMsgCntB = lMsgCntB + 1
    ReDim Preserve lTableB1(1 To lMsgCntB)
    ReDim Preserve lTableB2(1 To lMsgCntB)
    lTableB1(lMsgCntB) = lSocket
    lTableB2(lMsgCntB) = lObjPntr
    Subclass_PatchTableB
End Sub

Private Sub Subclass_DelSocketMessage(ByVal lSocket As Long)
    Dim Count As Long
    For Count = 1 To lMsgCntB
        If lTableB1(Count) = lSocket Then
            lTableB1(Count) = -1
            lTableB2(Count) = -1
            Exit Sub
        End If
    Next Count
End Sub

Private Function Subclass_Initialize() As Boolean
    Const PATCH_01 As Long = 16                   'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_03 As Long = 72                   'Relative address of SetWindowsLong
    Const PATCH_04 As Long = 77                   'Relative address of WSACleanup
    Const PATCH_06 As Long = 89                   'Relative address of KillTimer
    Const PATCH_08 As Long = 113                  'Relative address of CallWindowProc
    Const FUNC_EBM As String = "EbMode"           'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL As String = "SetWindowLongA"   'SetWindowLong allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const FUNC_CWP As String = "CallWindowProcA"  'We use CallWindowProc to call the original WndProc
    Const FUNC_WCU As String = "WSACleanup"       'closesocket is called when the program is closed to release the sockets
    Const FUNC_KTM As String = "KillTimer"        'KillTimer destroys the control timer
'    Const MOD_VBA5 As String = "vba5"             'Location of the EbMode function if running VB5
    Const MOD_VBA6 As String = "vba6"             'Location of the EbMode function if running VB6
    Const MOD_USER As String = "user32"           'Location of the SetWindowLong & CallWindowProc functions
    Const MOD_WS   As String = "ws2_32"           'Location of the closesocket function
    Dim lPntr    As Long                          'Loop index
    Dim nLen     As Long                          'String lengths
    Dim sHex     As String                        'Hex code string
    Dim bHex()   As Byte                          'Binary code string
    'Store the hex pair machine code representation in sHex
    sHex = "5850505589E55753515231C0FCEB09E8xxxxx01x85C074258B45103D0080000074543D01800000746CE8310000005A595B5FC9C21400E824000000EBF168xxxxx02x6AFCFF750CE8xxxxx03xE8xxxxx04x68xxxxx05x6A00E8xxxxx06xEBCFFF7518FF7514FF7510FF750C68xxxxx07xE8xxxxx08xC3BBxxxxx09x8B4514BFxxxxx0Ax89D9F2AF75A529CB4B8B1C9Dxxxxx0BxEB1DBBxxxxx0Cx8B4514BFxxxxx0Dx89D9F2AF758629CB4B8B1C9Dxxxxx0Ex895D088B1B8B5B1C89D85A595B5FC9FFE0"
    If HexToByte(sHex, bHex) Then
        nLen = UBound(bHex) + 1
        nAddrSubclass = VirtualAlloc(0, nLen, MEM_COMMIT, PAGE_RWX)  'Allocate executable memory
        'Copy the code to allocated memory
        Call CopyMemory(ByVal nAddrSubclass, ByVal VarPtr(bHex(0)), nLen)
        If Subclass_InIDE Then
            'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
            Call CopyMemory(ByVal nAddrSubclass + 13, &H9090, 2)
            lPntr = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)     'Get the address of EbMode in vba6.dll
            Debug.Assert lPntr                                'Ensure the EbMode function was found
            Call Subclass_PatchRel(PATCH_01, lPntr)           'Patch the relative address to the EbMode api function
        End If
        Call API_LoadLibrary(MOD_WS)                          'Ensure ws2_32.dll is loaded before getting WSACleanup address
        Call Subclass_PatchRel(PATCH_03, Subclass_AddrFunc(MOD_USER, FUNC_SWL))     'Address of the SetWindowLong api function
        Call Subclass_PatchRel(PATCH_04, Subclass_AddrFunc(MOD_WS, FUNC_WCU))       'Address of the WSACleanup api function
        Call Subclass_PatchRel(PATCH_06, Subclass_AddrFunc(MOD_USER, FUNC_KTM))     'Address of the KillTimer api function
        Call Subclass_PatchRel(PATCH_08, Subclass_AddrFunc(MOD_USER, FUNC_CWP))     'Address of the CallWindowProc api function
        Subclass_Initialize = True
    End If
End Function

Private Sub Subclass_PatchRel(ByVal nOffset As Long, ByVal nTargetAddr As Long)
    'Patch the machine code buffer offset with the relative address to the target address
    Call CopyMemory(ByVal (nAddrSubclass + nOffset), nTargetAddr - nAddrSubclass - nOffset - 4, 4)
End Sub

Private Sub Subclass_PatchTableB()
    Const PATCH_0D As Long = 158
    Const PATCH_0E As Long = 174
    Call Subclass_PatchVal(PATCH_0C, lMsgCntB)
    Call Subclass_PatchVal(PATCH_0D, Subclass_AddrMsgTbl(lTableB1))
    Call Subclass_PatchVal(PATCH_0E, Subclass_AddrMsgTbl(lTableB2))
End Sub

Private Sub Subclass_PatchVal(ByVal nOffset As Long, ByVal nValue As Long)
    'Patch the machine code buffer offset with the passed value
    Call CopyMemory(ByVal (nAddrSubclass + nOffset), nValue, 4)
End Sub

Private Function Subclass_SetTrue(bValue As Boolean) As Boolean
    'Worker function for InIDE - will only be called whilst running in the IDE
    Subclass_SetTrue = True
    bValue = True
End Function

Private Function Subclass_Subclass(ByVal hWnd As Long) As Boolean
    'Set the window subclass
    Const PATCH_02 As Long = 62     'Address of the previous WndProc
    Const PATCH_05 As Long = 82     'Control timer handle
    Const PATCH_07 As Long = 108    'Address of the previous WndProc
    If hWndSub = 0 Then
        Debug.Assert API_IsWindow(hWnd)     'Invalid window handle
        hWndSub = hWnd                      'Store the window handle
        'Get the original window proc
        nAddrOriginal = API_GetWindowLong(hWnd, GWL_WNDPROC)
        Call Subclass_PatchVal(PATCH_02, nAddrOriginal)     'Original WndProc address for CallWindowProc, call the original WndProc
        Call Subclass_PatchVal(PATCH_07, nAddrOriginal)     'Original WndProc address for SetWindowLong, unsubclass on IDE stop
        'Set our WndProc in place of the original
        nAddrOriginal = API_SetWindowLong(hWnd, GWL_WNDPROC, nAddrSubclass)
        If nAddrOriginal <> 0 Then
          Subclass_Subclass = True                          'Success
        End If
    End If
    If Subclass_InIDE Then
        hTimer = API_SetTimer(0, 0, TIMER_TIMEOUT, nAddrSubclass) 'Create the control timer
        Call Subclass_PatchVal(PATCH_05, hTimer)    'Patch the control timer handle
    End If
    Debug.Assert Subclass_Subclass
End Function

Private Sub Subclass_Terminate()
    'UnSubclass and release the allocated memory
    Call Subclass_UnSubclass                        'UnSubclass if the Subclass thunk is active
    Call VirtualFree(nAddrSubclass, 0, MEM_RELEASE)  'Release the allocated memory
    Call PrintDebug("OK Freed subclass memory at: " & Hex$(nAddrSubclass))
    nAddrSubclass = 0
    ReDim lTableA1(1 To 1)
    ReDim lTableA2(1 To 1)
    ReDim lTableB1(1 To 1)
    ReDim lTableB2(1 To 1)
End Sub

Private Function Subclass_UnSubclass() As Boolean
    'Stop subclassing the window
    If hWndSub <> 0 Then
        lMsgCntA = 0
        lMsgCntB = 0
        Call Subclass_PatchVal(PATCH_09, lMsgCntA)  'Patch the TableA entry count to ensure no further Proc callbacks
        Call Subclass_PatchVal(PATCH_0C, lMsgCntB)  'Patch the TableB entry count to ensure no further Proc callbacks
        'Restore the original WndProc
        Call API_SetWindowLong(hWndSub, GWL_WNDPROC, nAddrOriginal)
        If hTimer <> 0 Then
            Call API_KillTimer(0&, hTimer)          'Destroy control timer
            hTimer = 0
        End If
        hWndSub = 0                                 'Indicate the subclasser is inactive
        Subclass_UnSubclass = True                  'Success
    End If
End Function

Public Sub UnregisterSocket(ByVal lSocket As Long)
    'Removes the socket from the m_colSocketsInst collection
    'If it is the last socket in that collection, the window
    'and colection will be destroyed as well.
    Subclass_DelSocketMessage lSocket
    On Error Resume Next
    Call m_colSocketsInst.Remove("S" & lSocket)
    If m_colSocketsInst.Count = 0 Then
        Set m_colSocketsInst = Nothing
        Subclass_UnSubclass
        DestroyWinsockMessageWindow
        Call PrintDebug("OK Destroyed socket collection")
    End If
End Sub

Private Function Subclass_InIDE() As Boolean
    'Return whether we're running in the IDE. Public for general utility purposes
    Debug.Assert Subclass_SetTrue(Subclass_InIDE)
End Function

