Attribute VB_Name = "Module1"
Option Explicit

' //////////////////////////////////////////////////////////////////////////////
'
' From original post by Bonnie West:
'
'     [VB6] Clone ListView
'     https://www.vbforums.com/showthread.php?732433-VB6-Clone-ListView
'
' patched to work with a x64 target process: (not tested by me)
'     https://www.vbforums.com/showthread.php?893457-RESOLVED-Problem-with-LVM_GETITEM-and-64-bit-OS-please-help
'
' //////////////////////////////////////////////////////////////////////////////

Private Const HDI_TEXT              As Long = &H2
Private Const HDM_FIRST             As Long = &H1200
Private Const HDM_GETITEMCOUNT      As Long = (HDM_FIRST + 0)
Private Const HDM_GETITEMW          As Long = (HDM_FIRST + 11)

Private Const LVIF_TEXT             As Long = &H1
Private Const LVM_FIRST             As Long = &H1000
Private Const LVM_GETITEMCOUNT      As Long = (LVM_FIRST + 4)
Private Const LVM_GETHEADER         As Long = (LVM_FIRST + 31)
Private Const LVM_GETITEMW          As Long = (LVM_FIRST + 75)
Private Const LVM_GETITEMTEXTW      As Long = (LVM_FIRST + 115)

Private Const MAX_PATH              As Long = 260

Private Const MEM_COMMIT            As Long = &H1000
Private Const MEM_RELEASE           As Long = &H8000&
Private Const MEM_RESERVE           As Long = &H2000
Private Const PAGE_READWRITE        As Long = &H4
Private Const PROCESS_VM_OPERATION  As Long = &H8
Private Const PROCESS_VM_READ       As Long = &H10
Private Const PROCESS_VM_WRITE      As Long = &H20

Private Const MAX_LVMSTRING         As Long = (MAX_PATH * 2) + 2

Private Type HDITEM
    mask       As Long
    cxy        As Long
    pszText    As Long
    hbm        As Long
    cchTextMax As Long
    fmt        As Long
    lParam     As Long
    iImage     As Long
    iOrder     As Long
    type       As Long
    pvFilter   As Long
    state      As Long
End Type

Private Type LVITEM
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    state      As Long
    stateMask  As Long
    pszText    As Long
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
    iGroupId   As Long
    cColumns   As Long
    puColumns  As Long
    piColFmt   As Long
    iGroup     As Long
End Type

Private Type HDITEM64
    mask       As Long
    cxy        As Long
    pszText    As Currency
    hbm        As Currency
    cchTextMax As Long
    fmt        As Long
    lParam     As Currency
    iImage     As Long
    iOrder     As Long
    type       As Long
    lPad       As Long
    pvFilter   As Currency
    state      As Long
End Type

Private Type LVITEM64
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    state      As Long
    stateMask  As Long
    lPad       As Long
    pszText    As Currency
    cchTextMax As Long
    iImage     As Long
    lParam     As Currency
    iIndent    As Long
    iGroupId   As Long
    cColumns   As Long
    lPad2      As Long
    puColumns  As Currency
    piColFmt   As Currency
    iGroup     As Long
End Type

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, Optional ByRef lpdwProcessId As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByRef lpBuffer As Any, ByVal nSize As Long, Optional ByRef lpNumberOfBytesRead As Long) As Long
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hwnd As Long, ByVal uMsg As Long, Optional ByVal wParam As Long, Optional ByVal lParam As Long) As Long
Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, Optional ByVal lpAddress As Long, Optional ByVal dwSize As Long, Optional ByVal flAllocationType As Long = MEM_COMMIT Or MEM_RESERVE, Optional ByVal flProtect As Long = PAGE_READWRITE) As Long
Private Declare Function VirtualFreeEx Lib "kernel32.dll" (ByVal hProcess As Long, ByVal lpAddress As Long, Optional ByVal dwSize As Long, Optional ByVal dwFreeType As Long = MEM_RELEASE) As Long
Private Declare Function WriteProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByRef lpBuffer As Any, ByVal nSize As Long, Optional ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)

Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function MessageBoxW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long

' ////////////////////////////////////////////////////////////////////////////////////////

' https://forum.sources.ru/index.php?showtopic=406899

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
 
Private Declare Function NtWow64WriteVirtualMemory64 Lib "ntdll" ( _
                         ByVal ProcessHandle As Long, _
                         ByVal BaseAddressL As Long, _
                         ByVal BaseAddressH As Long, _
                         ByRef Buffer As Any, _
                         ByVal BufferSizeL As Long, _
                         ByVal BufferSizeH As Long, _
                         ByRef NumberOfBytesWritten As LARGE_INTEGER) As Long

Public Sub GetListViewItems(ByVal lngListViewWnd As Long, ByRef strColumns() As String, ByRef strItems() As String)

    If IsWindow(lngListViewWnd) = 0 Then Exit Sub
    
    Dim lngHeaderWnd As Long, lngCol As Long, lngCols As Long, lngRow As Long, lngRows As Long
    Dim lngProcId As Long, lngProcess As Long, lngBuffer As Long, lngRet As Long
    Dim lngHDI As Long, lngLenB_HDI As Long, tHDI As HDITEM64
    Dim lngLVI As Long, lngLenB_LVI As Long, tLVI As LVITEM64
    
    lngHeaderWnd = SendMessageW(lngListViewWnd, LVM_GETHEADER)
    lngCols = SendMessageW(lngHeaderWnd, HDM_GETITEMCOUNT)
    If lngCols = 0 Then Exit Sub
    
    lngRows = SendMessageW(lngListViewWnd, LVM_GETITEMCOUNT)
    If lngRows = 0 Then Exit Sub
    
    Call GetWindowThreadProcessId(lngListViewWnd, lngProcId)
    lngProcess = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, 0&, lngProcId)
    If lngProcess = 0 Then Exit Sub

    lngLenB_HDI = LenB(tHDI)
    lngHDI = VirtualAllocEx(lngProcess, , lngLenB_HDI)
    
    If lngHDI <> 0 Then
        tHDI.cchTextMax = MAX_LVMSTRING
        lngBuffer = VirtualAllocEx(lngProcess, , tHDI.cchTextMax)
        If lngBuffer <> 0 Then
            tHDI.mask = HDI_TEXT
            ReDim strColumns(0 To lngCols - 1) As String
            For lngCol = 0 To lngCols - 1
'                tHDI.pszText = lngBuffer
                PutMem4 tHDI.pszText, ByVal lngBuffer
                If WriteProcessMemory(lngProcess, lngHDI, tHDI, lngLenB_HDI, lngRet) <> 0 Then
                    Debug.Assert lngRet = lngLenB_HDI
                    If SendMessageW(lngHeaderWnd, HDM_GETITEMW, lngCol, lngHDI) <> 0 Then
                        If ReadProcessMemory(lngProcess, lngHDI, tHDI, lngLenB_HDI, lngRet) <> 0 Then
                            Debug.Assert lngRet = lngLenB_HDI
                            SysReAllocStringLen VarPtr(strColumns(lngCol)), , tHDI.cchTextMax \ 2 - 1
                            If ReadProcessMemory(lngProcess, lngBuffer, ByVal StrPtr(strColumns(lngCol)), tHDI.cchTextMax - 2, lngRet) <> 0 Then
                                Debug.Assert lngRet = tHDI.cchTextMax - 2
                                strColumns(lngCol) = Left$(strColumns(lngCol), lstrlenW(StrPtr(strColumns(lngCol))))
                            Else
                                strColumns(lngCol) = vbNullString
                            End If
                        End If
                    End If
                End If
            Next
            lngRet = VirtualFreeEx(lngProcess, lngBuffer): Debug.Assert lngRet
        End If
        lngRet = VirtualFreeEx(lngProcess, lngHDI): Debug.Assert lngRet
    End If
                        
    lngLenB_LVI = LenB(tLVI)
    lngLVI = VirtualAllocEx(lngProcess, , lngLenB_LVI)
    If lngLVI <> 0 Then
        tLVI.cchTextMax = MAX_LVMSTRING
        lngBuffer = VirtualAllocEx(lngProcess, , tLVI.cchTextMax)
        If lngBuffer Then
            tLVI.mask = LVIF_TEXT
            lngRows = lngRows - 1
            lngCols = lngCols - 1
            ReDim strItems(0 To lngRows, 0 To lngCols) As String
            For lngRow = 0 To lngRows
                tLVI.iItem = lngRow
                For lngCol = 0 To lngCols
                    tLVI.iSubItem = lngCol
'                    tLVI.pszText = lngBuffer
                    PutMem4 tLVI.pszText, ByVal lngBuffer
                    If WriteProcessMemory(lngProcess, lngLVI, tLVI, lngLenB_LVI, lngRet) <> 0 Then
                        Debug.Assert lngRet = lngLenB_LVI
                        If SendMessageW(lngListViewWnd, LVM_GETITEMW, , lngLVI) <> 0 Then
                            If ReadProcessMemory(lngProcess, lngLVI, tLVI, lngLenB_LVI, lngRet) <> 0 Then
                                Debug.Assert lngRet = lngLenB_LVI
                                SysReAllocStringLen VarPtr(strItems(lngRow, lngCol)), , tLVI.cchTextMax \ 2 - 1
                                If ReadProcessMemory(lngProcess, lngBuffer, ByVal StrPtr(strItems(lngRow, lngCol)), tLVI.cchTextMax - 2, lngRet) <> 0 Then
                                    Debug.Assert lngRet = tLVI.cchTextMax - 2
                                    strItems(lngRow, lngCol) = Left$(strItems(lngRow, lngCol), lstrlenW(StrPtr(strItems(lngRow, lngCol))))
                                Else
                                    strItems(lngRow, lngCol) = vbNullString
                                End If
                            End If
                        End If
                    End If
                Next lngCol
            Next lngRow
            lngRet = VirtualFreeEx(lngProcess, lngBuffer): Debug.Assert lngRet
        End If
        lngRet = VirtualFreeEx(lngProcess, lngLVI): Debug.Assert lngRet
    End If
    lngRet = CloseHandle(lngProcess): Debug.Assert lngRet

End Sub












