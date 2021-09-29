VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBinaryClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const formatName = "Binary Data"
Private formatID As Long
Private Const LANG_US = &H409
Private Const GMEM_DDESHARE = &H2000

Private ClipboardIsOpen As Boolean
Private mHwnd As Long


Property Let hWnd(h As Long)
    mHwnd = h
End Property

Private Function RegisterFormat(ByVal sName As String) As Boolean
    Dim wFormat As Long
    wFormat = RegisterClipboardFormat(sName & Chr$(0))
    If (wFormat > &HC000&) Then
        formatID = wFormat
        RegisterFormat = True
    End If
End Function

Public Function GetString(ByRef buf As String) As Boolean
    Dim tmp() As Byte
    Dim ret As Boolean
    buf = Empty
    If GetData(tmp) Then
        buf = StrConv(tmp, vbUnicode, LANG_US)
        GetString = True
    End If
End Function

Public Function GetData(ByRef buf() As Byte) As Boolean
    Dim hMem As Long, lSize As Long, lPtr As Long
    
    Erase buf
    
    If Not ClipboardOpen() Then Exit Function
    
    hMem = GetClipboardMemoryHandle(formatID)
    
    If (hMem <> 0) Then
        lSize = GlobalSize(hMem)
        lPtr = GlobalLock(hMem)
        If (lSize > 0) Then
            ReDim buf(0 To lSize - 1) As Byte
            CopyMemory buf(0), ByVal lPtr, lSize
        End If
        GlobalUnlock hMem
        GetData = (lSize > 0)
    End If

    ClipboardClose
    
End Function

Public Function SetString(ByRef buf As String) As Boolean
    Dim tmp() As Byte
    tmp() = StrConv(buf, vbFromUnicode, LANG_US)
    SetString = SetData(tmp)
End Function

Public Function SetData(ByRef buf() As Byte) As Boolean

    Dim lSize As Long
    Dim lPtr As Long
    Dim hMem As Long

    If Not ClipboardOpen() Then Exit Function
    EmptyClipboard
    
    lSize = UBound(buf) - LBound(buf) + 1
    hMem = GlobalAlloc(GMEM_DDESHARE, lSize) 'clipboard owns memory after this..
    
    If (hMem <> 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, buf(LBound(buf)), lSize
        GlobalUnlock hMem
                
        If (SetClipboardData(formatID, hMem) <> 0) Then
            SetData = True
        Else
            GlobalFree hMem 'wqweto bugfix 9.29.21
        End If
    End If
    
    ClipboardClose
 
End Function

Private Function IsDataAvailableForFormat(ByVal lFormatId As Long) As Boolean
    Dim lR As Long
    lR = IsClipboardFormatAvailable(lFormatId)
    IsDataAvailableForFormat = (lR <> 0)
End Function

Private Function GetClipboardMemoryHandle(ByVal lFormatId As Long) As Long
    If Not ClipboardOpen() Then Exit Function
    If (IsDataAvailableForFormat(lFormatId)) Then
        GetClipboardMemoryHandle = GetClipboardData(lFormatId)
    End If
End Function

Private Function SetClipboardMemoryHandle(ByVal lFormatId As Long, ByVal hMem As Long) As Boolean
    SetClipboardMemoryHandle = (SetClipboardData(lFormatId, hMem) <> 0)
End Function

Public Sub Clear()
    ClipboardOpen
    EmptyClipboard
    ClipboardClose
End Sub

Private Sub ClipboardClose()
    If Not ClipboardIsOpen Then Exit Sub
    CloseClipboard
    ClipboardIsOpen = False
End Sub

Private Function ClipboardOpen() As Boolean
    Dim lR As Long
    
    If mHwnd = 0 Then Err.Raise 1, "CBinaryClipboard", "hwnd must be set before accessing clipboard."
    
    If ClipboardIsOpen Then
        ClipboardOpen = True
        Exit Function
    End If
    
    lR = OpenClipboard(mHwnd)
    If (lR > 0) Then
        ClipboardIsOpen = True
        ClipboardOpen = True
    Else
        ClipboardIsOpen = False
        ClipboardOpen = False
    End If
    
End Function

Private Sub Class_Initialize()
    If Not RegisterFormat(formatName) Then
        Err.Raise 1, "CBinaryClipboard", "Failed to register binary format.."
    End If
    On Error Resume Next
    Dim f As Form
    For Each f In Forms
        mHwnd = f.hWnd
        Exit Sub
    Next
End Sub

Private Sub Class_Terminate()
    ClipboardClose
End Sub


