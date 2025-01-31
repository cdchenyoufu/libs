Attribute VB_Name = "modDebugPrint"
' http://www.vbforums.com/showthread.php?874127-Persistent-Debug-Print-Window
'
' This is a Stand-Alone module that can be thrown into any project.
' It works in conjunction with the PersistentDebugPrint program, and that program must be running to use this module.
' The only procedure you should worry about is the DebugPrint procedure.
' Basically, it does what it says, provides a "Debug" window that is persistent across your development IDE exits and starts (even IDE crashes).
'
Option Explicit
'
Private Type COPYDATASTRUCT
    dwData  As Long
    cbData  As Long
    lpData  As Long
End Type
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim mhWndTarget As Long
'
Const DoDebugPrint As Boolean = True
'

Private Function canStartServer() As Boolean
    Dim pth As String
    
    pth = GetSetting("dbgWindow", "settings", "path", "")
    If Not FileExists(pth) Then Exit Function
    
    Shell pth, vbNormalFocus
    
    If Err.Number = 0 Then
        Sleep 250
        ValidateTargetHwnd
        If mhWndTarget = 0 Then
            Sleep 250
            ValidateTargetHwnd
        End If
        canStartServer = (mhWndTarget <> 0)
    End If
    
End Function

Public Sub DebugPrint(sMsg As String)

    If Not DoDebugPrint Then Exit Sub

    Static bErrorMessageShown As Boolean
    
    ValidateTargetHwnd
    
    If mhWndTarget = 0& Then
        If Not bErrorMessageShown Then
            If Not canStartServer() Then
                MsgBox "The Persistent Debug Print Window could not be found. I can auto start it, but you havent run it yet for it to save its path to the registry.", vbCritical, "Persistent Debug Message"
                bErrorMessageShown = True
                Exit Sub
            End If
        End If
    End If

    SendStringToAnotherWindow sMsg
    
End Sub

Private Sub ValidateTargetHwnd()
    If IsWindow(mhWndTarget) Then
        Select Case WindowClass(mhWndTarget)
            Case "ThunderForm", "ThunderRT6Form"
                If WindowText(mhWndTarget) = "Persistent Debug Print Window" Then
                    Exit Sub
                End If
        End Select
    End If
    EnumWindows AddressOf EnumToFindTargetHwnd, 0&
End Sub

'callback - must be in a module
Private Function EnumToFindTargetHwnd(ByVal hWnd As Long, ByVal lParam As Long) As Long
    mhWndTarget = 0&                        ' We just set it every time to keep from needing to think about it before this is called.
    Select Case WindowClass(hWnd)
        Case "ThunderForm", "ThunderRT6Form"
            If WindowText(hWnd) = "Persistent Debug Print Window" Then
                mhWndTarget = hWnd
                Exit Function
            End If
    End Select
    EnumToFindTargetHwnd = 1&               ' Keep looking.
End Function

Private Function WindowClass(hWnd As Long) As String
    WindowClass = String$(1024&, vbNullChar)
    WindowClass = Left$(WindowClass, GetClassName(hWnd, WindowClass, 1024&))
End Function

Private Function WindowText(hWnd As Long) As String
    ' Form or control.
    WindowText = String$(GetWindowTextLength(hWnd) + 1&, vbNullChar)
    Call GetWindowText(hWnd, WindowText, Len(WindowText))
    WindowText = Left$(WindowText, InStr(WindowText, vbNullChar) - 1&)
End Function

Private Sub SendStringToAnotherWindow(sMsg As String)
    Dim cds             As COPYDATASTRUCT
    Dim lpdwResult      As Long
    Dim Buf()           As Byte
    Const WM_COPYDATA   As Long = &H4A&
    '
    ReDim Buf(1 To Len(sMsg) + 1&)
    Call CopyMemory(Buf(1&), ByVal sMsg, Len(sMsg)) ' Copy the string into a byte array, converting it to ASCII.
    cds.dwData = 3&
    cds.cbData = Len(sMsg) + 1&
    cds.lpData = VarPtr(Buf(1&))
    'Call SendMessage(hWndTarget, WM_COPYDATA, Me.hwnd, cds)
    SendMessageTimeout mhWndTarget, WM_COPYDATA, 0&, cds, 0&, 1000&, lpdwResult ' Return after a second even if receiver didn't acknowledge.
End Sub

