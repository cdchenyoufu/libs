Attribute VB_Name = "modDzTest"
Option Explicit

Public hLib As Long
Public hLib2 As Long
Public caBundle As String
Public ActiveResponse As CCurlResponse 'one active instance for simplicity

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Enum curlEvents
    ceInit
    ceProgress
    ceComplete
    ceHeader
    ceInfo
End Enum

Function initLib(errList As Collection) As Boolean
    
    If hLib <> 0 And hLib2 <> 0 Then
        initLib = True
        Exit Function
    End If
    
    Dim base() As String, b
    Const dll = "libcurl.dll"
    
    push base, App.path
    push base, App.path & "\bin"
    push base, GetParentFolder(App.path)
    push base, GetParentFolder(App.path) & "\bin"
    push base, GetParentFolder(App.path, 2)
    push base, GetParentFolder(App.path, 2) & "\bin"
    
    For Each b In base
        hLib = LoadLibrary(b & "\" & dll)
        If hLib <> 0 Then
            If Not errList Is Nothing Then errList.Add "Loaded " & b & "\" & dll
            hLib2 = LoadLibrary(b & "\vb" & dll)
            If hLib2 = 0 Then
                If Not errList Is Nothing Then errList.Add "Failed to load vbLibcurl.dll from same directory?!"
            Else
                If Not errList Is Nothing Then errList.Add "Loaded vbLibCurl.dll from same directory."
                caBundle = b & "\curl-ca-bundle.crt"
                If FileExists(caBundle) Then
                    If Not errList Is Nothing Then errList.Add "Found curl-ca-bundle.crt"
                Else
                    caBundle = Empty
                    If Not errList Is Nothing Then errList.Add "curl-ca-bundle.crt not found ssl will not work..."
                End If
            End If
            Exit For
        End If
    Next
        
    If hLib <> 0 And hLib2 <> 0 Then
        initLib = True
    Else
        If Not errList Is Nothing Then errList.Add "Could not load libcurl.dll"
    End If

End Function


Function DebugFunction(ByVal info As curl_infotype, ByVal rawBytes As Long, ByVal numBytes As Long, ByVal extra As Long) As Long

    Dim tmp As String, i As Long
    Dim b() As Byte
    
    If info >= 3 Then 'DATA_IN/OUT msgs
        vbcurl_easy_setopt ActiveResponse.owner.hCurl, CURLOPT_DEBUGFUNCTION, 0 'should we turn off dbg msgs now? less callbacks to us slow down
        Exit Function
    End If
    
    ReDim b(numBytes - 1)
    CopyMemory ByVal VarPtr(b(0)), ByVal rawBytes, numBytes
    tmp = StrConv(b, vbUnicode, &H409)
    
    If info = CURLINFO_HEADER_IN Then
        ActiveResponse.AddHeader tmp
    Else
        ActiveResponse.addInfo info, tmp
    End If
    
    DebugFunction = 0
    
End Function

'extra is the objPtr(activeClassObject)
Function WriteFunction(ByVal rawBytes As Long, ByVal sz As Long, ByVal nmemb As Long, ByVal extra As Long) As Long
    
    On Error Resume Next
    Dim totalBytes As Long, i As Long, b() As Byte, ret As CURLcode, v As Variant
    
    totalBytes = sz * nmemb
    
    If ActiveResponse.isMemFile Then
        ActiveResponse.memFile.memAppendBuf rawBytes, totalBytes
    Else
        ReDim b(totalBytes - 1)
        CopyMemory ByVal VarPtr(b(0)), ByVal rawBytes, totalBytes
        Put ActiveResponse.hFile, , b()
    End If
    
    DoEvents
    WriteFunction = ActiveResponse.notifyParent(totalBytes)  'if this returns 0 dl is aborted...
    
End Function

'multi way...

' This function illustrates a couple of key concepts in libcurl.vb.
' First, the data passed in rawBytes is an actual memory address
' from libcurl. Hence, the data is read using the MemByte() function
' found in the VBVM6Lib.tlb type library. Second, the extra parameter
' is passed as a raw long (via ObjPtr(buf)) in Sub EasyGet()), and
' we use the AsObject() function in VBVM6Lib.tlb to get back at it.

 'If extra <> 0 Then
        'Set obj = AsObject(extra)
        'Set buf = obj
        'buf.
        'ObjectPtr(obj) = 0&
        
        
'I would prefer the below: (and no tlb requirement)
'there isnt really a point in using the multi downloads though, unless one server was super slow
'out of several, your bandwidth is limited anyway so just do one at a time should be fine.
'rarely worth the extra complexity for most projects unless your literally making a multi downloader app.
'-----------------------------------------------
'Dim responses As New Collection 'of CCurlResponse
'
'Function getObj(ptr As Long) As CCurlResponse
'    Dim c As CCurlResponse
'    For Each c In responses
'        If ObjPtr(c) = ptr Then
'            Set getObj = c
'            Exit Function
'        End If
'    Next
'End Function
'
'Dim resp As CCurlResponse
'Set resp = getObj(extra)
'If resp Is Nothing Then Exit Function
