VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCurlResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public owner As CCurlDownload
Public CurlReturnCode As CURLcode
Public memFile As CMemBuffer  'can be nothing if downloading to file
Public headers As New Collection
Public DebugMsgs As New Collection

Public TotalTime
Public contentType
Public ResponseCode
Public DownloadLength 'if not set in header then will be -1
Public BytesReceived As Double

Public abort As Boolean
Public hFile As Long
Public localPath As String
Public url As String

Private gotSizes As Boolean


Friend Function Initilize(url As String, Optional toFile As String) As Boolean
    
    On Error GoTo hell

    Me.url = url
    If owner.hCurl = 0 Then Exit Function
    
    If Len(toFile) > 0 Then
        If FileExists(toFile) Then DeleteFile toFile
        localPath = toFile
        hFile = FreeFile
        Open toFile For Binary As hFile 'this can throw an error - also has a 2gb file size limitation
    Else
        Set memFile = New CMemBuffer
    End If
    
    Initilize = True
    
    Exit Function
hell:
    hFile = 0
    Set memFile = Nothing

End Function

Friend Sub Finalize()
    
    Dim v As Variant, ret As CURLcode
    
    ''set request stats in this class
    ret = vbcurl_easy_getinfo(owner.hCurl, CURLINFO_CONTENT_TYPE, v)
    If ret = CURLE_OK Then contentType = v
    
    ret = vbcurl_easy_getinfo(owner.hCurl, CURLINFO_RESPONSE_CODE, v)
    If ret = CURLE_OK Then ResponseCode = v
    
    ret = vbcurl_easy_getinfo(owner.hCurl, CURLINFO_TOTAL_TIME, v)
    If ret = CURLE_OK Then TotalTime = v
    
    If memFile Is Nothing Then
        Close hFile
        hFile = 0
    End If
    
    'If Not owner.INotify Is Nothing Then owner.INotify.Complete Me
    owner.TriggerEvent ceComplete, Me
    
End Sub

'we already have this w/ BytesReceived
'Property Get CurrentDownloadSize() As Long
'        'ret = vbcurl_easy_getinfo(hCurl, CURLINFO_SIZE_DOWNLOAD, v)
'        'If ret = CURLE_OK Then DownloadSize = v
'End Property

Property Get isMemFile() As Boolean
    isMemFile = Not memFile Is Nothing
End Property

'response header from server
Friend Function AddHeader(msg As String)
    headers.Add msg
    'If Not owner.INotify Is Nothing Then owner.INotify.Header Me, msg
    owner.TriggerEvent ceHeader, Me, msg
End Function

Friend Function addInfo(info As curl_infotype, msg As String)
    DebugMsgs.Add msg
    'If Not owner.INotify Is Nothing Then owner.INotify.InfoMsg Me, info, msg
    owner.TriggerEvent ceInfo, Me, info, msg
End Function

Friend Function notifyParent(size As Long) As Long
    
    On Error Resume Next
    Dim ret As CURLcode, v As Variant
    
    BytesReceived = BytesReceived + size
    
    If Not gotSizes Then
        ret = vbcurl_easy_getinfo(owner.hCurl, CURLINFO_CONTENT_LENGTH_DOWNLOAD, v) 'can be -1 if not set in http header
        If ret = CURLE_OK Then DownloadLength = v
        'If Not owner.INotify Is Nothing Then owner.INotify.Init Me
        owner.TriggerEvent ceInit, Me
        gotSizes = True
    End If
    
    'If Not owner.INotify Is Nothing Then owner.INotify.Progress Me 'consumer can set abort if desired...
    owner.TriggerEvent ceProgress, Me
    
    If abort Then
        notifyParent = 0
    Else
        notifyParent = size
    End If

End Function

Function dump() As String
    Dim tmp() As String
    
    push tmp, "CurlCode = " & CurlReturnCode 'curlCode2Text(Me.CurlReturnCode)
    push tmp, "ContentType = " & contentType
    push tmp, "ResponseCode = " & ResponseCode
    push tmp, "DownloadLength  = " & DownloadLength
    push tmp, "TotalTime = " & TotalTime
    
    If isMemFile Then
        push tmp, "isMemFile = True"
    Else
        push tmp, "localPath = " & localPath
    End If
    
    push tmp, "url = " & url
    
    push tmp, "headers:"
    push tmp, Join(c2a(headers), "")

    dump = Join(tmp, vbCrLf)
    
End Function

 
