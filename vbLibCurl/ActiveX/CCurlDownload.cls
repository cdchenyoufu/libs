VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCurlDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'so we broke this out into its own class because libcurl supports soooo many options
'once you get to wanting to configure it more you are beyond the number of args you want to
'pass to a single download function. You could just use the setopt api and enum vals but the names
'arent always friendly and are all longer by default

Public hCurl As Long
Private hList As Long
Private hForm As Long

Public errList As New Collection

Event Init(obj As CCurlResponse)
Event Progress(obj As CCurlResponse)
Event Complete(obj As CCurlResponse)
Event Header(obj As CCurlResponse, ByVal msg As String)
Event InfoMsg(obj As CCurlResponse, ByVal info As curl_infotype, ByVal msg As String)


Property Get Protocols() As String
    Protocols = Join(libcurlProtocols(), ",")
End Property

Property Get ZlibVersion() As String
    ZlibVersion = libZVersion()
End Property

Property Get SSL_Version() As String
    SSL_Version = sslVersion
End Property

Friend Sub TriggerEvent(e As curlEvents, obj As CCurlResponse, Optional arg1, Optional arg2)
    If e = ceInit Then RaiseEvent Init(obj)
    If e = ceProgress Then RaiseEvent Progress(obj)
    If e = ceComplete Then RaiseEvent Complete(obj)
    If e = ceHeader Then RaiseEvent Header(obj, arg1)
    If e = ceInfo Then RaiseEvent InfoMsg(obj, arg1, arg2)
End Sub

'you can overwrite an automatic header, or remove it by specifying "name;"
Function AddHeader(aryStrOrCol) As Boolean
    On Error GoTo hell
    Dim s As String, X
    If hList = 0 Then hList = vbcurl_slist_create()
    If IsArray(aryStrOrCol) Or TypeName(aryStrOrCol) = "Collection" Then
        For Each X In aryStrOrCol
            s = CStr(X)
            If Len(s) > 0 Then vbcurl_slist_append hList, strPtr(s)
        Next
    Else
        s = CStr(aryStrOrCol)
        If Len(s) > 0 Then vbcurl_slist_append hList, strPtr(s)
    End If
    AddHeader = True
Exit Function
hell:
End Function

Function AddFormElement(Name As String, Value As String) As CURLFORMcode
    Dim part As Long
    
    If hForm = 0 Then hForm = vbcurl_form_create()
    
    part = vbcurl_form_create_part(hForm)
    
    Call vbcurl_form_add_four_to_part( _
        part, _
        CURLFORM_COPYNAME, Name, _
        CURLFORM_COPYCONTENTS, Value _
    )
    
    AddFormElement = vbcurl_form_add_part(hForm, part)
    
End Function

Function AddFormFileUpload(Name As String, path As String, Optional contentType As String = "application/binary") As CURLFORMcode
    Dim part As Long
    
    If hForm = 0 Then hForm = vbcurl_form_create()
    
    part = vbcurl_form_create_part(hForm)
    
    Call vbcurl_form_add_six_to_part( _
        part, _
        CURLFORM_COPYNAME, Name, _
        CURLFORM_FILE, path, _
        CURLFORM_CONTENTTYPE, contentType _
    )
                    
    AddFormFileUpload = vbcurl_form_add_part(hForm, part)
    
End Function

Sub ClearHeaders()
     If hList <> 0 Then vbcurl_slist_free hList
     hList = 0
End Sub

Sub ClearForm()
    If hForm <> 0 Then vbcurl_form_free hForm
    hForm = 0
End Sub

Private Sub Class_Initialize()
    
    If Not initLib(errList) Then
        Exit Sub
    End If

    hCurl = vbcurl_easy_init()
    
End Sub

Private Sub Class_Terminate()
    vbcurl_easy_cleanup hCurl
    If hList <> 0 Then vbcurl_slist_free hList
    If hForm <> 0 Then vbcurl_form_free hForm
End Sub

'these ones are typically set once per class use so give one liner shortcut
Sub Configure(Useragent As String, _
             Optional followRedirects As Boolean = True, _
             Optional maxTimeout As Long, _
             Optional connectTimeout As Long _
)

    Me.Useragent = Useragent
    Me.followRedirects = followRedirects
    Me.maxTime = maxTimeout
    Me.connectTimeout = connectTimeout
    'If Not notifyUI Is Nothing Then Set INotify = notifyUI
    
End Sub

Property Let Useragent(ua As String)
    vbcurl_easy_setopt hCurl, CURLOPT_USERAGENT, ua
End Property

Property Let connectTimeout(secs As Long)
    vbcurl_easy_setopt hCurl, CURLOPT_CONNECTTIMEOUT, secs
End Property

Property Let followRedirects(v As Boolean)
     vbcurl_easy_setopt hCurl, CURLOPT_FOLLOWLOCATION, IIf(v, 1, 0)
End Property

Property Let maxTime(secs As Long)
    vbcurl_easy_setopt hCurl, CURLOPT_TIMEOUT, secs
End Property

Function escape(ByVal buf As String) As String
    Dim tmp As Long, s As String
    tmp = vbcurl_string_escape(strPtr(buf), Len(buf))
    CopyMemory ByVal VarPtr(s), tmp, 4 'steal a ref to an existing BSTR so we now own it
    escape = s
End Function

Function unescape(ByVal buf As String) As String
    Dim tmp  As Long, s As String
    tmp = vbcurl_string_unescape(strPtr(buf), Len(buf))
    CopyMemory ByVal VarPtr(s), tmp, 4 'steal a ref to an existing BSTR so we now own it
    unescape = s
End Function

Property Get CurlVersion() As String
    CurlVersion = libCurlVersion()
End Property

'these are usually set per request
Property Let Referrer(v As String)
     vbcurl_easy_setopt hCurl, CURLOPT_REFERER, v
End Property

Property Let Cookie(c As String)
     vbcurl_easy_setopt hCurl, CURLOPT_COOKIE, c
End Property

Property Get SupportedProtocols() As String()
    SupportedProtocols = libcurlProtocols()
End Property

Function Download(url As String, Optional toFile As String) As CCurlResponse
 
    Set Download = New CCurlResponse 'always a valid class to return
    'If Not INotify Is Nothing Then Set Me.INotify = INotify
     
    If Not initLib(errList) Then Exit Function
    
    If Not modDzTest.ActiveResponse Is Nothing Then
        If Not errList Is Nothing Then errList.Add "Limit Only one download active at a time right now you can fix this.."
        Exit Function
    End If
    
    Set modDzTest.ActiveResponse = Download 'one active instance
    
    With ActiveResponse
    
        Set .owner = Me
        If Not .Initilize(url, toFile) Then Exit Function
        
        vbcurl_easy_setopt hCurl, CURLOPT_URL, url
        vbcurl_easy_setopt hCurl, CURLOPT_WRITEDATA, ObjPtr(ActiveResponse)
        vbcurl_easy_setopt hCurl, CURLOPT_WRITEFUNCTION, AddressOf WriteFunction
        vbcurl_easy_setopt hCurl, CURLOPT_DEBUGFUNCTION, AddressOf DebugFunction
        vbcurl_easy_setopt hCurl, CURLOPT_VERBOSE, True
        If FileExists(caBundle) Then vbcurl_easy_setopt hCurl, CURLOPT_CAINFO, caBundle
        If hList <> 0 Then vbcurl_easy_setopt hCurl, CURLOPT_HTTPHEADER, hList
        If hForm <> 0 Then vbcurl_easy_setopt hCurl, CURLOPT_HTTPPOST, hForm
        
        .CurlReturnCode = vbcurl_easy_perform(hCurl)
        .Finalize
        
        'vbcurl_easy_reset hCurl  'wipe out entire config for next time
        
    End With
    
    Set ActiveResponse = Nothing
    
End Function


