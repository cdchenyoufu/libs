VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutPut 
      Height          =   2955
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "test.frx":0000
      Top             =   720
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   180
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   0
      TabIndex        =   1
      Top             =   3900
      Width           =   6675
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Text            =   "http://sandsprite.com"
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents curl As CCurlDownload
Attribute curl.VB_VarHelpID = -1


 

'uses libcurl to download files directly to the toFile location no cache used.

'  file:   vblibcurl.dll (std C dll)
'  file    vblibcurl.tlb (api declares and enums for above) - see modDeclares for all enums and partial declares
'  author: Jeffrey Phillips
'  date:   2.28.2005
 
'dzzie: 11.13.20
'   initLib() to find/load C dll dependencies on the fly from different paths
'   removed tlb references w/modDeclares.bas (all enums covered but not all api declares written yet)
'   added higher level framework around low level api
'   file progress, response object, abort, download to memory only
'   add/remove/overwrite headers
'   add form elements and file uploads
'   version info

Private Sub Form_Load()
    Set curl = New CCurlDownload
End Sub

Private Sub curl_Header(obj As CCurlResponse, ByVal msg As String)
     List1.AddItem "header: " & msg
End Sub

Private Sub curl_InfoMsg(obj As CCurlResponse, ByVal info As curl_infotype, ByVal msg As String)
    If info = CURLINFO_HEADER_OUT Then
        txtSent = msg
    Else
        'List1.AddItem "info: " & info & " (" & info2Text(info) & ") " & msg
        List1.AddItem "info: " & info & " " & msg
    End If
End Sub

Private Sub curl_Init(obj As CCurlResponse)
    On Error Resume Next
    If obj.DownloadLength > 0 Then pb.Max = obj.DownloadLength
End Sub

Private Sub curl_Progress(obj As CCurlResponse)
    On Error Resume Next
    pb.Value = obj.BytesReceived
    If abort Then
        List1.AddItem "Aborting at user request..."
        obj.abort = True
    End If
End Sub

Private Sub curl_Complete(obj As CCurlResponse)
    List1.AddItem "Download complete resp code: " & obj.ResponseCode & " time: " & obj.TotalTime
End Sub


 

Private Sub Command1_Click()

    On Error Resume Next
    
    Dim resp As CCurlResponse, e
    Dim ret As CURLcode
    Dim totalTimeout As Long, connectTimeout As Long
    
    List1.Clear
    txtOutPut = Empty
    abort = False
    
    totalTimeout = CLng(txtTimeout)
    If Err.Number <> 0 Then
        List1.AddItem "Invalid total timeout"
        Exit Sub
    End If
    
    connectTimeout = CLng(txtConnectTimeout)
    If Err.Number <> 0 Then
        List1.AddItem "Invalid connect timeout"
        Exit Sub
    End If

    On Error GoTo hell
    Set curl = New CCurlDownload
    
    If curl.errList.Count > 0 Then
        List1.AddItem "Error initilizing libcurl"
        For Each e In curl.errList
            List1.AddItem e
        Next
    End If
    
    Set curl.errList = New Collection
    curl.Configure "My Useragent", , totalTimeout, connectTimeout
    
    'curl.Referrer = "http://test.edition/yaBoy?" & curl.escape("this is my escape test!!")
    'curl.Cookie = "monster:true;"
    
    'curl.AddHeader "X-MyHeader: Works"
    'curl.AddHeader "X-LibCurl: Rocks"
    'curl.AddHeader "Accept:" 'this will remove the automatic Accept header we could also override it here
    'curl.AddHeader Array("X-Ary1: 1", "X-Ary2: 2")
    
    'todo a simple post:
    'vbcurl_easy_setopt curl.hCurl, CURLOPT_POSTFIELDS, "m-address=your@mail.com"

'    If chkPost.Value = 1 Then
'        'try: https://postman-echo.com/post
'         ret = curl.AddFormElement("test", "taco breath")
'         List1.AddItem "Add form field test: " & curlCode2Text(ret)
'
'         ret = curl.AddFormFileUpload("fart", App.Path & "\readme.txt")
'         List1.AddItem "Add form field fart: " & curlCode2Text(ret)
'    End If
    
    
'    If Len(txtSaveAs) = 0 Then
        Set resp = curl.Download(Text1.Text)
        txtOutPut = resp.dump & vbCrLf & vbCrLf & resp.memFile.asString
''    Else
'        Set resp = curl.Download(Text1. , txtSaveAs)
'        txtOutPut = resp.dump
'        'List1.AddItem "MD5: " & hash.HashFile(txtSaveAs)
'    'End If
    
    'List1.AddItem "Download 2 same handle received bytes: " & curl.Download(cboUrl.Text).BytesReceived
    
    Exit Sub
hell:
    List1.AddItem "Error: " & Err.Description
End Sub

