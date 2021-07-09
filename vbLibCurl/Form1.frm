VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "vbLibCurl Demo"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13365
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   13365
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUA 
      Height          =   360
      Left            =   1740
      TabIndex        =   23
      Text            =   "vbLibCurl Test Edition"
      Top             =   1740
      Width           =   3315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bulk Download"
      Height          =   465
      Left            =   11610
      TabIndex        =   21
      Top             =   1575
      Width           =   1590
   End
   Begin VB.CheckBox chkPost 
      Caption         =   "POST demo"
      Height          =   285
      Left            =   8775
      TabIndex        =   20
      Top             =   1485
      Width           =   1995
   End
   Begin VB.ComboBox cboUrl 
      Height          =   360
      Left            =   1170
      TabIndex        =   19
      Top             =   135
      Width           =   11130
   End
   Begin VB.TextBox txtSent 
      Height          =   2220
      Left            =   945
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   4860
      Width           =   12300
   End
   Begin VB.TextBox txtConnectTimeout 
      Height          =   360
      Left            =   7650
      TabIndex        =   15
      Text            =   "15"
      Top             =   1395
      Width           =   465
   End
   Begin VB.TextBox txtTimeout 
      Height          =   360
      Left            =   4365
      TabIndex        =   13
      Text            =   "0"
      Top             =   1395
      Width           =   510
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   420
      Left            =   10395
      TabIndex        =   11
      Top             =   1035
      Width           =   1140
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   5985
      TabIndex        =   10
      Top             =   1080
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   12375
      TabIndex        =   8
      Top             =   630
      Width           =   645
   End
   Begin VB.TextBox txtOutput 
      Height          =   3300
      Left            =   990
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   7515
      Width           =   12300
   End
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   945
      TabIndex        =   4
      Top             =   2475
      Width           =   12345
   End
   Begin VB.CommandButton cmdDl 
      Caption         =   "Download"
      Height          =   465
      Left            =   11700
      TabIndex        =   3
      Top             =   1035
      Width           =   1500
   End
   Begin VB.TextBox txtSaveAs 
      Height          =   360
      Left            =   1215
      TabIndex        =   2
      Top             =   540
      Width           =   11040
   End
   Begin VB.Label Label10 
      Caption         =   "User Agent"
      Height          =   315
      Left            =   180
      TabIndex        =   22
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Label Label9 
      Caption         =   "Sent"
      Height          =   330
      Left            =   90
      TabIndex        =   17
      Top             =   4995
      Width           =   600
   End
   Begin VB.Label Label8 
      Caption         =   "Note: VB file commands have 2gb max file size limit switch to API if necessary"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1020
      TabIndex        =   16
      Top             =   2100
      Width           =   11220
   End
   Begin VB.Label Label7 
      Caption         =   "Connect timeout    (s)"
      Height          =   285
      Left            =   5535
      TabIndex        =   14
      Top             =   1440
      Width           =   3120
   End
   Begin VB.Label Label6 
      Caption         =   "Total DL Timeout    (s)"
      Height          =   285
      Left            =   2115
      TabIndex        =   12
      Top             =   1440
      Width           =   3390
   End
   Begin VB.Label Label5 
      Caption         =   "Empty SaveAs path = mem only"
      Height          =   285
      Left            =   1260
      TabIndex        =   9
      Top             =   990
      Width           =   4155
   End
   Begin VB.Label Label4 
      Caption         =   "Output"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   7470
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Debug"
      Height          =   240
      Left            =   45
      TabIndex        =   5
      Top             =   2430
      Width           =   690
   End
   Begin VB.Label Label2 
      Caption         =   "Save As"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Implements ICurlProgress
Dim WithEvents curl As CCurlDownload
Attribute curl.VB_VarHelpID = -1


Private abort As Boolean
Dim fso As Object 'if they have dzrt.CFileSystem3 then browse button is available..

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


Private Sub cmdAbort_Click()
    abort = True
End Sub

Private Sub Command1_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub curl_Header(obj As CCurlResponse, ByVal msg As String)
     List1.AddItem "header: " & msg
End Sub

Private Sub curl_InfoMsg(obj As CCurlResponse, ByVal info As curl_infotype, ByVal msg As String)
    If info = CURLINFO_HEADER_OUT Then
        txtSent = msg
    Else
        List1.AddItem "info: " & info & " (" & info2Text(info) & ") " & msg
    End If
End Sub

Private Sub curl_Init(obj As CCurlResponse)
    On Error Resume Next
    If obj.DownloadLength > 0 Then pb.Max = obj.DownloadLength
End Sub

Private Sub curl_Progress(obj As CCurlResponse)
    On Error Resume Next
    pb.value = obj.BytesReceived
    If abort Then
        List1.AddItem "Aborting at user request..."
        obj.abort = True
    End If
End Sub

Private Sub curl_Complete(obj As CCurlResponse)
    pb.value = 0
    List1.AddItem "Download complete resp code: " & obj.ResponseCode & " time: " & obj.TotalTime
End Sub


Private Sub cmdBrowse_Click()
    On Error Resume Next 'late bound if found
    Const sf_DESKTOP As Long = &H0
    txtSaveAs = fso.dlg.OpenDialog(fso.GetSpecialFolder(sf_DESKTOP))
End Sub

Private Sub cmdDl_Click()

    On Error Resume Next
    
    Dim resp As CCurlResponse
    Dim ret As CURLcode
    Dim totalTimeout As Long, connectTimeout As Long
    
    List1.Clear
    txtOutput = Empty
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
    curl.Configure txtUA, , totalTimeout, connectTimeout, List1
    'curl.Referrer = "http://test.edition/yaBoy?" & curl.escape("this is my escape test!!")
    'curl.Cookie = "monster:true;"
    
    'curl.AddHeader "X-MyHeader: Works"
    'curl.AddHeader "X-LibCurl: Rocks"
    curl.AddHeader "Accept:" 'this will remove the automatic Accept header we could also override it here
    'curl.AddHeader Array("X-Ary1: 1", "X-Ary2: 2")
    
    'todo a simple post:
    'vbcurl_easy_setopt curl.hCurl, CURLOPT_POSTFIELDS, "m-address=your@mail.com"

    If chkPost.value = 1 Then
        'try: https://postman-echo.com/post
         ret = curl.AddFormElement("test", "taco breath")
         List1.AddItem "Add form field test: " & curlCode2Text(ret)
    
         ret = curl.AddFormFileUpload("fart", App.path & "\readme.txt")
         List1.AddItem "Add form field fart: " & curlCode2Text(ret)
    End If
    
    
    If Len(txtSaveAs) = 0 Then
        Set resp = curl.Download(cboUrl.Text)
        txtOutput = resp.dump & vbCrLf & vbCrLf & resp.memFile.asString
    Else
        Set resp = curl.Download(cboUrl.Text, txtSaveAs)
        txtOutput = resp.dump
        'List1.AddItem "MD5: " & hash.HashFile(txtSaveAs)
    End If
    
    'List1.AddItem "Download 2 same handle received bytes: " & curl.Download(cboUrl.Text).BytesReceived
    
    Exit Sub
hell:
    List1.AddItem "Error: " & Err.Description
End Sub

Private Sub Form_Load()

    If Not initLib(List1) Then
        cmdDl.Enabled = False
        List1.AddItem "This demo requires vblibcurl.dll and libcurl.dll"
        List1.AddItem "https://sourceforge.net/projects/libcurl-vb/files/libcurl-vb/libcurl.vb%201.01/"
    Else
        List1.AddItem "libCurl Version: " & libCurlVersion()
        List1.AddItem "Protocols: " & Join(libcurlProtocols(), ",")
        List1.AddItem "Zlib version: " & libZVersion()
        List1.AddItem "SSL Version: " & sslVersion()
    End If
    
    txtUA = GetSetting("vbLibCurl", "Settings", "UA", "vbLibCurl Demo")
    
    cboUrl.AddItem "http://sandsprite.com/tools.php"
    cboUrl.AddItem "https://postman-echo.com/get"
    cboUrl.AddItem "https://postman-echo.com/post"
    cboUrl.ListIndex = 0
    
    'membuf test
    '--------------------------------
    Dim b As New CMemBuffer, b1() As Byte, b2() As Byte, b3() As Byte
    
    b1 = StrConv("0123456789ABCDEF", vbFromUnicode, &H409)
    b2 = StrConv(String(16, "B"), vbFromUnicode, &H409)
    b3 = StrConv(String(16, "C"), vbFromUnicode, &H409)
    
    'b.appendBuf b1
    'b.appendBuf b2
    'b.appendBuf b3
    'txtOutput = "CMemBufTest: " & vbCrLf & HexDump(b.binData())
    
    b.memAppendBuf VarPtr(b1(0)), UBound(b1) + 1
    b.memAppendBuf VarPtr(b2(0)), UBound(b2) + 1
    b.memAppendBuf VarPtr(b3(0)), UBound(b3) + 1
    txtOutput = "CMemBufTest: " & vbCrLf & HexDump(b.binData()) & vbCrLf

    'escape unescape test
    '--------------------------------
    Dim c As New CCurlDownload, tmp As String, org As String, dec As String
    org = "this.is#my test!?%"
    tmp = c.escape(org)
    dec = c.unescape(tmp)
    txtOutput = txtOutput & vbCrLf & "Escape/Unescape test = " & (org = dec) & vbCrLf & HexDump(tmp)
    '--------------------------------
    
    On Error Resume Next
    Set fso = CreateObject("dzrt.CFileSystem3")
    If Err.Number <> 0 Then
        cmdBrowse.Enabled = False
        List1.AddItem "dzrt.dll not found browse file disabled"
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
       Call SaveSetting("vbLibCurl", "Settings", "UA", txtUA)
End Sub

Private Sub List1_DblClick()
    On Error Resume Next
    MsgBox List1.List(List1.ListIndex)
End Sub
