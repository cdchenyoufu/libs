VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10455
   LinkTopic       =   "Form2"
   ScaleHeight     =   9255
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopyList 
      Caption         =   "Copy"
      Height          =   375
      Left            =   630
      TabIndex        =   12
      Top             =   8280
      Width           =   1365
   End
   Begin VB.TextBox txtReferrer 
      Height          =   285
      Left            =   765
      TabIndex        =   11
      Top             =   540
      Width           =   8790
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   420
      Left            =   7110
      TabIndex        =   9
      Top             =   8325
      Width           =   1140
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   9630
      TabIndex        =   8
      Top             =   180
      Width           =   465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download"
      Height          =   420
      Left            =   8460
      TabIndex        =   7
      Top             =   8325
      Width           =   1500
   End
   Begin MSComctlLib.ProgressBar pb2 
      Height          =   285
      Left            =   5670
      TabIndex        =   6
      Top             =   1125
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   630
      TabIndex        =   5
      Top             =   1125
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   585
      TabIndex        =   4
      Top             =   5985
      Width           =   9465
   End
   Begin VB.TextBox txtUrls 
      Height          =   4335
      Left            =   585
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1440
      Width           =   9420
   End
   Begin VB.TextBox txtSave2 
      Height          =   330
      Left            =   765
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   135
      Width           =   8790
   End
   Begin VB.Label Label3 
      Caption         =   "Referrer"
      Height          =   285
      Left            =   45
      TabIndex        =   10
      Top             =   630
      Width           =   690
   End
   Begin VB.Label Label2 
      Caption         =   "Urls"
      Height          =   240
      Left            =   -45
      TabIndex        =   3
      Top             =   1485
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Save to:"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Dim WithEvents curl As CCurlDownload
Attribute curl.VB_VarHelpID = -1
 Dim abort As Boolean
 
Private Sub cmdAbort_Click()
     abort = True
End Sub

Private Sub cmdBrowse_Click()

    Dim fso As Object
    On Error Resume Next
    List1.Clear
    Set fso = CreateObject("dzrt.CFileSystem3")
    If Err.Number <> 0 Then
        cmdBrowse.Enabled = False
        List1.AddItem "dzrt.dll not found browse file disabled"
    Else
        Const sf_DESKTOP As Long = &H0
        txtSave2 = fso.dlg.FolderDialog2(fso.GetSpecialFolder(sf_DESKTOP))
    End If
    
End Sub

Private Sub cmdCopyList_Click()
    Dim i, t
    On Error Resume Next
    For i = 0 To List1.ListCount
        t = t & List1.List(i) & vbCrLf
    Next
    Clipboard.Clear
    Clipboard.SetText t
End Sub

Private Sub curl_Init(obj As CCurlResponse)
    On Error Resume Next
    If obj.DownloadLength > 0 Then pb2.Max = obj.DownloadLength
End Sub

Private Sub curl_Progress(obj As CCurlResponse)
    On Error Resume Next
    pb2.value = obj.BytesReceived
    If abort Then
        List1.AddItem "Aborting at user request..."
        obj.abort = True
    End If
End Sub

Private Sub curl_Complete(obj As CCurlResponse)
    pb2.value = 0
    List1.AddItem "Download complete resp code: " & obj.ResponseCode & " time: " & obj.TotalTime & " " & obj.url
End Sub


Private Sub Command1_Click()
    
    Dim fname As String, i As Long
    
    List1.Clear
    If Not FolderExists(txtSave2) Then
        List1.AddItem "Output folder not exist"
        Exit Sub
    End If
    
    Dim tmp() As String, u, resp As CCurlResponse
    
    tmp = Split(txtUrls, vbCrLf)
    If AryIsEmpty(tmp) Then Exit Sub
    
    pb.Max = UBound(tmp) + 1
    pb.value = 0
    
    On Error Resume Next
    Set curl = New CCurlDownload
    If Len(txtReferrer) > 0 Then curl.Referrer = txtReferrer
    curl.Useragent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:83.0) Gecko/20100101 Firefox/83.0"
    
    For Each u In tmp
        u = Trim(u)
        If Len(u) > 0 Then
            i = 1
            fname = WebFileNameFromPath(u)
            While FileExists(txtSave2 & "\" & fname)
                fname = WebFileNameFromPath(u) & "_" & i
                i = i + 1
            Wend
            List1.AddItem "Downloading: " & u & " -> " & fname
            Set resp = curl.Download(CStr(u), txtSave2 & "\" & fname)
            List1.List(List1.ListCount) = List1.List(List1.ListCount) & " -> " & resp.ResponseCode & " Bytes: " & resp.BytesReceived
        End If
        If abort Then Exit For
        pb.value = pb.value + 1
    Next
    
    pb.value = 0
    pb2.value = 0
    
End Sub

Private Sub txtSave2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim p As String
    p = Data.Files(1)
    If FolderExists(p) Then txtSave2 = p
End Sub
