VERSION 5.00
Begin VB.Form frmWhoIs 
   Caption         =   "IP WhoIs"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14205
   Icon            =   "frmWhoIs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   14205
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Tag             =   "DrpDwn"
      Top             =   0
      Width           =   3975
   End
   Begin VB.ListBox lstAddress 
      Height          =   2595
      ItemData        =   "frmWhoIs.frx":0742
      Left            =   120
      List            =   "frmWhoIs.frx":0744
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ComboBox cboHost 
      Height          =   315
      Left            =   4320
      TabIndex        =   6
      Text            =   "whois.arin.net"
      Top             =   0
      Width           =   2655
   End
   Begin VB.CheckBox chkIPv6 
      Caption         =   "Use IPv6"
      Enabled         =   0   'False
      Height          =   200
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox chkRef 
      Caption         =   "Use Referral"
      Height          =   200
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmWhoIs.frx":0746
      Top             =   960
      Width           =   14055
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "SEND"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image AddrDrop 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3960
      Picture         =   "frmWhoIs.frx":074E
      Tag             =   "DrpDwn"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   6480
      Width           =   14055
   End
End
Attribute VB_Name = "frmWhoIs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents WhoIs As SimpleSock
Attribute WhoIs.VB_VarHelpID = -1

Private WhoIsServer As String
Private WhoIsPort As Long
Private WhoIsCmd As String
Private Const WhoISRefer As String = "ReferralServer:"
Private Const gAppName As String = "IPWhois"
Private Const LB_SELECTSTRING = &H18C
Private Const LB_SETTOPINDEX = &H197
Private Const Key_Left = &H25
Private Const Key_Up = &H26
Private Const Key_Right = &H27
Private Const Key_Down = &H28

Dim CursorKeyFlg As Integer
Dim Int_Address() As String

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Sub QueryEnd()
    Dim Temp$
    Dim M%, N%
    On Error Resume Next
    Err.Clear
    'Look for referral if chkRef checked
    If chkRef Then
    N% = InStr(1, txtOut.Text, WhoISRefer, vbTextCompare)
        If N% > 0 Then
            M% = InStr(N%, txtOut.Text, vbCrLf, vbTextCompare)
            If M% > N% Then
                N% = InStr(N%, txtOut.Text, "://", vbTextCompare)
                N% = N% + 3
                If M% > N% Then
                    Temp$ = Mid$(txtOut.Text, N%, M% - N%)
                    N% = InStr(1, Temp$, ":", vbTextCompare)
                    If N% = 0 Then
                        WhoIsServer = Temp$
                        WhoIsPort = 43
                        WhoIsCmd = txtIP.Text
                        txtOut.Text = ""
                        lblStatus.Caption = "Referred to: " & Temp$
                        Call WhoIs.TCPConnect(WhoIsServer, WhoIsPort)
                   ElseIf N% > 0 And M% > N% Then
                        WhoIsServer = Mid$(Temp$, 1, N% - 1)
                        WhoIsPort = Mid$(Temp$, N% + 1)
                        WhoIsCmd = txtIP.Text
                        txtOut.Text = ""
                        lblStatus.Caption = "Referred to: " & Temp$
                        WhoIs.TCPConnect WhoIsServer, WhoIsPort
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Function GetSettings(sKey As String) As String
    GetSettings = GetSetting(gAppName, "Settings", sKey, "")
End Function

Private Sub SaveSettings(sKey As String, sValue As String)
    SaveSetting gAppName, "Settings", sKey, sValue
End Sub

Private Sub AddrDrop_Click()
    lstAddress.Visible = Not lstAddress.Visible
End Sub

Private Sub chkIPv6_Click()
    Dim M%, N%
    If chkIPv6.Value = 1 Then
        For N% = 0 To UBound(Int_Address)
            If InStr(Int_Address(N%), ":") Then GoTo v6_found
        Next N%
        MsgBox "IPv6 not Enabled!"
        chkIPv6.Value = 0
        Exit Sub
v6_found:
        M% = InStr(Int_Address(N%), "%")
        If M% Then
            txtIP = Left$(Int_Address(N%), M% - 1)
        Else
            txtIP = Int_Address(N%)
        End If
        WhoIs.IPvFlg = 6 'Set IP version
    Else
        Do Until InStr(Int_Address(N%), ".")
            N% = N% + 1
        Loop
        txtIP = Int_Address(N%)
        WhoIs.IPvFlg = 4
    End If
End Sub

Private Sub cmdSend_Click()
    Dim Result As Long
    frmWhoIs.MousePointer = vbHourglass  'Verify address
    If Len(WhoIs.GetIPFromHost(txtIP.Text, "0")) = 0 Then
        frmWhoIs.MousePointer = vbDefault
        MsgBox "You have entered an Invalid Address!"
        Exit Sub
    End If
    txtOut.Text = ""
    lstAddress.Visible = False
    txtIP.SelStart = 0
    txtIP.SelLength = Len(txtIP)
    WhoIsServer = cboHost.Text
    WhoIsPort = 43
    If cboHost.ListIndex = 0 Then
        WhoIsCmd = "n " & txtIP.Text    'arin request
    Else
        WhoIsCmd = txtIP.Text           'non-arin request
    End If
    lblStatus.Caption = "Whois: " & WhoIsServer
    WhoIs.TCPConnect WhoIsServer, WhoIsPort
    'Check if address already in list
    Result = SendMessage(lstAddress.hWnd, LB_SELECTSTRING, -1, txtIP.Text)
    If Result < 0 Then
        lstAddress.AddItem txtIP.Text 'add to list
        Debug.Print "Address added to list!"
    End If
End Sub


Private Sub Form_Activate()
    Dim Test_Address As String
    Dim M%, N%
    'Find and save all internal addresses
    Test_Address = WhoIs.GetIPFromHost(WhoIs.GetLocalHostName, "")
    Int_Address = Split(Test_Address, Chr$(0))
    For N% = 0 To UBound(Int_Address) 'Check if IPv6 enabled
        If InStr(Int_Address(N%), ":") Then
            If Left$(Int_Address(N%), 5) <> "fe80:" Then
                chkIPv6.Enabled = True
                Exit For
            End If
        End If
    Next N%
    If chkIPv6.Enabled = True And GetSettings("IPv6") = "1" Then
        chkIPv6.Value = 1
        WhoIs.IPvFlg = 6 'Set IP version
    Else
        WhoIs.IPvFlg = 4
        For N% = 0 To UBound(Int_Address)
            If InStr(Int_Address(N%), ".") Then
                txtIP = Int_Address(N%)
                Exit For
            End If
        Next N%
    End If
End Sub

Private Sub Form_Click()
    'FlashBox.Show 0
    DoEvents
End Sub

Private Sub Form_Load()
    Set WhoIs = New SimpleSock
    With cboHost
        .AddItem "whois.arin.net"
        .AddItem "whois.apnic.net"
        .AddItem "whois.ripe.net"
        .AddItem "whois.lacnic.net"
        .AddItem "whois.afrinic.net"
    End With
    cboHost.ListIndex = 0
End Sub


Private Sub Form_Resize()
    'FlashBox.Show 0
    DoEvents
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call SaveSettings("IPv6", CStr(chkIPv6.Value))
End Sub

Private Sub lstAddress_Click()
    If lstAddress.ListIndex > -1 Then
        txtIP.Text = lstAddress.List(lstAddress.ListIndex)
    Else
        Call SendMessage(lstAddress.hWnd, LB_SETTOPINDEX, 0, "")
    End If
    If CursorKeyFlg Then
        CursorKeyFlg = False
    Else
        lstAddress.Visible = False
    End If
End Sub

Private Sub lstAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Key_Down, Key_Up, Key_Right, Key_Left
            CursorKeyFlg = True
    End Select
End Sub


Private Sub lstAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstAddress.Visible = False
        KeyAscii = 0
        Call txtIP_KeyPress(13)
    End If
End Sub


Private Sub lstAddress_LostFocus()
    If ActiveControl.Tag <> "DrpDwn" Then
        If lstAddress.ListIndex > 0 Then
            lstAddress.Visible = False
        End If
    End If
End Sub


Private Sub txtIP_Change()
    Dim Result As Long
    Result = SendMessage(lstAddress.hWnd, LB_SELECTSTRING, -1, txtIP.Text)
    If Result >= 0 Then
        Debug.Print "Item " + Str$(Result) + " Selected!"
    Else
        CursorKeyFlg = True
        lstAddress.ListIndex = -1
    End If
End Sub

Private Sub txtIP_Click()
    txtIP.SelStart = 0
    txtIP.SelLength = Len(txtIP)
    lstAddress.Visible = True
End Sub


Private Sub txtIP_GotFocus()
    txtIP.SelStart = 0
    txtIP.SelLength = Len(txtIP)
End Sub


Private Sub txtIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdSend_Click
        cmdSend.SetFocus
    End If
End Sub







Private Sub txtIP_LostFocus()
    If ActiveControl.Tag <> "DrpDwn" Then
        If lstAddress.ListIndex > 0 Then
            lstAddress.Visible = False
        End If
    End If
End Sub

Private Sub WhoIs_CloseSck()
    frmWhoIs.MousePointer = vbDefault
    Call QueryEnd
End Sub


Private Sub WhoIs_Connect()
    Debug.Print "Query = " & WhoIsCmd
    WhoIs.sOutBuffer = WhoIsCmd & vbCrLf
    WhoIs.TCPSend
End Sub


Private Sub WhoIs_DataArrival(ByVal bytesTotal As Long)
    Dim strdata As String
    On Error Resume Next
    WhoIs.RecoverData
    strdata = WhoIs.sInBuffer
    If Len(strdata) <= 0 Then Exit Sub
    If Left(strdata, 1) = vbLf Then
        strdata = Mid(strdata, 2)
        If Len(strdata) <= 0 Then Exit Sub
    End If
    strdata = Replace(strdata, vbLf, vbCrLf)
    txtOut.Text = txtOut.Text & strdata
    strdata = Trim(strdata)
    If Len(strdata) <= 0 Then Exit Sub
End Sub


Private Sub WhoIS_Error(ByVal Number As Long, Description As String, ByVal Source As String)
    Beep
    WhoIs_CloseSck
'    WhoIs.LocalPort = 0
    lblStatus.Caption = "Whois: Error " & Description
    frmWhoIs.MousePointer = vbDefault
End Sub


