VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Bind Event Handler to Object"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   3945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private REQ As Object
Private WithEvents SinkRSChange As SinkRSChange
Attribute SinkRSChange.VB_VarHelpID = -1

Private Sub Form_Load()
    Show
    
    Set REQ = CreateObject("MSXML2.XMLHTTP")
    Set SinkRSChange = New SinkRSChange
    REQ.onreadystatechange = SinkRSChange
    REQ.open "GET", "http://www.google.com", True
    Text1.Text = "Starting async GET" & vbNewLine
    REQ.send
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        Text1.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub

Private Sub SinkRSChange_onreadystatechange()
    Text1.Text = Text1.Text & vbNewLine & "readyState = " & CStr(REQ.readyState)
    If REQ.readyState = 4 Then
        Text1.Text = Text1.Text & vbNewLine & vbNewLine & REQ.responseText
    End If
End Sub
