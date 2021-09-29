VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Point to a ListView & Press F12"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmCloneListView"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get ListView Items"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParentHwnd As Long, ByVal ChildhWnd As Long, ByVal lpClassName As String, ByVal lpCaption As String) As Long

Private Const GWL_STYLE      As Long = (-16&)
Private Const KEY_DOWN       As Integer = &H8000
Private Const LVS_OWNERDATA  As Long = &H1000
Private Const MAX_CLASS_NAME As Long = 256

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As KeyCodeConstants) As Integer
Private Declare Function GetClassNameW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowLongW Lib "user32.dll" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal PointX As Long, ByVal PointY As Long) As Long

Private m_lngWndOver As Long
Private m_tP As POINTAPI
Private m_strClassName As String

Private Sub Command1_Click()
    
    Dim strPath As String, lngListViewWnd As Long
    
'    strPath = App.Path
'    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    
'    If IsWindow(Val(Text1.Text)) = 0 Then
'        lngListViewWnd = FindListView
'    Else
        lngListViewWnd = Val(Text1.Text)
'    End If
    
    Dim strColumns() As String, strItems() As String, strItem As String
    Dim strText As String, lngRow As Integer, lngCol As Integer
    
    Call GetListViewItems(lngListViewWnd, strColumns, strItems)
        
    If Not Not strItems Then
        For lngRow = 0 To UBound(strItems, 1&)
            strItem = strItem & strItems(lngRow, 0&)
            For lngCol = 1 To UBound(strItems, 2&)
                strItem = strItem & vbTab & vbTab & strItems(lngRow, lngCol)
                If lngCol = UBound(strItems, 2&) Then ' Item completed
                    strText = strText & strItem & vbNewLine
                    strItem = vbNullString
                End If
            Next
        Next
    End If
    
    Call MessageBoxW(Me.hwnd, StrPtr(strText), StrPtr(""), 0&)
    
End Sub

'Private Function FindListView(Optional ByRef lngMainWnd As Long) As Long
'    lngMainWnd = FindWindow(StrPtr("VeraCryptCustomDlg"), StrPtr("VeraCrypt"))
'    If lngMainWnd <> 0 Then FindListView = FindWindowEx(lngMainWnd, ByVal 0&, "SysListView32", vbNullString)
'End Function

Private Function GetClassName(ByVal hwnd As Long) As String
    SysReAllocStringLen VarPtr(GetClassName), , MAX_CLASS_NAME
    SysReAllocStringLen VarPtr(GetClassName), StrPtr(GetClassName), _
    GetClassNameW(hwnd, StrPtr(GetClassName), MAX_CLASS_NAME + 1&)
End Function

Private Function IfStr(ByVal blnExpression As Boolean, ByRef strTruePart As String, ByRef strFalsePart As String) As String
    If blnExpression Then IfStr = strTruePart Else IfStr = strFalsePart
End Function

Private Sub Timer1_Timer()
    If GetCursorPos(m_tP) Then
        m_lngWndOver = WindowFromPoint(m_tP.X, m_tP.Y)
        Select Case m_lngWndOver
            Case Me.hwnd, Text1.hwnd, Command1.hwnd
                Caption = "Point to a ListView & Press F12"
            Case Else
                m_strClassName = GetClassName(m_lngWndOver)
                Me.Caption = """" & m_strClassName & """  (&H" & Hex$(m_lngWndOver) & ")  " & IfStr(IsWindowUnicode(m_lngWndOver), "[Unicode]", "[ANSI]")
                Select Case m_strClassName
                    Case "SysListView32", "ListViewWndClass", "ListView20WndClass"
                        If GetAsyncKeyState(vbKeyF12) And KEY_DOWN Then
                            Text1.Text = m_lngWndOver
                            Command1_Click
                        End If
                End Select
        End Select
    End If
End Sub














