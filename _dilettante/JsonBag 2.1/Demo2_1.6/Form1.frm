VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "JsonBag Demo 2"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11010
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtJsonBagDump 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   420
      Width           =   4515
   End
   Begin VB.TextBox txtJson 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   420
      Width           =   4455
   End
   Begin VB.Label lblJsonBagDump 
      Caption         =   "Dump of JsonBag"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   60
      Width           =   3015
   End
   Begin VB.Label lblJson 
      Caption         =   "JSON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GapHorizontal As Single
Private GapVertical As Single

Private Sub DumpText(ByVal Text As String)
    With txtJsonBagDump
        .SelStart = &H7FFF
        .SelText = Text
        .SelText = vbNewLine
    End With
End Sub

Private Sub HomeDump()
    txtJsonBagDump.SelStart = 0
End Sub

Private Sub JbDump(ByVal Name As String, ByVal JB As JsonBag, Optional ByVal Level As Integer)
    'Recursive dump.  Hazard: control characters or non-ANSI characters
    'within an item's value are not escaped for display.
    Dim IndentLevel As String
    Dim IndentNextLevel As String
    Dim ItemIndex As Long
    Dim Item As Variant
    
    IndentLevel = Space$(Level * 2)
    IndentNextLevel = Space$((Level + 1) * 2)
    DumpText IndentLevel & Name & "(" & IIf(JB.IsArray, "Array", "Object") & "):"
    For ItemIndex = 1 To JB.Count
        If TypeOf JB.Item(ItemIndex) Is JsonBag Then
            Set Item = JB.Item(ItemIndex)
            If JB.IsArray Then
                JbDump "", Item, Level + 1
            Else
                JbDump JB.Name(ItemIndex), Item, Level + 1
            End If
        Else
            Item = JB.Item(ItemIndex)
            If JB.IsArray Then
                DumpText IndentNextLevel & """" & Item & """"
            Else
                DumpText IndentNextLevel & JB.Name(ItemIndex) & ": """ & Item & """"
            End If
        End If
    Next
End Sub

Private Sub Form_Load()
    Dim F As Integer
    Dim JsonData As String
    Dim JB As JsonBag
    
    GapHorizontal = lblJson.Left
    GapVertical = lblJson.Top
    
    F = FreeFile(0)
    Open "JsonSample.txt" For Input As #F
    JsonData = Input$(LOF(F), #F)
    Close #F
    txtJson.Text = JsonData

    Set JB = New JsonBag
    JB.JSON = JsonData
    Show
    JbDump "*anonymous outer JsonBag*", JB
    HomeDump
    
    'Safe accessor syntax:
    MsgBox "JB.Item(""web-app"").Item(""servlet"")(1).Item(""init-param"").Item(""templateProcessorClass"")" _
         & vbNewLine & vbNewLine _
         & "= """ _
         & JB.Item("web-app").Item("servlet")(1).Item("init-param").Item("templateProcessorClass") & """"
    
    'Compact accessor syntax, but subject to error if the IDE changes the case
    'of the names of named items (in brackets [] here).  We get lucky here,
    'since there are no names used that the IDE will re-case on us:
    MsgBox "jb![web-app]![servlet](1)![init-param]![templateProcessorClass]" _
         & vbNewLine & vbNewLine _
         & "= """ _
         & JB![web-app]![servlet](1)![init-param]![templateProcessorClass] & """"
End Sub

Private Sub Form_Resize()
    Dim TxtWidth As Single
    Dim TxtHeight As Single
    
    If WindowState <> vbMinimized Then
        TxtWidth = (ScaleWidth - 3# * GapHorizontal) / 2#
        TxtHeight = ScaleHeight - (2# * GapVertical _
                                     + lblJson.Height)
        With txtJson
            .Move GapHorizontal, lblJson.Height + GapVertical, TxtWidth, TxtHeight
            txtJsonBagDump.Move 2# * GapHorizontal + .Width, .Top, TxtWidth, TxtHeight
            lblJsonBagDump.Left = txtJsonBagDump.Left
        End With
    End If
End Sub

