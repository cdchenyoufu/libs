VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "JsonBag to TreeView"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   6810
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetpath 
      Caption         =   "Getpath"
      Height          =   315
      Left            =   9780
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtGetPath 
      Height          =   330
      Left            =   6300
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   5220
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtPath 
      Height          =   330
      Left            =   420
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   60
      Width           =   4635
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   1020
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   4471
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Path"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3660
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   "Array"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":02BA
            Key             =   "Object"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":03BC
            Key             =   "Element"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":04BE
            Key             =   "Property"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private curBag As JsonBag
Private samplePath As String

'recursive
Private Sub TreeTheNode( _
    ByVal Name As Variant, _
    ByVal Item As Variant, _
    Optional ByVal Parent As ComctlLib.Node)

    Dim ImageKey As String
    Dim NewNode As ComctlLib.Node
    Dim i As Long
    Dim ItemAsText As String
    Dim Text As String


    With TreeView1.Nodes
        If VarType(Item) = vbObject Then
            If Item.IsArray Then
                ImageKey = "Array"
            Else
                ImageKey = "Object"
            End If
            If VarType(Name) <> vbString Then
                Name = "(" & CStr(Name) & ")"
            End If
            If Parent Is Nothing Then
                Set NewNode = .Add(, , Name, Name, ImageKey)
            Else
                Set NewNode = .Add(Parent.Key, tvwChild, Parent.Key & "\" & Name, Name, ImageKey)
            End If
            For i = 1 To Item.Count
                TreeTheNode Item.Name(i), Item.Item(i), NewNode
            Next
            NewNode.Expanded = True
        Else 'Value.
            Select Case VarType(Item)
                Case vbNull
                    ItemAsText = "Null"
                Case vbString
                    If InStr(Item, """") > 0 Then
                        ItemAsText = "'" & Item & "'"
                    Else
                        ItemAsText = """" & Item & """"
                    End If
                Case Else
                    ItemAsText = CStr(Item)
            End Select
            If VarType(Name) = vbString Then
                ImageKey = "Property"
                Text = Name & ": " & ItemAsText
            Else
                ImageKey = "Element"
                Name = "(" & CStr(Name) & ")"
                Text = Name & " " & ItemAsText
            End If
            If Parent Is Nothing Then
                .Add , , Name, Text, ImageKey
            Else
                .Add Parent.Key, tvwChild, Parent.Key & "\" & Name, Text, ImageKey
            End If
        End If
    End With
End Sub

Private Sub cmdGetpath_Click()

    On Error Resume Next
    Dim v As Variant
    
    If Len(txtGetPath) = 0 Then
        MsgBox "Enter an object or value path to retrieve (array indexes not supported yet in parser)"
        Exit Sub
    End If
    
    If curBag.getPath(txtGetPath, v) Then
        If IsObject(v) Then
            MsgBox "GetPath returned a " & TypeName(v)
        Else
            MsgBox "GetPath returned a value: " & v
        End If
    Else
        MsgBox "GetPath returned false Error: " & Err.Description
    End If
    
End Sub

Private Sub cmdLoad_Click()
    
    On Error GoTo hell
    
    Dim v() As Variant, dump() As String, i As Long
    Dim subBag As JsonBag
    
    TreeView1.Nodes.Clear
    Set curBag = New JsonBag
    Me.Caption = "Loaded from " & txtPath
    
    If curBag.fromFile(txtPath) Then
        TreeTheNode "[Root]", curBag
        
        If txtPath = samplePath Then
            txtGetPath.Text = "sample.story.title"
            If curBag.getPath("sample.text", subBag) Then
                v = subBag.toArray()
                ReDim dump(UBound(v))
                For i = 0 To UBound(v)
                    If IsObject(v(i)) Then
                        dump(i) = TypeName(v(i))
                    Else
                        dump(i) = "Value: " & v(i)
                    End If
                Next
                MsgBox "Sample.text toArray() dump: " & vbCrLf & vbTab & Join(dump, vbCrLf & vbTab)
            End If
        End If
 
        
    Else
        MsgBox "fromFile failed"
    End If
    
    Exit Sub
    
hell:
    MsgBox Err.Description, vbExclamation
    
End Sub

Private Sub Form_Load()
    Set curBag = New JsonBag
    samplePath = App.path & "\sample.json"
    txtPath = samplePath
    Me.Caption = "Loaded from string"
    curBag.JSON = "{'a':'quotes\'here',""B"":""quotes\""here""}"
    TreeTheNode "[Root]", curBag
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        TreeView1.Move 0, txtPath.Height + txtPath.Top + 40, ScaleWidth, ScaleHeight
    End If
End Sub

Private Sub txtPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If FileExists(Data.Files(1)) Then txtPath = Data.Files(1)
End Sub

Private Function FileExists(path) As Boolean
  On Error GoTo hell
    
  '.(0), ..(0) etc cause dir to read it as cwd!
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If InStr(path, Chr(0)) > 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

