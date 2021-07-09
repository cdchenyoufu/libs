VERSION 5.00
Begin VB.Form FlashBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   1815
   ClientLeft      =   2340
   ClientTop       =   2805
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1815
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4200
      Top             =   0
   End
   Begin VB.PictureBox Pic_ApplicationIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "Flashbox.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Lbl_CopyRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Copyright 2017"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Line lin_HorizontalLine1 
      BorderWidth     =   2
      X1              =   375
      X2              =   4410
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Label Lbl_TravPro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "IP WhoIs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   840
      TabIndex        =   1
      Top             =   270
      Width           =   3855
   End
   Begin VB.Label Lbl_Version 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Version3.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1470
   End
   Begin VB.Label Lbl_ComPany 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "JAC Computing, Vernon, BC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   390
      TabIndex        =   3
      Top             =   1320
      Width           =   4365
   End
End
Attribute VB_Name = "FlashBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
    Unload Me
End Sub


