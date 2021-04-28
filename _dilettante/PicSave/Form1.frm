VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008060&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PicSave.SavePicture Demo"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   6240
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4590
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   4920
      Picture         =   "Form1.frx":44AC
      Top             =   3360
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Menu mnuPicture 
      Caption         =   "Picture"
      Begin VB.Menu mnuPictureDraw 
         Caption         =   "Draw something"
      End
      Begin VB.Menu mnuPictureSave 
         Caption         =   "Save as JPEG"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Randomize
End Sub

Private Sub mnuPictureDraw_Click()
    Dim I As Long
    Dim Factor As Single
    
    AutoRedraw = True
    Cls
    With Image1.Picture
        For I = 1 To Int(Rnd() * 5) + 5
            Line (Rnd() * ScaleWidth, 0)-(Rnd() * ScaleWidth, ScaleHeight), vbRed
            Factor = Rnd() * 0.5 + 0.5
            .Render hDC, _
                    ScaleX(ScaleWidth * Rnd() * 0.8, ScaleMode, vbPixels), _
                    ScaleY(ScaleHeight * Rnd() * 0.8, ScaleMode, vbPixels), _
                    ScaleX(.Width * Factor, vbHimetric, vbPixels), _
                    ScaleY(.Height * Factor, vbHimetric, vbPixels), _
                    0, _
                    .Height, _
                    .Width, _
                    -.Height, _
                    ByVal 0&
        Next
    End With
    AutoRedraw = False
    mnuPictureSave.Enabled = True
    Caption = "Zombie Laser Attack"
End Sub

Private Sub mnuPictureSave_Click()
    PicSave.SavePicture Me.Image, "test.jpg", fmtJPEG, 70
End Sub
