VERSION 5.00
Begin VB.Form dmSplash 
   BorderStyle     =   0  'None
   Caption         =   "MathImagics"
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image dmImage 
      Height          =   8610
      Left            =   0
      Picture         =   "dmSplash.frx":0000
      Top             =   0
      Width           =   12540
   End
End
Attribute VB_Name = "dmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Me.Height = dmImage.Height
   Me.Width = dmImage.Width
   End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Hide
   End Sub

Private Sub dmImage_Click()
   Hide
   End Sub
