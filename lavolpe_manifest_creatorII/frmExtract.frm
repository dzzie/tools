VERSION 5.00
Begin VB.Form frmExtract 
   Caption         =   "Manifest Element"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtExtract 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmExtract.frx":0000
      Top             =   945
      Width           =   3855
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ever see a functional form with such little code?

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        txtExtract.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End If
End Sub

