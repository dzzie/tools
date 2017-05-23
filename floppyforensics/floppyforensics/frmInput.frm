VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1800
      Width           =   7575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Waiting As Boolean

Function GetInput(msg, def) As String
    Text1 = def
    Label1.Caption = Replace(Replace(Replace(msg, "\q", """"), "\n", vbCrLf), "\t", vbTab)
    Waiting = True
    Me.Show
    
    While Waiting
        DoEvents: Sleep 100
    Wend
    
    GetInput = Text1
    Unload Me
End Function

Private Sub Command1_Click()
    Waiting = False
End Sub

Private Sub Command2_Click()
    Text1 = Empty
    Waiting = False
End Sub

