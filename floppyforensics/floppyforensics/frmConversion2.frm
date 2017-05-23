VERSION 5.00
Begin VB.Form frmConv 
   Caption         =   "Character Conversion"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7275
   LinkTopic       =   "Form2"
   ScaleHeight     =   1245
   ScaleWidth      =   7275
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   2
      Left            =   2520
      TabIndex        =   15
      Top             =   525
      Width           =   645
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   1
      Left            =   1785
      TabIndex        =   14
      Top             =   525
      Width           =   645
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   525
      Width           =   645
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   195
      Left            =   1845
      TabIndex        =   11
      Top             =   1000
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   270
      Left            =   4875
      TabIndex        =   10
      Top             =   885
      Width           =   1005
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   165
      TabIndex        =   8
      Top             =   945
      Width           =   255
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   6150
      TabIndex        =   6
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox txtHex 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4080
      TabIndex        =   5
      Top             =   480
      Width           =   3030
   End
   Begin VB.TextBox txtAsc 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   60
      Width           =   3030
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Top             =   105
      Width           =   2670
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "KeyCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   105
      TabIndex        =   13
      Top             =   525
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Space Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2145
      TabIndex        =   12
      Top             =   945
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CGI Encode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   945
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Str"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   165
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hex :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3390
      TabIndex        =   3
      Top             =   525
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ascii :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3390
      TabIndex        =   2
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "frmConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal _
    hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx _
    As Long, ByVal cy As Long, ByVal wFlags As Long)
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
Dim inString As Boolean
Dim inHex As Boolean


Private Sub cmdClear_Click()
 Text1 = "": txtAsc = "": txtHex = ""
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub


Private Sub Form_Load()
inString = True
inHex = False
SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, _
        Me.Top / 15, Me.Width / 15, _
        Me.Height / 15, SWP_SHOWWINDOW
End Sub

Private Sub Text1_GotFocus()
inString = True: inHex = False
End Sub





Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        Text2(1) = KeyCode
        Text2(2) = Hex(KeyCode)
        Text2(0) = Empty
    End If
End Sub

Private Sub txtasc_GotFocus()
inString = False: inHex = False
End Sub

Private Sub txthex_GotFocus()
inString = False: inHex = True
End Sub

Private Sub Text1_Change()
If Not inString Then Exit Sub
If Text1 = Empty Then txtHex = Empty: txtAsc = Empty
On Error Resume Next

Dim letter As String
p = " "

If Check1.value = 1 Then p = "%"
If Check2.value = 0 Then p = "": Check1.value = 0

If Len(Text1) > 1 Then letter = Mid(Text1, Len(Text1), 1) _
Else: letter = Text1

txtAsc = txtAsc & " " & Asc(letter)
txtHex = txtHex & p & Hex(Asc(letter))
 
End Sub

Private Sub txthex_change()
If Not inHex Then Exit Sub
If txtHex = Empty Then Text1 = Empty: txtAsc = Empty

On Error Resume Next
If InStr(txtHex, " ") > 0 Then
    ary = Split(txtHex, " ")
    For i = 0 To UBound(ary)
        t1 = t1 & Chr(Int("&h" & ary(i))) & " "
        ta = ta & Int("&h" & ary(i)) & " "
    Next
    Text1 = t1: txtAsc = ta
Else
    Text1 = Chr(Int("&h" & txtHex))
    txtAsc = Int("&h" & txtHex)
End If
    
    
End Sub

Private Sub txtasc_change()
If inString Or inHex Then Exit Sub
If txtAsc = Empty Then txtHex = Empty: Text1 = Empty

On Error Resume Next
Text1 = Chr(txtAsc)
txtHex = Hex(txtAsc)

End Sub
