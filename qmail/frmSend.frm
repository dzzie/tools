VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sending Mail"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   105
      TabIndex        =   5
      Top             =   630
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   4200
      Top             =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3975
      TabIndex        =   4
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close Window"
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   3270
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   2040
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1125
      Width           =   4830
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4620
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   3
      Top             =   210
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Connecting to Server..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   975
      TabIndex        =   0
      Top             =   225
      Width           =   2910
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datafile As String
Dim convro() As String
Dim step As Integer
Dim db As Boolean
Dim msgStep As Integer

Private Sub Command2_Click()
  Winsock1.Close
  Unload Me
End Sub

Private Sub Form_Load()
    db = False
    step = 0
    Me.Height = 1410
    
    n = "frmSend"
    Me.Left = GetSetting("qmail", n, "MainLeft", 1000)
    Me.Top = GetSetting("qmail", n, "MainTop", 1000)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    n = "frmSend"
    If Me.WindowState <> vbMinimized Then
        SaveSetting "qmail", n, "MainLeft", Me.Left
        SaveSetting "qmail", n, "MainTop", Me.Top
    End If
End Sub
Private Sub Command1_Click()
 Unload Me
End Sub

'mabey make messagepath and read incrementally and have progressbar
'for large attachments

Public Sub sendMail(SendTo As String, SentFrom As String, Message As String, Server As String, Optional port = 25, Optional debugWindow = False)
On Error GoTo warn
    Me.Visible = True
    Label1.Caption = "Connecting to server"
    ProgressBar1.Value = 0
    If debugWindow Then Me.Height = 4080: db = True
    If Not db Then tmrTimeout.Enabled = True
    
    ReDim convro(0) '<-forgot this and caused horrible bug!
    
    push convro, "HELO " & Winsock1.LocalHostName & vbCrLf
    push convro, "MAIL FROM: " & SentFrom & vbCrLf
    
    If InStr(SendTo, ",") > 0 Then
        tmp = Split(SendTo, ",")
        For i = 0 To UBound(tmp)
            If tmp(i) <> Empty Then
               push convro, "RCPT TO: <" & tmp(i) & ">" & vbCrLf
            End If
        Next
    Else
      push convro, "RCPT TO: <" & SendTo & ">" & vbCrLf
    End If
    
    push convro, "DATA" & vbCrLf
    push convro, Message & vbCrLf & "." & vbCrLf
    msgStep = UBound(convro)
    ProgressBar1.Max = Len(Message)
    push convro, "QUIT " & vbCrLf & vbCrLf
    
    Winsock1.Close
    Winsock1.Connect Server, port
    
Exit Sub
warn: MsgBox "Error in mail function " & Err.Description
End Sub


Private Sub Label1_Click()
  MsgBox Label1.Caption
End Sub

Private Sub tmrTimeout_Timer()
    If step <> msgStep Then
        If MsgBox("We are tired of waiting for mail server to respond would you like to let this send attempt to time out?", vbYesNo) = vbYes Then
            MsgBox "Message not sent", vbExclamation
            Unload Me
        End If
    End If
End Sub

Private Sub Winsock1_Connect()
  Call mailer
End Sub

Private Sub mailer()
   Winsock1.SendData convro(step)
   Label1.Caption = convro(step)
   tmrTimeout.Enabled = False
   If step <> msgStep And Not db Then tmrTimeout.Enabled = True
   If db Then Text1 = Text1 & convro(step) & vbCrLf
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData datafile
    reply = Mid(datafile, 1, 3)
    If db And Len(Text1) < 2000 Then Text1 = Text1 & datafile & vbCrLf
    Label1.Caption = datafile
      If reply = 250 Or reply = 354 Then  'acepted command response
            step = step + 1
            Call mailer
      ElseIf reply = 220 Then  'welcome response
            step = step + 1
            Call mailer
      ElseIf reply = 221 Then 'server closing
            Winsock1.Close
            step = 0
            If Not db Then Unload Me
      Else                     'error (50x) response
            MsgBox Err.Description & vbCrLf & vbCrLf & datafile
            Winsock1.Close
            step = 0
      End If
End Sub


Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    If step = msgStep Then
       nval = ProgressBar1.Value + bytesSent
       ProgressBar1.Value = IIf(nval <= ProgressBar1.Max, nval, ProgressBar1.Max)
    End If
End Sub
