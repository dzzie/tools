VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCheckMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Mail"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   630
      Top             =   4620
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   315
      Left            =   1980
      TabIndex        =   6
      Top             =   4680
      Width           =   1590
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   210
      TabIndex        =   5
      Top             =   735
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox Text5 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1170
      Width           =   5295
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4530
      TabIndex        =   1
      Top             =   210
      Width           =   900
   End
   Begin VB.TextBox Text3 
      Height          =   2040
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2550
      Width           =   5295
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   135
      Top             =   4635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblMessage 
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
      Height          =   300
      Left            =   1140
      TabIndex        =   4
      Top             =   315
      Width           =   3270
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
      Left            =   225
      TabIndex        =   3
      Top             =   315
      Width           =   825
   End
End
Attribute VB_Name = "frmCheckMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gm(8) As String
Dim datafile As String
Dim step As Integer
Dim db As Boolean

Private Type Message
  total As Integer
  retr As Integer
  bytes As Long
  dele As Boolean
  receiving As Boolean
  fhand As Integer
  'fname As String
End Type

Dim m As Message     'holds stats of currently downloading msg
Dim msgs() As String 'array hold msg filenames
Dim Cancel As Boolean
Dim readyToReturn As Boolean


Private Sub cmdStop_Click()
  Cancel = True
End Sub

Private Sub Command3_Click()
Unload Me
readyToReturn = True
End Sub

Private Sub Form_Load()
    Me.Height = 1440
    m.receiving = False
    pb.Value = 0
    db = False
    readyToReturn = False
    Cancel = False
    
    n = "frmCheckMail"
    Me.Left = GetSetting("qmail", n, "MainLeft", 1000)
    Me.Top = GetSetting("qmail", n, "MainTop", 1000)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    n = "frmCheckMail"
    If Me.WindowState <> vbMinimized Then
        SaveSetting "qmail", n, "MainLeft", Me.Left
        SaveSetting "qmail", n, "MainTop", Me.Top
    End If
End Sub

Public Function getMail(Server As String, user As String, pass As String, Optional port = 110, Optional debugWindow = False, Optional hidden = False) As String()
  On Error GoTo warn
  
  If Not hidden Or debugWindow Then Me.Visible = True
  If debugWindow Then Me.Height = 5475: db = True
  tmrTimeout.Enabled = True
  
  gm(1) = "USER " & user & vbCrLf
  gm(2) = "PASS " & pass & vbCrLf
  gm(3) = "STAT" & vbCrLf
    
  Winsock1.Close
  Winsock1.Connect Server, port
  
  'from here server initiates process with inital data arrival event
  'so for now we have to just wait until mail is downloaded and global
  'array msgs() is full of messages before we can return
  
  While Not readyToReturn
     DoEvents
     DoEvents
     DoEvents
  Wend
  
  getMail = msgs()
  If Not db Then Unload Me
  
  Exit Function
warn:
  MsgBox Err.Description
  
End Function

Private Sub tmrTimeout_Timer()
    If Me.Visible = True Then
        If MsgBox("I am tired of waiting for the server to respond. Do you want to time out?", vbYesNo) = vbYes Then
            readyToReturn = True
        End If
    Else
        readyToReturn = True
    End If
End Sub

Private Sub Winsock1_Connect()
    step = 0
    ReDim msgs(0)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo warn
Dim RESP As String

Winsock1.GetData RESP, vbString
OK = InStr(1, Mid(RESP, 1, 4), "OK ")
If InStr(1, Mid(RESP, 1, 5), "ERR ") > 0 Then GoTo warn

tmrTimeout.Enabled = False
tmrTimeout.Enabled = True

If Cancel Then
    Winsock1.SendData "QUIT" & vbCrLf
    step = 5
    Cancel = False
    Exit Sub
End If

l (RESP)
  If OK And Not m.receiving Then 'acepted command response
    Select Case step
       Case 0, 1, 2 'login and see if there is mail
           step = step + 1
           l (gm(step))
           Winsock1.SendData gm(step)
       Case 3  'do we send UIDL or QUIT ?
            fspc = InStr(1, RESP, " ")
            lspc = InStr(fspc + 1, RESP, " ")
            Msg = Mid(RESP, fspc, (lspc - fspc))
            If Msg > 0 Then
              ReDim msgs(Msg) 'set array size once cause we know how many msgs
              m.fhand = FreeFile
              msgs(Msg) = fso.CreateTempFile(uc.folders.inbox, ".txt")
              Open msgs(Msg) For Binary As m.fhand
              m.total = Msg
              m.retr = Msg
              m.dele = False
              step = 4
              l ("RETR " & m.retr)
              Winsock1.SendData "RETR " & m.retr & vbCrLf
              m.receiving = True
            Else
              step = 5
              l ("QUIT")
              Winsock1.SendData "QUIT" & vbCrLf
            End If
       Case 4
           If m.retr > 0 Then
            l ("RETR " & m.retr)
              m.fhand = FreeFile
              msgs(m.retr) = fso.CreateTempFile(uc.folders.inbox, ".txt")
              Open msgs(m.retr) For Binary As m.fhand
              m.receiving = True
              Winsock1.SendData "RETR " & m.retr & vbCrLf
           Else
              step = 5
              l ("QUIT")
              Winsock1.SendData "QUIT" & vbCrLf
           End If
       Case 5
           Winsock1.Close
           readyToReturn = True 'now checkmail can return and unload form
       End Select
  
  ElseIf m.receiving Then
       If OK Then Call setbyte(RESP)
       If db Then
          Text3 = Text3 & RESP
          If Len(Text3) > 5000 Then Text3 = ""
       End If
       Put m.fhand, , RESP
       m.bytes = m.bytes - bytesTotal
       Call incProgbar(bytesTotal)
       If m.bytes < 0 Or (InStr(1, RESP, vbCrLf & "." & vbCrLf) And m.bytes < 10) Then
            l ("Bytes Left : " & m.bytes)
            l ("DELE " & m.retr)
            pb.Value = 0
            Close m.fhand
            Winsock1.SendData "DELE " & m.retr & vbCrLf
            If m.retr > 0 Then m.retr = m.retr - 1 Else step = 5
            m.receiving = False
       End If
  End If

Exit Sub
warn:
    l (RESP & vbCrLf)
    l ("!!ERROR-" & m.receiving & " " & m.retr & " " & OK)
    MsgBox "Communication Error,Exiting Program  " & vbCrLf & vbCrLf & "Err Desc= " & RESP, , Err.Description
    'Winsock1.SendData "QUIT" & vbCrLf dont want to quit or else it will delete all marked messages !
    Winsock1.Close
    readyToReturn = True 'triggers return val in checkmail function
    step = 0
End Sub

Private Sub setbyte(inp As String) 'teh ok after the retr has the message
       tmp = Split(inp, vbCrLf)    'size in it, after the stat is bytes of total dl
       fspc = InStr(1, tmp(0), " ")
       lspc = InStr(fspc + 1, tmp(0), " ")
       m.bytes = Int(Mid(tmp(0), fspc, lspc - fspc))
       pb.Max = m.bytes + 60
End Sub

Private Sub incProgbar(X)
     On Error Resume Next
     pb.Value = pb.Value + X
End Sub

Private Sub l(it)
 If db Then
    If Len(Text5) > 5000 Then Text5 = ""
    Text5 = Text5 & it
 End If

 lblMessage = it

End Sub


'example convro with server
'   +OK hello from popgate
'   user dzzie
'   +OK password required
'   pass xxxxxx
'   +OK maildrop ready, 1 message (719 octlets)(no[7] no[7])
'   STAT
'   +OK 1 719 [its saying msg 1 is 719 bytes]
'   retr 1
'   [message sent]
'   dele 1
'   +OK message 1 marked as deleted
'   QUIT
'   +OK server signing off
'
