VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Small Server                                                         Updated 10/00 -Q"
   ClientHeight    =   5340
   ClientLeft      =   165
   ClientTop       =   420
   ClientWidth     =   7020
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRaw 
      Height          =   1365
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frmMain.frx":000C
      Top             =   3915
      Width           =   6825
   End
   Begin VB.Frame Frame 
      Caption         =   "Config"
      Height          =   1575
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   2265
      Width           =   6915
      Begin VB.TextBox txtConfig 
         Height          =   315
         Index           =   2
         Left            =   1020
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         Top             =   1020
         Width           =   4875
      End
      Begin VB.TextBox txtConfig 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Text            =   "http://www.geocities.com/"
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtConfig 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   6
         Top             =   660
         Width           =   4515
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   1
         Left            =   6000
         TabIndex        =   3
         Top             =   180
         Width           =   855
         Begin VB.CheckBox chkRaw 
            Alignment       =   1  'Right Justify
            Caption         =   "Raw"
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   570
            Width           =   735
         End
         Begin VB.CommandButton cmdLog 
            Caption         =   "ViewLog"
            Height          =   285
            Left            =   30
            TabIndex        =   12
            Top             =   855
            Width           =   840
         End
         Begin VB.CheckBox chkLog 
            Alignment       =   1  'Right Justify
            Caption         =   "Log"
            Height          =   195
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   735
         End
         Begin VB.CheckBox chkAuth 
            Alignment       =   1  'Right Justify
            Caption         =   "Auth"
            Height          =   195
            Left            =   60
            TabIndex        =   4
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Serve File :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Redirect Probes to :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Authorized URL :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "  Connections  "
      Height          =   2175
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   6915
      Begin ComctlLib.ListView lvIPs 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   3201
         View            =   3
         Arrange         =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "IP"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Time"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Req's"
            Object.Width           =   661
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "User Agent"
            Object.Width           =   8072
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Authcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "PROBED"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   7080
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   7080
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuCopyMyIP 
         Caption         =   "Copy Authorized URL"
      End
      Begin VB.Menu mnuImgAuth 
         Caption         =   "IMG SRC=""AuthURL"""
      End
      Begin VB.Menu mnuDyndns 
         Caption         =   "Set DynDns Name"
      End
      Begin VB.Menu mnuSetLogFile 
         Caption         =   "Set Log File Path"
      End
      Begin VB.Menu mnuspaca 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHowTo 
         Caption         =   "HowTo..."
      End
      Begin VB.Menu mnuReadMe 
         Caption         =   "ReadMe..."
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopyIP 
         Caption         =   "Copy IP"
      End
      Begin VB.Menu mnuCopyInfo 
         Caption         =   "Copy Info"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy ALL"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear ALL"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I just did a major recode on this because there were a couple of weird
'bugs with teh auth url bit and things felt a bit sloppy ( i coded the
'original over a year ago so!) ...anyway this UI is more streamlined
' as soon asyou drop a file in the serve file textbox the server comes online
'if your port 80 is already running somthing it will not liek it :P
'i am still recoding this is an in between edition so come back in a week
'and everythign should be ironed otu again : )  -Sept 20 2001

Private Sub chkRaw_Click()
  Form_Resize
  If chkRaw.value = 1 Then txtRaw = Script.httpHeader & Script.data
End Sub

Private Sub cmdLog_Click()
  If FileExists(cfg.LogFile) Then Shell "notepad " & cfg.LogFile, vbNormalFocus _
  Else: MsgBox "Logfile: " & cfg.LogFile & vbCrLf & " Could not be found"
End Sub

Private Sub mnuDyndns_Click()
    cfg.DynDns = InputBox("If you are on dsl or use a dynamic dns entry set it here. If this value is blank then it will default to your current ip. Note do NOT include the http://", , DynDns)
    cfg.DynDns = Trim(cfg.DynDns)
End Sub

Private Sub mnuHowTo_Click()
    Dim f As String
    f = App.path & "\Smallserver.txt"
    If fso.FileExists(f) Then
        Shell "notepad """ & f & """", vbNormalFocus
    Else
        MsgBox "Howto File not found - " & vbCrLf & vbCrLf & f, vbExclamation
    End If
End Sub

Private Sub mnuReadMe_Click()
Dim f As String
    f = App.path & "\README.txt"
    If fso.FileExists(f) Then
        Shell "notepad """ & f & """", vbNormalFocus
    Else
        MsgBox "Readme File not found - " & vbCrLf & vbCrLf & f, vbExclamation
    End If
End Sub

Private Sub mnuSetLogFile_Click()
  cfg.LogFile = InputBox("Fill in the path where you would like the Log file saved", , LogFile)
  If Not FolderExists(fso.GetParentFolder(cfg.LogFile)) Then
    MsgBox "Parent Folder does not exist please choose valid path"
    mnuSetLogFile_Click
  End If
End Sub

Private Sub Form_Load()
  Call loadConfig
  If cfg.DynDns = Empty Then cfg.DynDns = sckServer(0).LocalIP
  txtConfig(1) = "http://" & cfg.DynDns & "/" & cfg.LastPath
  Me.Caption = "Small Server --Offline"
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call saveConfig
End Sub

Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
    X = -1
    For i = 1 To sckServer.UBound
        If sckServer(i).State <> sckConnected And _
           sckServer(i).State <> sckConnecting And _
           sckServer(i).State <> sckConnectionPending Then
           ' use first available socket...unload all extras
           If X = -1 Then X = i Else Unload sckServer(i)
        End If
    Next

    If X < 1 Then X = sckServer.UBound + 1: Load sckServer(X)
    
    sckServer(X).Close
    sckServer(X).Accept requestID
End Sub


Private Sub sckServer_DataArrival(index As Integer, ByVal bytesTotal As Long)
With sckServer(index)
  Dim h As HTTPRequest

  .GetData strdata, vbString
  h = Globals.ParseRequest(strdata, .RemoteHostIP)
    
  ip = .RemoteHostIP
  AddNewIP .RemoteHostIP, h.uAgent
   
  If chkAuth.value = 1 Then
        If InStr(1, Script.AuthURL, h.page) < 1 Then
          rejectHeader = buildHeader(301, "Close", "text/html", , Script.probeServer & h.page)
          .SendData rejectHeader
          AddNewIP .RemoteHostIP, "PROBING-" & UsrAgent, , False, h.page
          logit .RemoteHostIP, "PROBING FROM " & .RemoteHost & "  PORT=" & .RemotePort, strdata
          WaitForSentAndClosed
          Exit Sub
        End If
  End If
        
  If Len(h.BasicAuth) > 0 Then
        AddNewIP .RemoteHostIP, h.uAgent, Base64Decode(h.BasicAuth), False
        .SendData Script.AuthedHeader
  Else
        If chkRaw.value = 1 Then
           .SendData txtRaw
           WaitForSentAndClosed
        ElseIf Script.RespCode = 401 Then
          .SendData Script.AuthedHeader
          WaitForSentAndClosed
        Else
          .SendData Script.httpHeader & Script.data
          WaitForSentAndClosed
        End If
  End If
    
  'On Error Resume Next
  If chkLog.value = 1 Then Call logit(ip, , strdata)
    
End With
End Sub

Private Sub sckServer_SendComplete(index As Integer)
  sckServer(index).Close
  If index > 0 Then Unload sckServer(index)
  ReadyToClose = True
End Sub

Private Sub AddNewIP(strIP As String, _
                     UsrAgent As String, _
                     Optional AUTH = False, _
                     Optional INC = True, _
                     Optional PROBED = False _
                    )
 'yes yes swiss army functions are bad I know but
 'how the macgyver in me loves them so!

With lvIPs.ListItems
  If .Count = 0 Then GoTo First

  For i = 1 To .Count
    If strIP = .Item(i) Then
      .Item(.Count).SubItems(1) = time:                             IPs(.Count).time = time
      If INC Then .Item(i).SubItems(2) = .Item(i).SubItems(2) + 1:  IPs(.Count).reqs = IPs(.Count).reqs + 1
      .Item(.Count).SubItems(3) = UsrAgent:                         IPs(.Count).userAgent = UsrAgent
      If AUTH <> False Then .Item(.Count).SubItems(4) = AUTH:       IPs(.Count).authed = AUTH
      If PROBED <> False Then .Item(.Count).SubItems(5) = PROBED:   IPs(.Count).PROBED = PROBED
      Exit Sub
    ElseIf i = .Count Then GoTo First
    End If
  Next

Exit Sub
First:
      .Add = strIP:                                                 ReDim Preserve IPs(.Count)
      .Item(.Count).SubItems(1) = time:                             IPs(.Count).time = time
      .Item(.Count).SubItems(2) = 1:                                IPs(.Count).reqs = 1
      .Item(.Count).SubItems(3) = UsrAgent:                         IPs(.Count).userAgent = UsrAgent
      .Item(.Count).SubItems(4) = AUTH:                             IPs(.Count).authed = AUTH
      .Item(.Count).SubItems(5) = PROBED:                           IPs(.Count).PROBED = PROBED
End With
End Sub

Private Sub mnuImgAuth_Click()
   Clipboard.Clear
   Clipboard.SetText "<img src=""" & txtConfig(1) & """>"
End Sub

Private Sub mnuCopyMyIP_Click()
  Clipboard.Clear
  Clipboard.SetText txtConfig(1)
  lvIPs.SetFocus
End Sub

Private Sub mnuCopyInfo_Click()
On Error Resume Next
 With lvIPs.ListItems.Item(lvIPs.SelectedItem.index)
  Clipboard.Clear
  strdata = lvIPs.ListItems.Item(lvIPs.SelectedItem.index) & "  " & _
            .SubItems(1) & "  " & .SubItems(2) & "  " & .SubItems(3) & "  " & _
            "AuthCode [" & .SubItems(4) & "]  Probing [" & .SubItems(5) & "]"
  Clipboard.SetText strdata
 End With
End Sub

Private Sub mnuCopyAll_Click()
On Error Resume Next
sp = "     "
Dim strdata As String
For i = 0 To UBound(IPs)
 With IPs(i)
  strdata = strdata & .ip & sp & .time & sp & .reqs & sp & .userAgent & sp & .authed & sp & .PROBED & vbCrLf
 End With
Next
Clipboard.Clear
Clipboard.SetText strdata
End Sub

Private Sub mnuClear_Click()
    lvIPs.ListItems.Clear
  ReDim IPs(0)
End Sub

Private Sub Form_Resize()
On Error Resume Next
  stat = Me.Width
  If chkRaw.value = 1 Then Me.Height = 6030 Else Me.Height = 4560
  For i = 0 To Frame.UBound: Frame(i).Width = stat - 300: Next
  lvIPs.Width = Frame(0).Width - 235
  Frame(1).Left = stat - (535 + 795)
  For i = 0 To txtConfig.UBound: txtConfig(i).Width = stat - (txtConfig(i).Left + 1500): Next
End Sub
Private Sub mnuCopyIP_Click()
  Clipboard.Clear
  Clipboard.SetText lvIPs.ListItems.Item(lvIPs.SelectedItem.index)
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then frmMain.PopupMenu mnuFile
End Sub

Private Sub lvIPs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then frmMain.PopupMenu mnuEdit
End Sub

Private Sub logit(who, Optional it = "", Optional dat = "")
      If Not fso.FileExists(cfg.LogFile) Then fso.CreateFile (cfg.LogFile)
      nfo = "[ " & who & " ]   " & it & vbCrLf & dat & vbCrLf & vbCrLf
      fso.AppendFile cfg.LogFile, nfo
End Sub

Private Sub txtConfig_OLEDragDrop(index As Integer, data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If index = 2 Then
    txtConfig(2).Text = data.Files(1)
    Call startup
  End If
End Sub

Function WaitForSentAndClosed()
    ReadyToClose = False
    While Not ReadyToClose
        DoEvents: DoEvents: DoEvents: DoEvents
    Wend
End Function

Sub startup()
  Call resetScript
  
  sckListen.Close
  sckListen.LocalPort = 80
  sckListen.Listen
  sckServer(0).Close
  Me.Caption = "Small Server -- Online"
  
  Script.path = txtConfig(2)
  Script.data = ReadFile(Script.path)
  Script.AuthURL = txtConfig(1)
  Script.probeServer = txtConfig(0)
  Script.extension = fso.GetExtension(Script.path)
  
  Select Case Script.extension
    Case ".jpg", ".jpeg"
      Script.httpHeader = buildHeader(200, "Keep-Alive", "image/jpeg")
    Case ".gif"
      Script.httpHeader = buildHeader(200, "Keep-Alive", "image/gif")
    Case ".exe", ".zip"
        Script.httpHeader = buildHeader(200, "Close", "application/x-compress")
    Case ".loc"
        Script.httpHeader = buildHeader(301, "Close", "text/html", , Script.data)
    Case ".raw"
        Script.httpHeader = Script.data
        Script.data = Empty
    Case ".auth"
       authinfo = Split(Script.data, vbCrLf)
       If UBound(authinfo) <> 1 Then
          MsgBox "Format of .auth files is" & vbCrLf & vbCrLf & "Line1 : Authorization Dialogue" & vbCrLf & "Line2 : Url to relay them to on authorization", vbCritical, "USER ERROR !"
          txtConfig(2) = ""
          Exit Sub
       End If
       Script.httpHeader = buildHeader(401, "Keep-Alive", "text/html", authinfo(0))
       Script.AuthedHeader = buildHeader(301, "Close", "text/html", , authinfo(1))
    Case Else
      Script.httpHeader = buildHeader(200, "Close", "text/html")
  End Select
  
  If chkRaw.value = 1 Then txtRaw = Script.httpHeader
End Sub


