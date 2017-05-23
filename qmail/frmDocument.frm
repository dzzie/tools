VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCompose 
   Caption         =   "frmCompose"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7530
   Begin VB.TextBox txtProps 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   885
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   1245
      Width           =   6390
   End
   Begin VB.TextBox txtProps 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   870
      TabIndex        =   6
      Top             =   990
      Width           =   6405
   End
   Begin VB.TextBox txtProps 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   870
      TabIndex        =   5
      Top             =   735
      Width           =   6420
   End
   Begin VB.TextBox txtProps 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   900
      TabIndex        =   4
      Top             =   465
      Width           =   6405
   End
   Begin VB.TextBox txtProps 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   900
      TabIndex        =   3
      Top             =   195
      Width           =   6405
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1275
      Left            =   180
      TabIndex        =   1
      Top             =   210
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   2249
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "mine"
         Object.Width           =   706
      EndProperty
   End
   Begin VB.TextBox txtMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1620
      Width           =   7410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1515
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   7395
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRctp 
         Caption         =   "Insert Rcpt"
         Begin VB.Menu mnuEmail 
            Caption         =   "[ Empty ]"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuAdvanced 
      Caption         =   "mnuAdvanced"
      Visible         =   0   'False
      Begin VB.Menu mnuSend 
         Caption         =   "Send "
      End
      Begin VB.Menu mnuSaveNoSend 
         Caption         =   "Save (NoSend)"
      End
      Begin VB.Menu mnuSpellChk 
         Caption         =   "Spell Check"
      End
      Begin VB.Menu spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHtmlMsg 
         Caption         =   "Type: text/html"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug Mode"
      End
   End
End
Attribute VB_Name = "frmCompose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private index As Integer, Qued As Boolean, fpath As String, loadsize As Long

Function buildStdHeaders() As String
    Dim eMail As String
    
    For i = 0 To txtProps.Count - 2
      header = ListView1.ListItems(i + 1).key
      If i < 3 And txtProps(i) = "" Then
           MsgBox header & " field is blank. Please Fill in."
           Exit Function
      Else
           If txtProps(i) <> "" Then eMail = eMail & header & txtProps(i) & vbCrLf
      End If
    Next
        
    If Not AryIsEmpty(uc.xHeaders) Then
      For i = 1 To UBound(uc.xHeaders)
        eMail = uc.xHeaders(i) & vbCrLf & eMail
      Next
    End If
    
    buildStdHeaders = "X-Mailer: Qmail v.8" & vbCrLf & eMail
End Function

Private Sub SendMessage()
    Dim eMail As String
    
    eMail = buildStdHeaders()
    If eMail = Empty Then Exit Sub
    
    If txtProps(4) <> Empty Then
        SetupAttachment eMail 'modifies parent object
    Else
       eMail = eMail & "MIME-Version: 1.0" & vbCrLf
       eMail = eMail & "Content-Type: " & getType() & vbCrLf
       eMail = eMail & vbCrLf & txtMsg
       If fso.FileExists(uc.folders.sigFile) Then
            eMail = eMail & vbCrLf & vbCrLf & fso.readFile(uc.folders.sigFile)
       End If
    End If
    
    SendTo = txtProps(1) & IIf(txtProps(3) = Empty, Empty, "," & txtProps(3))
    Call frmSend.sendMail(CStr(SendTo), txtProps(0), eMail, uc.Send.Server, , mnuDebug.Checked)
    
    If Qued Then
        frmMessages.lv(2).ListItems(index).Text = "S"
        fso.writeFile fpath, eMail
    Else
        If uc.Prefs.saveSent Then Call Library.saveSent(eMail)
    End If
    
    uc.SentMessages = uc.SentMessages + 1
    Unload Me
End Sub

Private Sub mnuDebug_Click()
  If mnuDebug.Checked = False Then mnuDebug.Checked = True _
  Else mnuDebug.Checked = False
End Sub

Private Sub mnuEmail_Click(index As Integer)
  On Error GoTo other
  txtProps(1).SetFocus 'thorws error if already has focus
  txtProps(3) = txtProps(3) & IIf(Trim(txtProps(1)) = Empty, Empty, ",") & mnuEmail(index).Caption
  Exit Sub
other:
  txtProps(1) = txtProps(1) & IIf(Trim(txtProps(1)) = Empty, Empty, ",") & mnuEmail(index).Caption
End Sub



Private Sub mnuHtmlMsg_Click()
    With mnuHtmlMsg
        If .Checked Then .Checked = False Else .Checked = True
    End With
End Sub

Private Sub mnuSaveNoSend_Click()
   If Trim(txtMsg) = Empty Then MsgBox "Yeah right save what :P": Exit Sub
   eMail = buildStdHeaders & vbCrLf & vbCrLf & txtMsg
   
   If Not Qued Then Call Library.saveSent(CStr(eMail), True) _
   Else fso.writeFile fpath, eMail
   
   Unload Me
End Sub

Private Sub mnuSend_Click()
  SendMessage
End Sub

Private Sub mnuSpellChk_Click()
    txtMsg = SpellCheck(txtMsg.Text)
End Sub

Private Sub txtMsg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then ShowRtClkMenu Me, txtMsg, mnuAdvanced
End Sub

Private Sub txtProps_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If index = 1 Or index = 3 Then _
     If Button = 2 Then ShowRtClkMenu Me, txtProps(index), mnuPopup
End Sub

Private Sub txtProps_OLEDragDrop(index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If index = 4 Then txtProps(4) = Data.Files(1)
End Sub

Private Sub Txtmsg_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
        Case 1: txtMsg.SelStart = 0: txtMsg.SelLength = Len(txtMsg) 'ctrl a -> select all
        Case 2:  MsgBox "Selected string Length: " & Len(txtMsg.SelText)  'ctrl b -> add letters
        Case 12: txtMsg.SelText = LCase(txtMsg.SelText) 'ctrl l -> lcase selection
        Case 21: txtMsg.SelText = UCase(txtMsg.SelText) 'ctrl u -> ucase selection
        'Case Else: MsgBox KeyAscii
   End Select
End Sub


Public Sub loadSentMail(path As String, q As Boolean, i As Integer)
    Dim p As parsed
    If FileLen(path) > 58500 Then txt = ReadFirst(path, 20000) _
    Else txt = readFile(path)
    p = Library.parseMail(txt)
    txtProps(0) = p.from
    txtProps(1) = p.to
    txtProps(2) = p.subj
    txtMsg = Mid(p.body, InStr(p.body, vbCrLf & vbCrLf) + 6, Len(p.body))
    fpath = path
    Qued = q
    index = i
    loadsize = Len(txtMsg)
End Sub

Private Sub SetupAttachment(header)

   If Not FileExists(txtProps(4)) Then
      MsgBox "Couldnt find Attachment!", vbCritical
      Exit Sub
   ElseIf FileLen(txtProps(4)) > 1196505 Then '1.14Mb max size
      MsgBox br("Cmon now...that is to big to email, protocall abuser!\n\n *crosses arms and shakes head..nugghh ughh i aint gunna do it ! :P"), vbInformation
      Exit Sub
   End If
    
   boundry = "=====================_" & fso.RandomNum & "==_"
   fname = fso.FileNameFromPath(txtProps(4))
   mType = MimeTypeFromName(fname)
   mimedFile = b64.MimeFileToString(txtProps(4))
   
   Dim h()
    push h, header & "Mime-Version: 1.0"
    push h, "Content-Type: multipart/mixed; boundary=""" & boundry & """"
    push h, Empty
    push h, "--" & boundry
    push h, "Content-Type: text/plain; charset=""us-ascii"""
    push h, Empty
    push h, txtMsg
    push h, "--" & boundry
    push h, "Content-Type: " & mType & "; name=""" & fname & """"
    push h, "Content-Transfer-Encoding: base64"
    push h, "Content-Disposition: attachment; filename=""" & fname & """"
    push h, Empty
    push h, mimedFile
    push h, "--" & boundry
    push h, "Content-Type: text/plain; charset=""us-ascii"""
    push h, Empty
           
    If fso.FileExists(uc.folders.sigFile) Then
            push h, fso.readFile(uc.folders.sigFile)
    End If
            
    push h, vbCrLf & "--" & boundry & "--" & vbCrLf
    
    'modifies parent object !
    header = Join(h, vbCrLf)
End Sub

Private Sub Form_Load()
  On Error GoTo warn
   winApi.SetForegroundWindow Me.hWnd
   winApi.BringWindowToTop Me.hWnd
  
   Form_Resize
   ListView1.ListItems.Add 1, "From: ", "From:"
   ListView1.ListItems.Add 2, "To: ", "   To:"
   ListView1.ListItems.Add 3, "Subject: ", "Subj:"
   ListView1.ListItems.Add 4, "Cc: ", "   Cc:"
   ListView1.ListItems.Add 5, "Attached: ", "Atch:"
   txtProps(0).Text = uc.Send.sender
      
   n = "frmCompose"
   Me.Left = GetSetting(App.Title, n, "MainLeft", 1000)
   Me.Top = GetSetting(App.Title, n, "MainTop", 1000)
   Me.Width = GetSetting(App.Title, n, "MainWidth", 6500)
   Me.Height = GetSetting(App.Title, n, "MainHeight", 6500)
   
   txtMsg.FontSize = uc.fonts.size
   txtMsg.FontName = uc.fonts.face
   txtMsg.ForeColor = CLng(uc.fonts.color)
   txtMsg.backcolor = CLng(uc.fonts.backcolor)
   txtMsg.FontBold = uc.fonts.bold

   If AryIndexExists(uc.recipants, 1) Then
        For i = 1 To UBound(uc.recipants)
              If i > 1 Then Load mnuEmail.Item(i)
              mnuEmail.Item(i).Caption = uc.recipants(i)
              mnuEmail.Item(i).Enabled = True
              mnuEmail.Item(i).Visible = True
        Next
   End If
   
  'mnuDebug.Checked = True
Exit Sub
warn: MsgBox "err in frmcompose form load : " & Err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    n = "frmCompose"
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, n, "MainLeft", Me.Left
        SaveSetting App.Title, n, "MainTop", Me.Top
        SaveSetting App.Title, n, "MainWidth", Me.Width
        SaveSetting App.Title, n, "MainHeight", Me.Height
    End If
    Set e = Nothing
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width > 6000 Then
      Label1.Width = Me.Width - Label1.Left - 200
      For i = 0 To txtProps.Count - 1
        txtProps(i).Width = Me.Width - txtProps(i).Left - 400
      Next
      txtMsg.Width = Me.Width - txtMsg.Left - 200
    End If
    If Me.Height > 6000 Then txtMsg.Height = Me.Height - txtMsg.Top - 400
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
       On Error Resume Next
       
       l = Len(txtMsg)
       If UnloadMode <> 0 Then Exit Sub
       If l < 20 Or l = loadsize Then Exit Sub
       
       'if new msg or contents changed
       If fpath = Empty Then
            msg = "Document is unsaved are you sure you want to exit?"
            If MsgBox(msg, vbExclamation + vbYesNo) = vbNo Then Cancel = -1
       End If
End Sub

Private Function MimeTypeFromName(filename)
    Dim X
    Select Case fso.GetExtension(filename)
        Case ".txt": X = "text/plain"
        Case ".html", ".htm": X = "text/html"
        Case ".zip": X = "application/zip"
        Case Else: X = "application/octet-stream"
    End Select
    MimeTypeFromName = X
End Function

Private Function getType() As String
    If mnuHtmlMsg.Checked Then getType = "text/html" _
    Else getType = "text/plain"
End Function
