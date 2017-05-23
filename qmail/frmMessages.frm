VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMessages 
   Caption         =   "   Loading"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   8655
   Icon            =   "frmMessages.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   8655
   Begin ComctlLib.ListView lv 
      Height          =   1605
      Index           =   4
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   2831
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "R"
         Object.Width           =   35
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "From"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Subject"
         Object.Width           =   14111
      EndProperty
   End
   Begin qmail.longTimer tmrAutoCheck 
      Left            =   8040
      Top             =   165
      _ExtentX        =   794
      _ExtentY        =   741
   End
   Begin ComctlLib.ListView lv 
      Height          =   1845
      Index           =   3
      Left            =   885
      TabIndex        =   4
      Top             =   1395
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   3254
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "R"
         Object.Width           =   35
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "From"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Subject"
         Object.Width           =   14111
      EndProperty
   End
   Begin ComctlLib.ListView lv 
      Height          =   2220
      Index           =   2
      Left            =   450
      TabIndex        =   3
      Top             =   1005
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "S"
         Object.Width           =   41
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "To:"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Subject"
         Object.Width           =   14111
      EndProperty
   End
   Begin ComctlLib.ListView lv 
      Height          =   2565
      Index           =   1
      Left            =   135
      TabIndex        =   0
      Top             =   660
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "R"
         Object.Width           =   406
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "From"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Subject"
         Object.Width           =   14111
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3135
      Left            =   45
      TabIndex        =   2
      Top             =   210
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   5530
      TabWidthStyle   =   2
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Inbox"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Outbox"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Trash"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Archieve"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Q=quit    D=delete    S=new msg   A=Archieve Mail   Z=check mail   R=rebuild toc"
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
      Left            =   45
      TabIndex        =   1
      Top             =   -45
      Width           =   7020
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "txtPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Selected"
      End
      Begin VB.Menu mnuCheckmail 
         Caption         =   "Check Mail"
         Begin VB.Menu mnuChkAccount 
            Caption         =   "[Empty]"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu mnuAutoCheck 
         Caption         =   "Use AutoCheck"
      End
      Begin VB.Menu div5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewMessage 
         Caption         =   "New message"
      End
      Begin VB.Menu mnuNewMsgTo 
         Caption         =   "New message to"
         Begin VB.Menu mnuRecp 
            Caption         =   "[Empty]"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu div3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmptyTrash 
         Caption         =   "Empty Trash"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug Mode"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Options"
      End
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toc As String
Dim myRole As listStyle

Private Sub Lv_KeyPress(index As Integer, KeyAscii As Integer)
  Select Case KeyAscii
    Case 65: If myRole <> saved Then ArchieveMail 'shift a -> archieve
    Case 68: mnuDelete_Click      'shift d -> delete
    Case 81: Unload Me            'shift q -> quit
    Case 13: Lv_dblClick index    'return
    Case 82: RebuildTocFromFiles lv(myRole), myRole    'shift r
    Case 83: mnuNewMessage_Click  'shift s -> new message
    Case 90: Library.checkMail 1, mnuDebug.Checked 'shift z = check default accnt
    'Case Else: MsgBox KeyAscii
  End Select
End Sub

Private Sub mnuChkAccount_Click(index As Integer)
  Call Library.checkMail(index, mnuDebug.Checked)
End Sub

Private Sub mnuDebug_Click()
  If mnuDebug.Checked Then mnuDebug.Checked = False _
  Else mnuDebug.Checked = True
End Sub

Public Sub TabStrip1_Click()
    setMyRole
End Sub

Private Function setMyRole() As listStyle
    X = TabStrip1.SelectedItem.index
    For i = 1 To 4: lv(i).Visible = False: Next
    lv(X).Visible = True
    myRole = X
End Function


Public Sub DeleteMessage(filePath As String, box As listStyle)
    
    Dim txt, from, subj
    For i = 1 To lv(box).ListItems.Count
        With lv(box).ListItems(i)
            If filePath = .key Then
               txt = .Text
               from = .SubItems(1)
               subj = .SubItems(2)
               Call lv(box).ListItems.Remove(i)
               Exit For
            End If
        End With
    Next
    
    If myRole = trash Or uc.Prefs.useTrash = False Then
        fso.Delete filePath
    Else
        newPath = fso.Move(filePath, uc.folders.trash)
        With lv(3).ListItems
             i = .Count + 1
            .Add i, newPath, txt
            .Item(i).SubItems(1) = from
            .Item(i).SubItems(2) = subj
        End With
    End If
End Sub

Public Function loadTOC(l As ListView, it As listStyle)
  
  With l.ListItems
    .Clear
    toc = Library.getTocPath(it)
    info = fso.readFile(toc)
    info = Split(info, vbCrLf)
         
    For i = 0 To UBound(info)
        Sect = Split(info(i), Chr(5))
        If UBound(Sect) = 3 Then
            'key=filepath, text=status (S,Q,<>)
            'subitem 1=from, 2=subject
            .Add i + 1, Sect(3), Sect(0)
            For j = 1 To 2
                .Item(i + 1).SubItems(j) = Sect(j)
            Next
        End If
    Next
  End With

End Function


Private Sub Lv_dblClick(index As Integer)
On Error GoTo oops
  Me.Caption = " Qmail Folders"
  With lv(index).SelectedItem
    Dim path As String
    path = .key
    
    If .Text <> "S" And .Text <> "Q" Then .Text = Empty
    
    If FileLen(path) > 58500 Then
          msg = "The mail is to large to fully view with this mail client" & vbCrLf & "Would you like to open it in your default editor?"
          If MsgBox(msg, vbYesNo, "Mail to big") = vbYes Then
             Shell uc.folders.editor & " " & path, vbNormalFocus
             Exit Sub
          End If
    End If
    
    If myRole = outbx Then
        Dim e As frmCompose
        Set e = New frmCompose
        Qued = IIf(.Text = "Q", True, False)
        Call e.loadSentMail(path, CBool(Qued), .index)
        e.Show
        Set e = Nothing
    Else
        Dim d As frmRead
        Set d = New frmRead
        Call d.loadMail(path, myRole)
        d.Show
        Set d = Nothing
    End If
  End With
  Exit Sub
oops:
    If Err.Number = 53 Then
        MsgBox "Oops, Looks like you have a corrupt TOC" & vbCrLf & "This file path wasnt found..removing entry ", vbInformation
        lv(index).ListItems.Remove (lv(index).SelectedItem.index)
    Else
         MsgBox "Unknown err in lv_dblclick: " & Err.Description, vbExclamation
    End If
End Sub


'---------------  Menu functions ------------------------
'--------------------------------------------------------

Private Sub mnuDelete_Click()
    fpath = lv(myRole).SelectedItem.key
    If FileExists(fpath) Then
        Call DeleteMessage(CStr(fpath), myRole)
    Else
        'let the err handling in this function deal with it
        Lv_dblClick CInt(myRole)
    End If
End Sub

Private Sub mnuNewMessage_Click()
   Library.ComposeNewMail Library.MonitorClipboard
End Sub

Private Sub mnuEmptyTrash_Click()
   Library.EmptyTrash
   lv(trash).ListItems.Clear
End Sub

Private Sub mnuOption_Click()
  Call ShellAndWait("notepad """ & uc.folders.IniFile & """", vbNormalFocus)
  Call Startup.reloadConfig
End Sub

Private Sub mnuRecp_Click(index As Integer)
   On Error GoTo blah 'if they edited ini and removed one
   Library.ComposeNewMail uc.recipants(index)
   Exit Sub
blah: mnuRecp(index).Visible = False
End Sub

Private Sub mnuAutoCheck_Click()
   If mnuAutoCheck.Checked Then
      mnuAutoCheck.Checked = False
      tmrAutoCheck.Enabled = False
   Else
      mnuAutoCheck.Checked = True
      tmrAutoCheck.SetMinutes uc.Prefs.AutoCheckDelay
      tmrAutoCheck.Enabled = True
   End If
End Sub

'--------------  General form events --------------------
'--------------------------------------------------------

Private Sub Form_Unload(Cancel As Integer)
    Call Library.SaveTocChanges(lv(1), inbox)
    Call Library.SaveTocChanges(lv(2), outbx)
    Call Library.SaveTocChanges(lv(3), trash)
    Call Library.SaveTocChanges(lv(4), saved)
    
    n = "frmMessage"
    SaveSetting App.Title, n, "hWnd", 0
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, n, "MainLeft", Me.Left
        SaveSetting App.Title, n, "MainTop", Me.Top
        SaveSetting App.Title, n, "MainWidth", Me.Width
        SaveSetting App.Title, n, "MainHeight", Me.Height
    End If
    
    winApi.Unhook
    
    End
    'For i = 0 To Forms.Count - 1
    '    Unload Forms(i)
    'Next


End Sub

Private Sub Form_Load()

    If App.PrevInstance Then
       'cant have 2 instances or can loose messages from toc overwrites
       If Len(Command) = 0 Then winApi.ShowPreviousInstance _
       Else winApi.SendMsgToPreviousInstance Command
       End
    End If
    
    If Not IsIde() Then winApi.Hook Me.hWnd
    
    n = "frmMessage"
    SaveSetting App.Title, n, "hWnd", CStr(Me.hWnd)
    Me.Left = GetSetting(App.Title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.Title, n, "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, n, "MainHeight", 6500)
    Me.Caption = " Qmail v.8"
        
    For i = 2 To 4
      lv(i).Width = lv(1).Width
      lv(i).Height = lv(1).Height
      lv(i).Left = lv(1).Left
      lv(i).Top = lv(1).Top
    Next
    
    Call loadTOC(lv(1), inbox)
    Call loadTOC(lv(2), outbx)
    Call loadTOC(lv(3), trash)
    Call loadTOC(lv(4), saved)
    
    Call loadDynamicMenus
    
    setMyRole
End Sub

Sub loadDynamicMenus()
   
   If AryIndexExists(uc.recipants, 1) Then
      For i = 1 To UBound(uc.recipants)
         If i > mnuRecp.Count Then Load mnuRecp.Item(i)
         mnuRecp.Item(i).Caption = uc.recipants(i)
         mnuRecp.Item(i).Enabled = True
         mnuRecp.Item(i).Visible = True
      Next
   End If
    
    For i = 1 To UBound(uc.Users)
        If i > mnuChkAccount.Count Then Load mnuChkAccount.Item(i)
        If i <> UBound(uc.Users) Then
            mnuChkAccount.Item(i).Caption = uc.Users(i).user
        Else
            mnuChkAccount.Item(i).Caption = "All Accounts"
        End If
        mnuChkAccount.Item(i).Enabled = True
        mnuChkAccount.Item(i).Visible = True
    Next
    
    mnuAutoCheck.Caption = "Check every " & uc.Prefs.AutoCheckDelay & " min."
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
       If UnloadMode = 0 And Forms.Count > 1 + uc.SentMessages Then
            msg = "There are still Windows Open are you SURE you want to exit now?"
            If MsgBox(msg, vbExclamation + vbYesNo) = vbNo Then Cancel = -1
       End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Me.Width > 6000 Then
    TabStrip1.Width = Me.Width - TabStrip1.Left - 100
    For i = 1 To 4
        lv(i).Width = TabStrip1.Width - lv(i).Left - 100
    Next
  End If
  If Me.Height > 3000 Then
     TabStrip1.Height = Me.Height - TabStrip1.Top - 400
     For i = 1 To 4
         lv(i).Height = TabStrip1.Height - lv(i).Top
     Next
  End If
End Sub

Private Sub lv_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub tmrAutoCheck_Activate()
    Me.Caption = " Checking..."
    Library.checkMail 1, False, True
    c = 0
    For i = 1 To lv(1).ListItems.Count
       If Left(lv(1).ListItems(i).Text, 1) = "<" Then c = c + 1
    Next
    Me.Caption = "  " & c & " Unread Messages"
End Sub

Private Sub tmrAutoCheck_Tick()
    mnuAutoCheck.Caption = "Check due in " & tmrAutoCheck.TicksTillTrigger & " min."
End Sub

Private Sub ArchieveMail()
    With lv(myRole).SelectedItem
        i = .index
        c = lv(4).ListItems.Count + 1
        newfile = uc.folders.saved & FileNameFromPath(.key)
        fso.Move .key, uc.folders.saved
        lv(4).ListItems.Add c, newfile, .Text
        lv(4).ListItems(c).SubItems(1) = .SubItems(1)
        lv(4).ListItems(c).SubItems(2) = .SubItems(2)
        lv(myRole).ListItems.Remove (i)
    End With
End Sub
