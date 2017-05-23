VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "SimpleFtp   "
   ClientHeight    =   4080
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRaw 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   165
      TabIndex        =   6
      ToolTipText     =   "Send raw command"
      Top             =   3650
      Width           =   8415
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   5535
      TabIndex        =   3
      Top             =   -45
      Width           =   3150
      Begin VB.ComboBox comboFTP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Form1.frx":030A
         Left            =   90
         List            =   "Form1.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   150
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1620
         TabIndex        =   4
         Top             =   165
         Width           =   1425
      End
   End
   Begin VB.Timer tmrPause 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   8805
      Top             =   750
   End
   Begin VB.TextBox txtftp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      TabIndex        =   2
      Top             =   90
      Width           =   5310
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1860
      Left            =   150
      TabIndex        =   1
      Top             =   525
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   3281
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      _Version        =   393217
      Icons           =   "image"
      SmallIcons      =   "image"
      ColHdrIcons     =   "image"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   7408
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Permissions"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList image 
      Left            =   8865
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":062A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0946
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C62
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":129A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2445
      Width           =   8460
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh File List"
      End
      Begin VB.Menu mnuQView 
         Caption         =   "Quick View File"
      End
      Begin VB.Menu mnuspacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRnmFile 
         Caption         =   "Rename File"
      End
      Begin VB.Menu mnuMkFldr 
         Caption         =   "Make Folder"
      End
      Begin VB.Menu spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBonus 
         Caption         =   "Bonus Items"
         Begin VB.Menu mnuBonusItem 
            Caption         =   "Copy File Names"
            Index           =   0
         End
         Begin VB.Menu mnuBonusItem 
            Caption         =   "Create Index.html"
            Index           =   1
         End
      End
      Begin VB.Menu mnumode 
         Caption         =   "Set Mode"
         Begin VB.Menu mnuPasv 
            Caption         =   "PASV"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPort 
            Caption         =   "PORT"
         End
         Begin VB.Menu mnuAppend 
            Caption         =   "APPE"
         End
         Begin VB.Menu mnuNOOP 
            Caption         =   "NOOP"
         End
      End
      Begin VB.Menu spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetDL 
         Caption         =   "Set DL Dir"
      End
   End
   Begin VB.Menu mnuBookmarks 
      Caption         =   "mnuBookmarks"
      Visible         =   0   'False
      Begin VB.Menu mnuAddBookmark 
         Caption         =   "Bookmark Sight"
      End
      Begin VB.Menu mnuDeleteBookmark 
         Caption         =   "Delete Bookmark"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cfgfile As String       'path to config file (bookmarked sights)
Dim DefaultSightIndex As Integer  'upload to this bookmark on icon drop
Dim WaitToUpload As Boolean  'junk hack to deal with new timer code :-\

Dim favs()         'bookmarked sights connection strings

Private Sub Form_Load()
    oFtp.ConnectionMode = cPasv
    cfgfile = App.path & "\ftp.txt"
    
    DefaultSightIndex = -1
    Call loadsights
    
    n = "frmFTP"
    Me.Left = GetSetting(App.Title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.Title, n, "MainWidth", 9000)
    Me.Height = GetSetting(App.Title, n, "MainHeight", 6500)
    
    dlDir = GetSetting(App.Title, n, "DLto", "c:\windows\desktop\")
    oFtp.useNOOP = CBool(GetSetting(App.Title, n, "useNoop", 0))
    mnuNOOP.Checked = oFtp.useNOOP
    
    If Len(Command) > 0 Then 'upload from file dropped on icon
        If InStr(1, Command, "fudd:", 1) > 0 Then
            handleProtocol Command
        Else
            handleFileDrop Command
        End If
    ElseIf DefaultSightIndex > -1 Then
        comboFTP.ListIndex = DefaultSightIndex
    End If
    
End Sub

Sub handleProtocol(ByRef args)
   Me.Visible = True
   args = Replace(args, "fudd://", "ftp://")
   args = Replace(args, "fudd:", "ftp://")
   txtftp = args
   Command1_Click
End Sub

Sub handleFileDrop(arg)
        If DefaultSightIndex = -1 Then MsgBox "To use automatic upload functionality you must set a default sight by editing the ftp.txt file and putting as astrics (*) as the first char of the line of the bookmark to use by default", vbInformation: Exit Sub
        Dim s(1) As String
        s(1) = Replace(arg, """", "")
        If Not FileExists(s(1)) Then MsgBox "Can only Upload Files": Exit Sub
        comboFTP.ListIndex = DefaultSightIndex
        Me.Visible = True
        Command1_Click
        WaitToUpload = True
        While WaitToUpload
            DoEvents: DoEvents: DoEvents
        Wend
        uploadEngine s()
End Sub

Public Sub loadsights()
    'format of cfg file is = 'menominic name <tab> connection string'
    'default sight starts with * for menomic name...if > 1 then last one is used
    If FileExists(cfgfile) Then
        comboFTP.Clear
        txtftp = Empty
        dat = readFile(cfgfile)
        dat = Split(dat, vbCrLf)
        favs() = skinny(dat)
        For i = 0 To UBound(favs)
             On Error Resume Next
             nme = Mid(favs(i), 1, InStr(favs(i), vbTab) - 1)
             comboFTP.AddItem nme, i
             If Left(nme, 1) = "*" Then DefaultSightIndex = i
        Next
    End If
End Sub

Private Sub Command1_Click()
 On Error GoTo warn
    If Command1.Caption = "Connect" Then
      Command1.Caption = "Disconnect"
      If txtftp = Empty Then MsgBox "No Sight to Connect to!", vbCritical: Exit Sub
      txtLog = Empty
      ftp.Connect txtftp
      'geocities waits and sends another packet after connect with quota
      'data...this unexpected data is then intercepted in mid showfilelist
      'which sucks because showlist opens data transfer and that fx cant
      'handle the unexpected data : (  <--took forever to find that out!
      tmrPause.Enabled = True 'waits just in case quota data is sent :-\
    Else
      Me.Caption = "SimpleFtp   "
      Command1.Caption = "Connect"
      ftp.Disconnect
      lv.ListItems.Clear
    End If
 Exit Sub
warn: MsgBox err.Description, vbCritical, "Caught in Command1 Click"
End Sub

Private Sub lv_DblClick()
 On Error GoTo nope
    i = lv.SelectedItem.Index
    If f(i).ftype = fldr Or i = 1 Then 'f(1).name always = ".."
       ftp.ChangeDirectory f(i).fname
       Call showFileList
    Else
       If f(i).fname = "" Then Exit Sub
       ftp.GetFile f(i).fname, dlDir, f(i).byteSize
    End If
    
Exit Sub
nope: MsgBox err.Description, vbCritical, "Error"
End Sub


Sub showFileList()
 On Error GoTo nope
    lv.ListItems.Clear
    f() = ftp.ListContents 'returns base 2 ftpfile() array
    lv.ListItems.Add 1, , "Up to Parent Folder..", , 8
    
    For i = 2 To UBound(f)
            lv.ListItems.Add i, , f(i).fname, , f(i).ftype
            b = f(i).byteSize
            If Len(b) > 5 Then b = Round(b / 1000000, 2) & " M"
            lv.ListItems.Item(i).SubItems(1) = b
            lv.ListItems.Item(i).SubItems(2) = f(i).permissions
    Next
    lv.Refresh

Exit Sub
nope: MsgBox err.Description, vbCritical, "err # " & err.Number & "Connection Error"
End Sub

Private Sub uploadEngine(fpaths() As String)
On Error GoTo noperm
        For i = 1 To UBound(fpaths)
            Index = i
            If AlreadyExists(fpaths(i), Index) Then
                If mnuAppend.Checked = True Then
                    If f(Index).byteSize < FileLen(fpaths(i)) Then
                        bleft = FileLen(fpaths(i)) - f(Index).byteSize
                        bleft = IIf(Len(bleft) > 5, Round(bleft / 1000000, 1) & " M", Round(bleft / 1000, 1) & " Kb")
                        ans = MsgBox("A file of this name already exists on the server but is smaller ...would you like to resume a previous upload? There are " & bleft & " Left to upload", vbOKCancel + vbInformation)
                        If ans = vbOK Then ftp.AppendFile f(Index).fname, fpaths(i), f(Index).byteSize
                    Else
                        MsgBox "A file of this name already exists on the server and is at least the same size as the file you are trying to upload...exiting", vbCritical
                    End If
                Else 'we are replaceing old files w/ new versions
                    ftp.RemoveFile f(Index).fname
                    ftp.PutFile fpaths(i)
                End If
            Else
                If FileExists(fpaths(i)) Then
                   Call ftp.PutFile(fpaths(i))
                ElseIf FolderExists(fpaths(i)) Then
                   MsgBox "Folder transfers not yet supported", vbInformation
                End If
            End If
        Next
        Call showFileList
 Exit Sub
noperm: MsgBox err.Description, vbCritical

End Sub


'export data to upload engine so we can use default upload functionality
Private Sub lv_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim s() As String
   For i = 1 To data.Files.Count
        ReDim Preserve s(i)
        s(i) = data.Files(i)
   Next
   uploadEngine s()
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuPopup
End Sub

Private Sub mnuAddBookmark_Click()
   Dim lst As String
   If txtftp = Empty Then MsgBox "There is no info to save right now!", vbCritical: Exit Sub
   If Not AryIsEmpty(favs) Then
        For i = 0 To UBound(favs)
           If InStr(favs(i), txtftp) > 0 Then
             MsgBox "This Connection string is already saved!", vbExclamation
             Exit Sub
           End If
        Next
        lst = Join(favs, vbCrLf)
   End If
   fname = InputBox("Enter a name you will recgonize this sight by")
   If Trim(fname) = Empty Then MsgBox "Save Failed..You must enter a name try again", vbExclamation: Exit Sub
    
   lst = lst & vbCrLf & fname & vbTab & Replace(txtftp, "ftp://", "")
   Call writeFile(cfgfile, CStr(lst))
   Call loadsights
   MsgBox "Sight added Successfully !", vbInformation
End Sub

Private Sub mnuBonusItem_Click(Index As Integer)
On Error GoTo out
    Dim tmp()
    For i = 1 To lv.ListItems.Count
        t = lv.ListItems(i).Text
        Select Case Index
            Case 0: 'copy file names
                 push tmp(), t
            Case 1: 'create index.html
                 push tmp(), "<a href='" & t & "'>" & t & "</a>"
        End Select
    Next
    Clipboard.Clear
    Clipboard.SetText Join(tmp, vbCrLf)
    MsgBox "Done, data saved in clipboard", vbInformation
out:
End Sub

Private Sub mnuDeleteBookmark_Click()
    If comboFTP.Text = Empty Then MsgBox "Nothing to delete !", vbCritical: Exit Sub
    If MsgBox("Are you sure you want to delete " & comboFTP.Text & " ?", vbYesNo + vbExclamation) = vbYes Then
        favs(comboFTP.ListIndex) = ""
        favs = skinny(favs)
        Call writeFile(cfgfile, Join(favs, vbCrLf))
        Call loadsights
        MsgBox "Sight Successfully removed.", vbInformation
    End If
End Sub

Private Sub mnuDelete_Click()
  i = lv.SelectedItem.Index
  If i = 1 Then Exit Sub
  If f(i).ftype = fldr Then
        If Not ftp.RemoveDirectory(f(i).fname) Then MsgBox ftp.ErrReason, vbCritical, "Error Deleting File" _
         Else Call showFileList
  Else
        If Not ftp.RemoveFile(f(i).fname) Then MsgBox ftp.ErrReason, vbCritical, "Error Deleting File" _
         Else Call showFileList
  End If
End Sub

Private Sub mnuNOOP_Click()
    If mnuNOOP.Checked Then
        mnuNOOP.Checked = False
        oFtp.useNOOP = False
        frmFTP.tmrKeepAlive.Enabled = False
    Else
        mnuNOOP.Checked = True
        oFtp.useNOOP = True
    End If
End Sub

Private Sub mnuPasv_Click()
  ftp.ConnectionMode = cPasv
  mnuPasv.Checked = True
  mnuPort.Checked = False
End Sub

Private Sub mnuPort_Click()
  ftp.ConnectionMode = cPort
  mnuPasv.Checked = False
  mnuPort.Checked = True
End Sub

Private Sub mnuAppend_click()
    If mnuAppend.Checked Then mnuAppend.Checked = False _
    Else: mnuAppend.Checked = True
End Sub

Private Sub mnuQView_Click()
   On Error GoTo out
    i = lv.SelectedItem.Index
    If i < 2 Then Exit Sub
    If f(i).ftype <> txt And f(i).ftype <> unknown Then MsgBox "Ughh QuickView is for text files only :P", vbExclamation: Exit Sub
    t = ftp.QuickView(f(i).fname, f(i).byteSize)
    If t = Empty Then Exit Sub
    Dim q As New frmQuickView
    q.Show
    q.txtMsg = t
    q.RemoteFileName = f(i).fname
out:
End Sub

Private Sub mnuMkFldr_Click()
    fname = InputBox("Enter Name of Folder to Create", "Create Folder", "NewFolder")
    If fname <> "" Then
       If Not ftp.MakeDirectory(fname) Then MsgBox ftp.ErrReason, vbCritical, "Error Creatign Directory" _
        Else Call showFileList
    End If
End Sub

Private Sub mnuRefresh_Click()
   Call showFileList
End Sub

Private Sub mnuSearch_Click()
    If lv.ListItems.Count < 2 Then
        MsgBox "There are no files to search through!", vbExclamation
        Exit Sub
    End If
    frmSearch.Show
End Sub

Private Sub mnuRnmFile_Click()
    i = lv.SelectedItem.Index
    If i = 1 Or f(i).ftype = fldr Then Exit Sub
    oldname = f(i).fname
    newName = InputBox("Enter New Name for " & oldname, "Rename FIle", oldname)
    If newName <> "" Then
       If Not ftp.RenameFile(oldname, newName) Then MsgBox ftp.ErrReason, vbCritical, "Error Renaming File"
    End If
End Sub

Private Sub mnuSetDL_Click()
    old = GetSetting(App.Title, "frmFTP", "DLto", "c:\windows\desktop")
    msg = "Please enter the full path to the folder you wish downloads to be saved in.Sorry this is so incovient but it really only has to be set once"
    X = InputBox(msg, "Set Download Directory", old)
    If FolderExists(X) Then
        dlDir = IIf(Right(X, 1) = "\", X, X & "\")
        SaveSetting App.Title, "frmFTP", "DLto", dlDir
    Else
        MsgBox "Sorry specified folder does not exist try again!", vbCritical
        mnuSetDL_Click
    End If
End Sub


Private Sub comboFTP_Click()
    On Error Resume Next
    dat = favs(comboFTP.ListIndex)
    txtftp = Mid(dat, InStr(dat, vbTab) + 1, Len(dat))
End Sub

Private Sub tmrPause_Timer()
   showFileList
   tmrPause.Enabled = False
   WaitToUpload = False
End Sub

Private Sub txtftp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        LockWindowUpdate txtftp.hWnd
        txtftp.Enabled = False
        DoEvents
        PopupMenu mnuBookmarks
        txtftp.Enabled = True
        LockWindowUpdate 0&
    End If
End Sub

Private Sub txtLog_DblClick()
    If txtLog = Empty Then Exit Sub
    Dim q As New frmQuickView
    q.Show
    q.txtMsg = txtLog
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Me.Width > 8130 Then
        lv.Width = Me.Width - 400
        lv.ColumnHeaders(1).Width = lv.Width - 3500
        txtLog.Width = lv.Width
        txtRaw.Width = lv.Width
        Frame1.Left = lv.Width - Frame1.Width
        txtftp.Width = Frame1.Left - txtftp.Left
  End If
  If Me.Height > 3700 Then
        lv.Height = (0.7 * Me.Height) - lv.Top - 400
        txtLog.Top = lv.Top + lv.Height + 100
        txtRaw.Top = Me.Height - 500 - txtRaw.Height
        txtLog.Height = txtRaw.Top - txtLog.Top - 150
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    n = "frmFTP"
    SaveSetting App.Title, n, "useNoop", oFtp.useNOOP
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, n, "MainLeft", Me.Left
        SaveSetting App.Title, n, "MainTop", Me.Top
        SaveSetting App.Title, n, "MainWidth", Me.Width
        SaveSetting App.Title, n, "MainHeight", Me.Height
    End If
    End
End Sub

Private Function AlreadyExists(Localfile, Index) As Boolean 'Locfile = fullpath to file
   fname = FileNameFromPath(Localfile)
   For j = 1 To UBound(f)
       If f(j).fname = fname And f(j).ftype <> fldr Then
            Index = j 'this passes back the value of j into index!
            AlreadyExists = True
            Exit Function
       End If
   Next
   AlreadyExists = False
End Function

Sub ChangedCurrentFolder(nextSubFolder)
    m = Me.Caption
    If nextSubFolder = ".." Then 'remove last folder from path
        If InStr(m, "/") > 1 Then Me.Caption = Mid(m, 1, InStrRev(m, "/") - 1) _
        Else Me.Caption = "SimpleFtp   "
    Else                         'add new folder to path
        s = "/" 'so we can use this to set full paths too
        If InStr(nextSubFolder, "/") > 0 Then s = ""
        Me.Caption = m & s & nextSubFolder
    End If
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
 On Error GoTo out
    Select Case KeyAscii
       'Case Else: MsgBox KeyAscii
    End Select
out:
End Sub

Private Sub txtRaw_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo oops
    If KeyCode = 13 And txtRaw <> Empty Then
         Dim cmd, arg, args() As String
         fs = InStr(txtRaw, " ")
         If fs > 0 Then
            cmd = Mid(txtRaw, 1, fs - 1)
            arg = Mid(txtRaw, fs + 1, Len(txtRaw))
            args = Split(arg, " ")
         Else
            cmd = txtRaw
         End If
         Select Case UCase(cmd)
            Case "MKDIR": ftp.MakeDirectory arg
            Case "GET":  FindNameFromRegExp arg, True: lv_DblClick
            Case "PUT": ftp.PutFile arg
            Case "RM", "DEL", "D": ftp.RemoveFile arg
            Case "RMDIR": ftp.RemoveDirectory arg
            Case "RNTO", "RN", "RNFR": ftp.RenameFile args(0), arg(1)
            Case "QUIT", "EXIT", "OPEN", "CONNECT": Command1_Click
            Case "SEARCH", "FIND", "S", "F": frmSearch.PredefinedSearch arg
            Case "CWD", "CD": ftp.ChangeDirectory arg: showFileList
            Case "LIST", "LS": Call showFileList
            Case "VIEW", "V": FindNameFromRegExp arg, True: mnuQView_Click
            Case Else: ftp.SendRawCommand txtRaw
         End Select
         txtRaw = Empty
    End If
    Exit Sub
oops: MsgBox err.Description, vbCritical
End Sub

Function FindNameFromRegExp(match, Optional giveFocus As Boolean = False)
    If lv.ListItems.Count < 2 Then Exit Function
    Dim find As String
    find = CStr(match) 'lcase can be weird with non strings
    For i = 1 To lv.ListItems.Count
        If LCase(lv.ListItems(i).Text) Like LCase(find) Then
            FindNameFromRegExp = lv.ListItems(i).Text
            If giveFocus Then lv.ListItems(i).Selected = True
            Exit Function
        End If
    Next
End Function
