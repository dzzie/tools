VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRead 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   9375
   Begin VB.CheckBox chkShow 
      Caption         =   "Show Body"
      Height          =   270
      Left            =   3750
      TabIndex        =   7
      Top             =   165
      Width           =   1155
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   315
      Left            =   6765
      TabIndex        =   6
      Top             =   75
      Width           =   885
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   5205
      TabIndex        =   5
      Top             =   105
      Width           =   1485
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Goto Url"
      Height          =   330
      Index           =   2
      Left            =   1275
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   9780
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Reply"
      Height          =   330
      Index           =   3
      Left            =   2625
      TabIndex        =   3
      Top             =   75
      Width           =   1065
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Message"
      Height          =   330
      Index           =   4
      Left            =   7785
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      UseMaskColor    =   -1  'True
      Width           =   1425
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save Mail"
      Height          =   330
      Index           =   1
      Left            =   75
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   75
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3825
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   495
      Width           =   9225
   End
   Begin VB.Menu mnuAdvanced 
      Caption         =   "Advanced"
      Visible         =   0   'False
      Begin VB.Menu mnuUrl 
         Caption         =   "Goto Url"
      End
      Begin VB.Menu mnuSaveSel 
         Caption         =   "Save Sel"
      End
      Begin VB.Menu mnuBits 
         Caption         =   "To BitFile"
         Begin VB.Menu mnuBitFiles 
            Caption         =   "[Empty]"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu spacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookMark 
         Caption         =   "BookMark Url"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "Download Url"
      End
      Begin VB.Menu mnuAddRcpt 
         Caption         =   "Add As Recpt"
      End
      Begin VB.Menu spacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDecode 
         Caption         =   "Decode Text"
      End
      Begin VB.Menu mnuDecodeFile 
         Caption         =   "Decode File"
      End
      Begin VB.Menu mnuHeaders 
         Caption         =   "Full Headers"
      End
      Begin VB.Menu spacer2 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Private Type mail
    attachment As String
    subject As String
    body As String
    header As String
    path As String
    fromBox As listStyle
    replyTo As String
    showHeader As Boolean
End Type

Private Type find
    position As Long
    place As Long
    what As String
End Type

Dim find As find
Dim mail As mail

Private Sub cmd_Click(index As Integer)
   With cmd(index)
     Select Case index
       Case 1 'save whole mail
                msg1 = "Doh! you alreaded saved a file with this name and this exact file size here :)\"
                msg2 = "Oops that name is taken with another file of the same name but different size. Do you want to append this message to the other file?"
                fname = getFileName(uc.folders.saveTo, mail.subject)
                If Len(fname) = 0 Then: MsgBox "Save Failed No filename!": Exit Sub
                If fso.FileExists(fname) Then
                    If (FileLen(fname) - 2) = Len(mail.body) Then 'writeFile adds a vbcrlf
                         MsgBox msg1, vbExclamation
                         Exit Sub
                    ElseIf MsgBox(msg2, vbYesNo + vbExclamation) = vbYes Then
                         v = vbCrLf & vbCrLf & vbCrLf
                         v = v & String(50, "-") & v
                         fso.AppendFile fname, v & mail.body
                         Exit Sub
                    End If
                End If
                'note this saves the viewed message so you can determine
                'how it is saved! including headers, dequoting etc..
                fso.writeFile fname, Text1
       Case 2 'go url
                If Text1.SelLength = 0 Then MsgBox "Nothing Selected !": Exit Sub
                Shell uc.folders.browser & " " & Text1.SelText, vbNormalFocus
       Case 3 'reply
                isRe = InStr(1, Left(mail.subject, 5), "Re:", vbTextCompare)
                subj = IIf(isRe > 0, mail.subject, "Re: " & mail.subject)
                body = IIf(chkShow.Value, breakIt(mail.body), Empty)
                If Text1.SelLength > 0 Then
                   seltxt = Text1.SelText
                   If InStr(seltxt, "@") > 0 Then ComposeNewMail Text1.SelText, subj, body _
                   Else ComposeNewMail mail.replyTo, subj, breakIt(seltxt)
                Else
                    Library.ComposeNewMail mail.replyTo, subj, body
                End If
       Case 4 'delete message
                Call frmMessages.DeleteMessage(mail.path, mail.fromBox)
                Unload Me
       End Select
   End With
End Sub


Public Sub loadMail(path As String, fromBox As listStyle)

    mail.path = path
    mail.fromBox = fromBox
    mail.showHeader = False
        
    If FileLen(path) > 58500 Then it = ReadFirst(path, 20000) _
    Else it = readFile(path)
    
    '1 = vbTextCompare
    X = InStr(1, it, vbCrLf & "Subject", 1) + 10
    ex = InStr(X, it, vbCrLf)
    mail.subject = Trim(Mid(it, X, ex - X))
    
    Y = InStr(1, it, vbCrLf & "From:", 1) + 7
    ex = InStr(Y, it, vbCrLf)
    mail.replyTo = Trim(Mid(it, Y, ex - Y))
    mail.header = Mid(it, 1, Y - 6)
    mail.body = Mid(it, Y - 5, Len(it))
        
    Text1 = mail.body
    Me.Caption = mail.subject
    chkShow.Value = IIf(uc.Prefs.MsgInReply, 1, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    n = "frmRead"
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, n, "MainLeft", Me.Left
        SaveSetting App.title, n, "MainTop", Me.Top
        SaveSetting App.title, n, "MainWidth", Me.Width
        SaveSetting App.title, n, "MainHeight", Me.Height
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo warn
    find.what = Empty
    find.place = 0
    find.position = 1
    
    n = "frmRead"
    Me.Left = GetSetting(App.title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.title, n, "MainWidth", 6500)
    Me.Height = GetSetting(App.title, n, "MainHeight", 6500)
    
    Text1.FontSize = uc.fonts.size
    Text1.FontName = uc.fonts.face
    Text1.ForeColor = CLng(uc.fonts.color)
    Text1.backcolor = CLng(uc.fonts.backcolor)
    Text1.FontBold = uc.fonts.bold
    
    If AryIndexExists(uc.bitfiles, 1) Then
      For i = 1 To UBound(uc.bitfiles)
         If i > 1 Then Load mnuBitFiles.Item(i)
         mnuBitFiles.Item(i).Caption = fso.GetFullName(uc.bitfiles(i))
         mnuBitFiles.Item(i).Enabled = True
         mnuBitFiles.Item(i).Visible = True
      Next
    End If
Exit Sub
warn: MsgBox "error in frmread form load : " & Err.Description
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Me.Width > 6000 Then Text1.Width = Me.Width - Text1.Left - 200
  If Me.Height > 3000 Then Text1.Height = Me.Height - Text1.Top - 400
End Sub

Private Sub mnuBitFiles_Click(index As Integer)
    If index = mnuBitFiles.Count Then
       fname = SafeFileName(InputBox("Enter Name for new bit file"))
       If fname = Empty Then Exit Sub
       X = UBound(uc.bitfiles)
       pf = GetParentFolder(uc.bitfiles(1)) & "\"
       Ini.LoadFile uc.folders.IniFile
       Ini.SetValue "bitfiles", "number", X
       Ini.AddKey "bitfiles", "bitfile" & X, pf & fname
       Ini.Save
       Ini.Release
       uc.bitfiles(UBound(uc.bitfiles)) = pf & fname
       push uc.bitfiles, "- Add New -"
       With mnuBitFiles
         .Item(.Count).Caption = fname
         Load .Item(.Count + 1)
         .Item(.Count).Caption = "- Add New -"
       End With
    Else
       If Text1.SelLength > 0 Then
            fso.AppendFile uc.bitfiles(index), vbCrLf & vbCrLf & " " & Text1.SelText & vbCrLf & vbCrLf & String(50, "-")
       Else
            Shell "notepad " & uc.bitfiles(index), vbNormalFocus
       End If
    End If
End Sub



Private Sub mnuDecode_Click()
    On Error GoTo oops
    If Text1.SelLength = 0 Then MsgBox "Nothing Selected !": Exit Sub
    Dim tmp
    seltxt = Text1.SelText
    If InStr(seltxt, " ") < 1 Then 'is a base64 encoded string
        tmp = UnixToDos(b64.DecodeString(seltxt))
    Else 'is a quoted printable string
        tmp = parseHtml(seltxt)
        tmp = DeQuote(tmp)
    End If
    Text1.SelText = tmp
    Exit Sub
oops: MsgBox Err.Description
End Sub

Private Sub mnuDecodeFile_Click()
    On Error GoTo oops
    If Text1.SelLength > 0 Then
        fname = getFileName(uc.folders.attach, mail.replyTo)
        If Len(fname) = 0 Then Exit Sub
        b64.UnMimeStringToFile fname, Text1.SelText
    Else
        fname = getFileName(Empty, mail.subject, "Open File to Decode", False)
        saveAs = getFileName(Empty, fso.GetBaseName(fname) & ".zip", "Save Decoded File As")
        If Len(fname) = 0 Or Len(saveAs) = 0 Then Exit Sub
        b64.UnMimeFileToFile fname, saveAs
        
    End If
    MsgBox "File Decode Complete" & vbCrLf & vbCrLf & "Saved as:  " & fname
    Exit Sub
oops: MsgBox Err.Description
End Sub

Private Sub mnuDownload_Click()
    On Error GoTo out
    seltxt = Text1.SelText
    If Len(seltxt) = 0 Then Exit Sub
    tmp = Split(seltxt, "/")
    fname = getFileName(uc.folders.attach, tmp(UBound(tmp)))
    If DownloadFile(seltxt, fname) Then MsgBox "Download Complete :)" _
    Else MsgBox br("Download Failed: \n\n\t") & seltxt
out:
End Sub

Private Sub cmdFind_Click()
  If Text1 = "" Or txtFind = "" Then Exit Sub
  If find.what = Empty Then find.what = txtFind
  find.place = InStr(find.position, Text1.Text, find.what, vbTextCompare)
  If find.place > 0 Then
    Text1.SetFocus
    Text1.SelStart = find.place - 1
    Text1.SelLength = Len(find.what)
    find.position = find.place + Len(find.what)
  Else
    find.position = 1: find.place = 0: find.what = Empty
    MsgBox "Search Completed..", vbInformation
  End If
End Sub

Private Sub mnuSaveSel_Click()
    If Text1.SelLength < 1 Then MsgBox "Nothing Text Highlighted to try to save!": Exit Sub
    f = getFileName(uc.folders.saveTo, mail.subject)
    If f <> Empty Then
        fso.AppendFile f, Text1.SelText
        MsgBox "Selection Saved Successfully to" & vbCrLf & vbCrLf & f
    End If
End Sub

Private Sub txtFind_Change()
    find.what = Empty
End Sub

Private Sub mnuHeaders_Click()
    If Not mail.showHeader Then
        mail.showHeader = True
        Text1 = mail.header & mail.body
        mnuHeaders.Caption = "Hide Full Headers"
    Else
        mail.showHeader = False
        Text1 = mail.body
        mnuHeaders.Caption = "Show Full Headers"
    End If
End Sub

Private Sub mnuUrl_Click()
   Call cmd_Click(2)
End Sub

Private Sub mnuBookMark_Click()
    'THIS IS IE SPECIFIC ! 'mabey should save to local html file?
    If Not fso.FolderExists(uc.folders.bookmark) Then MsgBox uc.folders.bookmark & br("\n\nDoes not exist Please create it first"): Exit Sub
    seltxt = Text1.SelText
    n = InputBox(br("Enter the Name you want\n\t" & seltxt & "\nSaved as"), , WebFileNameFromPath(seltxt))
    If n = Empty Then Exit Sub
    dat = br("[Default]\nBASEURL=___\n[InternetShortcut]\nURL=___")
    dat = Replace(dat, "___", seltxt)
    path = endSlash(uc.folders.bookmark) & SafeFileName(n) & ".url"
    fso.writeFile path, dat
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
        Case 1: Text1.SelStart = 0: Text1.SelLength = Len(Text1) 'ctrl a -> select all
        Case 2:  MsgBox "Selected string Length: " & Len(Text1.SelText)  'ctrl b -> add letters
        Case 4: Call cmd_Click(2)                     'ctrl d -> delete
        Case 6: mnuHeaders_Click                      'ctrl f -> full headers
        Case 12: Text1.SelText = LCase(Text1.SelText) 'ctrl l -> lcase selection
        Case 21: Text1.SelText = UCase(Text1.SelText) 'ctrl u -> ucase selection
        Case 24: Unload Me                            'ctrl x -> close
        'Case Else: MsgBox KeyAscii
   End Select
End Sub

Public Function DownloadFile(URL, LocalFilename) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, CStr(URL), CStr(LocalFilename), 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function

Function getFileName(startDir, Optional proposedName = "", Optional title = "Save File As", Optional SaveDlg = True) As String
    With CommonDialog
       On Error GoTo skip
            .CancelError = True
            .DialogTitle = title
            .filename = fso.SafeFileName(proposedName)
            .InitDir = startDir
            .filter = "Text Files (*.txt)|*.txt"
            .DefaultExt = ".txt"
            If SaveDlg Then .ShowSave Else .ShowOpen
            getFileName = .filename
            Exit Function
skip: getFileName = ""
    End With
End Function

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then ShowRtClkMenu Me, Text1, mnuAdvanced
End Sub

Private Sub mnuAddRcpt_click()
    eMail = Text1.SelText
    If InStr(eMail, "@") < 1 Then MsgBox eMail & " Not a valid email Address!", vbCritical: Exit Sub
    X = UBound(uc.recipants) + 1
    Ini.LoadFile uc.folders.IniFile
    Ini.SetValue "recipients", "number", X
    Ini.AddKey "recipients", "recp" & X, eMail
    Ini.Save
    Ini.Release
    push uc.recipants, eMail
    frmMessages.loadDynamicMenus
End Sub

