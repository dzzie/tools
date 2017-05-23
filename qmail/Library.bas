Attribute VB_Name = "Library"
Public Sub SaveTocChanges(lv As ListView, l As listStyle)
    Dim tmp()
    v = Chr(5)
    
    With lv.ListItems
        For i = 1 To .Count
           push tmp, .Item(i).Text & v & .Item(i).SubItems(1) & v & _
                     .Item(i).SubItems(2) & v & .Item(i).key
        Next
    End With
    
    Call fso.writeFile(getTocPath(l), Join(tmp, vbCrLf))
    
End Sub

Public Sub checkMail(userIndex As Integer, Optional debugmode As Boolean = False, Optional hidden = False)
  On Error GoTo warn
    Dim msg() As String 'checkmail returns base 1 array
    
    If userIndex = UBound(uc.Users) Then
        For i = 1 To UBound(uc.Users) - 1
            With uc.Users(i)
                msg() = frmCheckMail.getMail(.Server, .user, .pass, .port, debugmode, hidden)
            End With
            FillBox frmMessages.lv(1), msg()
        Next
    Else
        With uc.Users(userIndex)
                msg() = frmCheckMail.getMail(.Server, .user, .pass, .port, debugmode, hidden)
        End With
        FillBox frmMessages.lv(1), msg()
    End If
    
Exit Sub
warn: MsgBox "err in checkmail() " & Err.Description & Err.Source
End Sub

Public Sub FillBox(lv As ListView, msg() As String)
    If AryIsEmpty(msg) Then Exit Sub
    Dim Outbox As Boolean, t
    If lv.ColumnHeaders(2).Text = "To:" Then Outbox = True
    
    For j = 1 To UBound(msg)
       Dim p As parsed
       p = parseMail(GetMailHeader(msg(j)))
       If Outbox Then t = "?" Else t = IIf(p.atch = 0, "<>", "<A>")
       With lv.ListItems
          i = .Count + 1
          .Add i, msg(j), t
          If Outbox Then .Item(i).SubItems(1) = p.to _
          Else .Item(i).SubItems(1) = p.from
          .Item(i).SubItems(2) = p.subj
       End With
    Next
    
End Sub

Public Sub saveSent(msg As String, Optional Qued As Boolean = False)
     Dim p As parsed, q As String
     
     p = parseMail(msg)
     toc = getTocPath(outbx)
     tf = fso.CreateTempFile(fso.GetParentFolder(toc), ".txt")
     fso.writeFile tf, msg
     If Qued Then q = "Q" Else q = "S"
     With frmMessages.lv(2).ListItems
        i = .Count + 1
        .Add i, tf, q
        .Item(i).SubItems(1) = p.to
        .Item(i).SubItems(2) = p.subj
     End With
End Sub


Public Function getTocPath(style As listStyle)
    
    Dim box As String, toc As String
    Select Case style
        Case inbox
             box = endSlash(uc.folders.inbox)
             toc = box & "inbox.toc"
        Case outbx
             box = endSlash(uc.folders.oubox)
             toc = box & "outbox.toc"
        Case trash
             box = endSlash(uc.folders.trash)
             toc = box & "trash.toc"
        Case saved
             box = endSlash(uc.folders.saved)
             toc = box & "saved.toc"
    End Select
    
    If Not fso.FolderExists(box) Then fso.buildPath (box)
    If Not fso.FileExists(toc) Then fso.CreateTextFile (toc)
    
    getTocPath = toc
    
End Function


Public Sub ComposeNewMail(Optional recpt = "", Optional subject = "", Optional body = "")
    Dim frmD As frmCompose
    Set frmD = New frmCompose
    frmD.Caption = ""
    frmD.txtProps(0) = uc.Send.sender
    frmD.txtProps(1) = recpt
    frmD.txtProps(2) = subject
    frmD.txtMsg = body
    frmD.Show
    Set frmD = Nothing
End Sub

Public Sub EmptyTrash()
    Call fso.DeleteFolder(uc.folders.trash, True)
    Call Library.getTocPath(trash)
End Sub

Function GetMailHeader(path)
  On Error GoTo out
    tmp = ""
    If FileLen(path) < 10 Then Exit Function
    f = FreeFile
    Open path For Input As f
    While Right(tmp, 4) <> vbCrLf & vbCrLf
       Line Input #f, it
       tmp = tmp & it & vbCrLf
    Wend
    Close #f
out:
    GetMailHeader = tmp
End Function

Public Function parseMail(msg) As parsed
    On Error Resume Next
    Dim p As parsed
    p.body = msg
    p.atch = 0
    If InStr(p.body, "boundary=""") > 0 Then p.atch = 1
    tmp = Split(msg, vbCrLf)
      
      For j = 0 To UBound(tmp)
          col = InStr(1, tmp(j), ":") + 1
          
          If InStr(1, Left(tmp(j), 7), "From", vbTextCompare) > 0 Then
              p.from = Trim(Mid(tmp(j), col, Len(tmp(j))))
          End If
            
          If InStr(1, Left(tmp(j), 9), "Subject", vbTextCompare) > 0 Then
              p.subj = Trim(Mid(tmp(j), col, Len(tmp(j))))
          End If
          
          If InStr(1, Left(tmp(j), 5), "To", vbTextCompare) > 0 Then
              p.to = Trim(Mid(tmp(j), col, Len(tmp(j))))
          End If
          
          If p.from <> "" And p.subj <> "" And p.to <> "" Then Exit For
      Next
    
      parseMail = p
End Function

Function MonitorClipboard() As String
    X = FirstLine(LTrim(Trim(Clipboard.GetText)))
    If InStr(X, "@") > 0 Then
        sp = InStr(2, X, " ")
        If sp > 0 Then MonitorClipboard = Mid(X, 1, sp - 1) _
        Else MonitorClipboard = CStr(X)
    End If
End Function

Sub RebuildTocFromFiles(lv As ListView, box As listStyle)
  On Error GoTo out
    lv.ListItems.Clear
    toc = Library.getTocPath(box)
    If fso.FileExists(toc) Then Kill toc
    pf = fso.GetParentFolder(toc)
    Dim dc As DirectoryContents
    dc = fso.GetDirContents(pf, "txt")
    Dim fnames() As String
    ReDim fnames(1 To dc.sFiles.Count)
    For i = 1 To dc.sFiles.Count
        fnames(i) = CStr(pf & dc.sFiles(i))
    Next
    FillBox lv, fnames()
    Call SaveTocChanges(lv, box)
    Exit Sub
out: MsgBox "There are no files in this folder", vbInformation
End Sub

Function DeQuote(it)
  On Error GoTo out
    Dim f(): Dim c()
    n = it
    If InStr(n, "=") > 0 Then
        t = Split(n, "=")
        For i = 0 To UBound(t)
            a = Left(t(i), 2)
            b = IsHex(a)
            If b <> Empty Then
                push f(), "=" & a
                push c(), b
            End If
        Next
        If Not AryIsEmpty(f) Then
            For i = 0 To UBound(f): n = Replace(n, f(i), c(i)): Next
        End If
        n = Replace(n, "=" & vbCrLf, "")
    End If
    DeQuote = n
Exit Function
out: DeQuote = it: MsgBox Err.Description, vbCritical, "Err in deQuote"
End Function

Function Escape(it)
    Dim f(): Dim c()
    n = Replace(it, "+", " ")
    If InStr(n, "%") > 0 Then
        t = Split(n, "%")
        For i = 0 To UBound(t)
            a = Left(t(i), 2)
            b = IsHex(a)
            If b <> Empty Then
                push f(), "%" & a
                push c(), b
            End If
        Next
        For i = 0 To UBound(f)
            n = Replace(n, f(i), c(i))
        Next
    End If
    Escape = n
End Function

Public Function breakIt(it)
    breakIt = uc.Prefs.ReplyChar & Replace(it, vbCrLf, vbCrLf & uc.Prefs.ReplyChar)
End Function

Sub ParseCommandLine(arg)
    Dim t() As String
    arg = BatchReplace(arg, """-,?-&,mailto:-")
    t = Split(arg, "&")
    subj = Escape(StrFindValFromKey(t, "subject"))
    body = Escape(StrFindValFromKey(t, "body"))
    Call ComposeNewMail(t(0), subj, body)
End Sub
