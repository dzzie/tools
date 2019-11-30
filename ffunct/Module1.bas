Attribute VB_Name = "Module1"
Public Type optn
   extensions As String
   renameTo As String
End Type

Public Type Seq
    count As Integer
    Appnd As String
End Type

Global opt As optn
Global Seq As Seq
Global Selopt As Integer
Global Warned As Boolean
Global Together As Boolean         'used in zip & unzip


Public Sub ListEngine(folderpath)
    Dim files() As String
    
    If fso.FileExists(folderpath) Then
        push files, folderpath
    ElseIf fso.FolderExists(folderpath) Then
        'zip the whole folder if conditions right
        If Selopt = 7 And Together Then
            Zip CStr(folderpath)
        Else
            files() = fso.GetFolderFiles(folderpath)
            If Form1.Check1.value = 1 Then
                Dim folders() As String
                folders() = fso.GetSubFolders(folderpath)
                If Not aryIsEmpty(folders) Then
                    For i = 0 To UBound(folders)
                        ListEngine folders(i)
                    Next
                End If
            End If
        End If
    Else
        MsgBox "Specified folder could not be found exiting"
        Exit Sub
    End If
    
    For i = 0 To UBound(files)
      Dim f As String
      f = files(i)
      ext = fso.GetExtension(f) 'has .xxx format need xxx format
      
      If Not exclude(Mid(ext, 2, Len(ext)), opt.extensions) Then
           Select Case Selopt
               Case 0:  Call batchRename(f)
               Case 1:  Call folderJoin(f)
               Case 2:  Call directList(f)
               Case 3:  Call htmlIndex(f)
               Case 4:  Call deScript(f)
               Case 5:  Call deHtml(f)
               Case 6:  Call Unzip(f)
               Case 7:  Call Zip(f)
               Case 8:  Call SetHidden(f)
               Case 9:  Call ShowHidden(f)
               Case 10: Call SetReadOnly(f)
               Case 11: Call UnSetReadOnly(f)
               Case 12: Call ImgIndex(f)
               Case 13: Call SeqName(f)
               Case 14: Call unix2Dos(f)
               Case 15: Call dos2Unix(f)
               Case 16: Call batchReplace(f)
               Case 17: Call PrefixFiles(f)
               Case 18: Call StripFileNames(f)
           End Select
      End If
    Next
    
End Sub

Private Sub batchRename(arg As String)
   If arg = Empty Then Exit Sub
   fso.Rename arg, fso.GetBaseName(arg) & opt.renameTo
End Sub

Private Sub StripFileNames(arg As String)
   If arg = Empty Then Exit Sub
   fso.Rename arg, Replace(fso.FileNameFromPath(arg), Seq.Appnd, Empty, , , vbTextCompare)
End Sub

Private Sub folderJoin(arg As String)
   If InStr(1, arg, "JOINED.txt") > 0 Then Exit Sub
   Dim txt As String
   
    pf = fso.GetParentFolder(arg)
    jf = pf & "\JOINED.txt"
    txt = fso.readFile(arg)
    
    If Not fso.FileExists(jf) Then
        fso.AppendFile jf, "Combined Files of " & pf & "   " & Date & vbCrLf & vbCrLf
    End If
    
    fso.AppendFile jf, String(3, vbCr) _
     & "===============  " & arg & " ==============" & vbCrLf & txt
    
    
End Sub

Private Sub directList(arg As String)
  If InStr(1, arg, "LISTING.txt") > 0 Then Exit Sub
  
  pf = fso.GetParentFolder(arg)
  jf = pf & "\LISTING.txt"
  
  If Not fso.FileExists(jf) Then
    fso.AppendFile jf, "Directory Listing of" & pf & "   " & Date & vbCrLf & vbCrLf
  End If
  
  fso.AppendFile jf, arg & vbCrLf
End Sub

Private Sub htmlIndex(arg As String)
  If InStr(1, arg, "index.html") > 0 Then Exit Sub
  
  pf = fso.GetParentFolder(arg)
  jf = pf & "\index.html"
  nom = fso.GetFullName(arg)
  
  If Not fso.FileExists(jf) Then
     fso.AppendFile jf, "<h3><pre>Contents of   " & LCase(pf) & "          " & Date & "</pre></h3></br>" & vbCrLf & vbCrLf & "<a href='../'>Parent Directory</a><br>" & vbCrLf
  End If
  
  fso.AppendFile jf, "- <a href='" & nom & "'>" & nom & "</a><br>"
 
End Sub

Private Sub deScript(arg As String)
  'pf = fso.GetParentFolder(arg)
  'nom = fso.GetFullName(arg)
  'fso.writeFile pf & "\_" & nom, parseScript(fso.readFile(arg))
  fso.writeFile arg, parseScript(fso.readFile(arg))
End Sub

Private Sub deHtml(arg As String)
  pf = fso.GetParentFolder(arg)
  nom = pf & "\" & fso.GetBaseName(arg)
  fso.writeFile nom & "_html.txt", parseHtml(fso.readFile(arg))
End Sub

Private Sub Unzip(arg As String)
    pf = fso.GetParentFolder(arg)
    bn = fso.GetBaseName(arg)
    If Together Then
      Shell "wzunzip -d """ & arg & """ """ & pf & "\""", vbHide
    Else
      Shell "wzunzip -d """ & arg & """ """ & pf & "\" & bn & """", vbHide
    End If
End Sub

Private Sub Zip(arg As String)
     ext = fso.GetExtension(arg) 'has .xxx format need xxx format
     If exclude(Mid(ext, 2, Len(ext)), opt.extensions) Then Exit Sub
     
     pf = fso.GetParentFolder(arg)
     bn = fso.GetBaseName(arg)
     If Together Then
       Shell "wzzip """ & pf & "\DIRECTORY.ZIP"" """ & pf & """", vbHide
     Else
       Shell "wzzip """ & arg & ".zip"" """ & arg & """", vbHide
     End If
End Sub

Private Sub SetHidden(arg As String)
    fso.SetAttribute arg, vbHidden
End Sub

Private Sub ShowHidden(arg As String)
   fso.SetAttribute arg, vbNormal
End Sub

Private Sub SetReadOnly(arg As String)
   fso.SetAttribute arg, vbReadOnly
End Sub

Private Sub UnSetReadOnly(arg As String)
   fso.SetAttribute arg, vbNormal
End Sub

Private Sub ImgIndex(arg As String)
  
  pf = fso.GetParentFolder(arg)
  n = fso.GetFullName(arg)
  If InStr(1, n, "%20") > 0 Then Call fso.Rename(arg, Replace(n, "%20", ""))
  If InStr(1, n, "%") > 0 Then Call fso.Rename(arg, Replace(n, "%", ""))
  n = fso.GetFullName(arg)
  fso.AppendFile pf & "\Imgindex.html", "<img src='" & n & "' alt='" & n & "'><br>"
  
End Sub


Private Sub SeqName(arg As String)
  ext = fso.GetExtension(arg)
  Seq.count = Seq.count + 1
  Call fso.Rename(arg, Seq.count & Seq.Appnd & ext)
End Sub

Private Sub PrefixFiles(arg As String)
  ext = fso.GetExtension(arg)
  bn = fso.GetBaseName(arg)
  Seq.count = Seq.count + 1
  Call fso.Rename(arg, Seq.Appnd & bn & ext)
End Sub



Private Function setTo(X)
  p = InStr(1, X, "->") + 1
  ed = Mid(X, p, Len(X))
  ed = filt(ed, " ,;,*,-,>")
  If InStr(1, ed, ".") < 1 Then ed = "." & ed
  opt.renameTo = ed
  setTo = Mid(X, 1, p)
End Function

Public Sub Vdate(which)
  it = Form1.Text2
  If InStr(1, it, "->") > 0 Then it = setTo(it)
  it = filt(it, " ,*,-,>,.")
  If InStr(1, it, ";") < 1 Then it = it & ";"
  opt.extensions = Replace(it, ";", ",")
End Sub

Public Function filt(txt, remove As String)
  If Right(txt, 1) = "," Then txt = Mid(txt, 1, Len(txt) - 1)
  tmp = Split(remove, ",")
  For i = 0 To UBound(tmp)
     txt = Replace(txt, tmp(i), "", , , vbTextCompare)
  Next
  filt = txt
End Function

Public Function exclude(test, acpt) As Boolean
    Dim a As Boolean
    a = False
    tmp = Split(acpt, ",")
    For i = 0 To UBound(tmp)
      If LCase(test) = LCase(tmp(i)) Or tmp(i) = "*" Then
            a = True
            Exit For
      End If
    Next
    If InStr(1, Form1.lblMsg, "Exclude") < 1 Then a = Not a
    'MsgBox test & " " & acpt & "  " & a
    exclude = a
End Function

Sub d(it)
  Debug.Print it
End Sub

Sub unix2Dos(arg)
    writeFile arg, UnixToDos(readFile(arg))
End Sub


Sub dos2Unix(arg)
    writeFile arg, Replace(readFile(arg), vbCrLf, vbLf)
End Sub


Function UnixToDos(it) As String
    If InStr(it, vbLf) > 0 Then
        tmp = Split(it, vbLf)
        For i = 0 To UBound(tmp)
            If InStr(tmp(i), vbCr) < 1 Then tmp(i) = tmp(i) & vbCr
        Next
        UnixToDos = Join(tmp, vbLf)
    Else
        UnixToDos = CStr(it)
    End If
End Function


Sub batchReplace(arg)
    s = Split(Seq.Appnd, "->")
    writeFile arg, Replace(readFile(arg), s(0), s(1))
End Sub
