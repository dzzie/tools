Attribute VB_Name = "misc"
Sub push(ary, value)
  On Error GoTo fresh
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
  Exit Sub
fresh: ReDim ary(0): ary(0) = value
End Sub

Function readFile(FileName)
  ff = FreeFile
  Temp = ""
   Open FileName For Binary As #ff        ' Open file.(can be text or image)
     Temp = Input(FileLen(FileName), #ff) ' Get entire Files data
   Close #ff
   readFile = Temp
End Function

Public Sub writeFile(path, it As String)
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, it
    Close ff
End Sub

Function skinny(t, Optional base = 0)  'remove empty elements
    Dim ret()                                'return adjustable base array
    c = base
    For i = base To UBound(t)
      If t(i) <> "" Then
        ReDim Preserve ret(c)
        ret(c) = t(i)
        c = c + 1
      End If
    Next
   
    skinny = ret()
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function FolderExists(path) As Boolean
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function br(it)
 tmp = Replace(it, "\n", vbCrLf)
 tmp = Replace(tmp, "\t", vbTab)
 br = Replace(tmp, "\q", """")
End Function

Function chngArrayBase(t, base)
   Dim ret()
   lbT = LBound(t)
   elem = UBound(t) - LBound(t) + 1
   ReDim ret(base To elem)
   For i = 0 To elem - 1
      ret(base + i) = t(lbT + i)
   Next
   chngArrayBase = ret()
End Function

Function shave(it)
    it = Trim(it)
    shave = LTrim(it)
End Function

Function endSlash(it)
   endSlash = IIf(Right(it, 1) = "\", it, it & "\")
End Function

Function filt(txt, remove As String)
  If Right(txt, 1) = "," Then txt = Mid(txt, 1, Len(txt) - 1)
  tmp = Split(remove, ",")
  For i = 0 To UBound(tmp)
     txt = Replace(txt, tmp(i), "", , , vbTextCompare)
  Next
  filt = txt
End Function

Function Slice2Str(ary, lbnd, ubnd, Optional joinChr As String = ",")
    If lbnd > ubnd Then Slice2Str = "ERROR": Exit Function
    Dim tmp()
    ReDim tmp(ubnd - lbnd)
    For i = 0 To UBound(tmp)
        tmp(i) = ary(lbnd + i)
    Next
    Slice2Str = Join(tmp, joinChr)
End Function

Public Sub scroll(t As TextBox)
  t.SelStart = Len(t)
End Sub

Public Sub AddtoLog(LogData As String)
    LogData = Replace(LogData, Chr(10), vbCrLf)
    If Len(frmMain.txtLog) + Len(LogData) > 40000 Then
        frmMain.txtLog = LogData
    Else
        frmMain.txtLog = frmMain.txtLog & LogData
    End If
    scroll frmMain.txtLog
End Sub

Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function SafeFileName(proposed) As String
  badChars = ">,<,&,/,\,:,|,?,*,"""
  bad = Split(badChars, ",")
  For i = 0 To UBound(bad)
    proposed = Replace(proposed, bad(i), "")
  Next
  SafeFileName = CStr(proposed)
End Function

'deals with servers that may terminate lines with cr, lf , or crlf
Function Standardize(it) As String()
    If it = "" Or it = Empty Or (InStr(it, Chr(10)) < 0 And InStr(it, Chr(13)) < 0) Then Exit Function
    
    Dim s() As String
    If InStr(1, it, Chr(10)) Then
      it = Replace(it, Chr(13), "")
      s() = Split(it, Chr(10))
    Else
      s() = Split(it, Chr(13))
    End If
    
    For i = 0 To UBound(s)
      s(i) = LTrim(Trim(s(i)))
    Next
    
    Standardize = s()
End Function

Function RemoveTotal(ByVal it As String) As String
    it = Trim(it)
    'vbcrlf = 0D 0A = chr(13) chr(10)
    If LCase(Left(it, 5)) = "total" Then
        ten = InStr(it, Chr(10))
        If ten > 0 Then
            it = Mid(it, ten + 1, Len(it))
        Else
            it = Mid(it, InStr(it, Chr(13)) + 1, Len(it))
        End If
     End If
     
     'these two crlf are a horrible hack :-\ ...
     'see the listview cant have index 0..so it has to start at
     'index 1 ....then I always want teh first element to be the
     'Upto parent directory list...so the first one we can use
     'is the 3rd array element ! worse off i just discovered this
     'two off error after like 3 months of use ! *hangs head in shame*
     RemoveTotal = vbCrLf & vbCrLf & it
End Function
