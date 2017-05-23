Attribute VB_Name = "General"
'These are all common library functions applicaple to any project
'Library functions are specific to this email program

Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
    
Sub ShowRtClkMenu(f As Form, t As TextBox, m As Menu)
        LockWindowUpdate t.hWnd
        t.Enabled = False
        DoEvents
        f.PopupMenu m
        t.Enabled = True
        LockWindowUpdate 0&
End Sub

Function AryIndexExists(ary, index) As Boolean
    On Error GoTo nope
    X = ary(index)
    AryIndexExists = True
    Exit Function
nope: AryIndexExists = False
End Function

Function BatchReplace(ByRef it, them, Optional compare As VbCompareMethod = vbTextCompare) As String
    t = Split(them, ",")
    For i = 0 To UBound(t)
        If InStr(t(i), "-") > 1 Then
            s = Split(t(i), "-")
            it = Replace(it, s(0), s(1), , , compare)
        End If
    Next
    BatchReplace = CStr(it)
End Function

Function skinny(t, base) 'remove empty elements
    Dim ret()                   'return adjustable base array
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


Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    ReDim Preserve ary(UBound(ary) + 1) '<-throws Error If Not initalized
    ary(UBound(ary)) = Value
    Exit Sub
init: ReDim ary(0): ary(0) = Value
End Sub

Function AryIsEmpty(ary) As Boolean
 On Error GoTo out
    X = UBound(ary) '<- throws error if uninitalized
    AryIsEmpty = False
 Exit Function
out: AryIsEmpty = True
End Function

Function IsHex(it)
    On Error GoTo out
      IsHex = Chr(Int("&H" & it))
    Exit Function
out:  IsHex = Empty
End Function

Function endSlash(it)
   ret = IIf(Right(it, 1) = "\", it, it & "\")
   endSlash = ret
End Function

Function aryize(it As String, base As Integer) As String()
    Dim r() As String
    
    For i = base To Len(it)
        ReDim Preserve r(i)
        r(i) = Mid(it, i, 1)
    Next
    aryize = r()

End Function

Function StrFindValFromKey(ary() As String, key)
 On Error GoTo out
    For i = 0 To UBound(ary)
        pos = InStr(ary(i), "=")
        If pos > 0 Then
            k = Mid(ary(i), 1, pos - 1)
            v = Mid(ary(i), pos + 1, Len(ary(i)))
            If LCase(k) = LCase(key) Then
                StrFindValFromKey = v
                Exit Function
            End If
        End If
    Next
out:
    StrFindValFromKey = Empty
End Function

Function br(it) As String
    t = Replace(it, "\n", vbCrLf)
    t = Replace(t, "\t", vbTab)
    br = Replace(t, "\q", """")
End Function

Function FirstLine(it) As String
    If InStr(it, vbCrLf) > 0 Then FirstLine = Mid(it, 1, InStr(it, vbCrLf)) _
    Else FirstLine = CStr(it)
End Function

'remove all html tags (can be buggered if html
'tag contains quoted > or <
Function parseHtml(info) As String
     Dim Temp As String, EndOfTag As Integer
     fmat = Replace(info, "&nbsp;", " ")
     cut = Split(fmat, "<")

   For i = 0 To UBound(cut)  'cut at all html start tags
     EndOfTag = InStr(1, cut(i), ">")
        If EndOfTag > 0 Then
          EndOfText = Len(cut(i))
          NL = False
          If Left(cut(i), 2) = "br" Then NL = True
          cut(i) = Mid(cut(i), EndOfTag + 1, EndOfText)
          If NL Then cut(i) = vbCrLf & cut(i)
          If cut(i) = vbCrLf Then cut(i) = ""
        End If
     Temp = Temp & cut(i)
    Next
    
    parseHtml = Temp
End Function

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


Function SpellCheck(strText As String) As String
    'This function opens the MS Word Object and uses its spell checker
    'passing back the corrected string
    
    If strText = Empty Then Exit Function
    
    On Error GoTo out
    Dim oWDBasic As Object, sTmpString As String

    Set oWDBasic = CreateObject("Word.Basic")
    Screen.MousePointer = vbHourglass
   
    With oWDBasic
        .FileNew
        .Insert strText
        .ToolsSpelling oWDBasic.EditSelectAll
        .SetDocumentVar "MyVar", oWDBasic.Selection
    End With
    
    sTmpString = oWDBasic.GetDocumentVar("MyVar")
    sTmpString = Left(sTmpString, Len(sTmpString) - 1)

    If sTmpString = "" Then SpellCheck = strText _
    Else SpellCheck = Replace(sTmpString, Chr$(13), vbCrLf, , , 1)
        
    oWDBasic.FileCloseAll 2
    oWDBasic.AppClose
    Set oWDBasic = Nothing
    Screen.MousePointer = vbNormal
    MsgBox "Spell check is completed.", vbInformation
    Exit Function
out:
    MsgBox "For this feature to work you have to have MS word Installed on this computer sorry :-\", vbInformation
    SpellCheck = strText
End Function

Function IsIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Function ReadFirst(path, xBytes) As String
    Dim tmp() As Byte
    Dim ret() As String
    f = FreeFile
    ReDim tmp(1 To xBytes)
    
    Open path For Binary As f
    Get f, , tmp()
    Close f
    
    ReDim ret(1 To UBound(tmp))
    For i = 1 To UBound(tmp)
        ret(i) = Chr(tmp(i))
    Next
    ReadFirst = Replace(Join(ret, ""), Chr(0), Empty)
End Function
