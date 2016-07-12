Attribute VB_Name = "Ini"
Private Type Sect
  name As String
  Key() As String
  Value() As String
End Type

Private IniObj() As Sect
Private IniFile As String
Private s As Integer

Public Sub LoadFile(INIFileLoc As String)
   
   If Dir(INIFileLoc) = "" Then Exit Sub
   IniFile = INIFileLoc
   inidata = readFile(INIFileLoc)
   
   tmp = Split(inidata, vbCrLf)
   For i = 0 To UBound(tmp)
      If Left(tmp(i), 1) = "[" Then tmp(i) = Replace(tmp(i), "[", Chr(5))
   Next
   
   inidata = Join(tmp, vbCrLf)
   sec = Split(inidata, Chr(5))
   ReDim IniObj(UBound(sec))
   
    For i = 1 To UBound(sec)
      IniObj(i).name = Mid(sec(i), 1, InStr(1, sec(i), "]") - 1)
      subs = Split(sec(i), vbCrLf)
      ReDim IniObj(i).Key(UBound(subs))
      ReDim IniObj(i).Value(UBound(subs))
      For j = 1 To UBound(subs)
        If Trim(subs(j)) <> "" Then
          a = Split(subs(j), "=")
          IniObj(i).Key(j) = a(0)
          IniObj(i).Value(j) = a(1)
        End If
      Next
    Next
End Sub

Public Function GetValue(Section, Key) As String
    s = GetSectionIndex(Section)
    k = GetKeyIndex(s, Key)
    GetValue = IniObj(s).Value(k)
End Function

Public Function EnumSections() As String()
    Dim r() As String
    ReDim r(UBound(IniObj))
    For i = 0 To UBound(IniObj)
        r(i) = IniObj(i).name
    Next
    EnumSections = r()
End Function

Public Function EnumKeys(Section) As String()
    Dim r() As String
    s = GetSectionIndex(Section)
    ReDim r(UBound(IniObj(s).Key))
    For j = 0 To UBound(IniObj(s).Key)
       r(j) = IniObj(s).Key(j)
    Next
    EnumKeys = r()
End Function


Public Function AddKey(Section, Key, Value) As Boolean
    s = GetSectionIndex(Section)
    If s = -1 Then AddKey = False: Exit Function
    ub = UBound(IniObj(s).Key) + 1
    ReDim Preserve IniObj(s).Key(ub)
    ReDim Preserve IniObj(s).Value(ub)
    IniObj(s).Key(ub) = Key
    IniObj(s).Value(ub) = Value
    AddKey = True
End Function

Public Function AddSection(SectionName) As Boolean
        If GetSectionIndex(SectionName) <> -1 Then AddSection = False: Exit Function
        ub = UBound(IniObj) + 1
        ReDim Preserve IniObj(ub)
        IniObj(ub).name = SectionName
        ReDim IniObj(ub).Key(0)
        ReDim IniObj(ub).Value(0)
        AddSection = True
End Function

Public Function DeleteSection(Section) As Boolean
    s = GetSectionIndex(Section)
    If s = -1 Then DeleteSection = False: Exit Function
    ReDim IniObj(s).Key(0)
    ReDim IniObj(s).Value(0)
    IniObj(s).name = ""
    DeleteSection = True
End Function

Public Function DeleteKey(Section, Key) As Boolean
    s = GetSectionIndex(Section)
    k = GetKeyIndex(s, Key)
    If s = -1 Then DeleteKey = False: Exit Function
    If k = -1 Then DeleteKey = False: Exit Function
    IniObj(s).Key(k) = ""
    IniObj(s).Value(k) = ""
    DeleteKey = True
End Function

Public Function SetValue(Section, Key, newVal) As Boolean
    s = GetSectionIndex(Section)
    k = GetKeyIndex(s, Key)
    If s = -1 Then SetValue = False: Exit Function
    If k = -1 Then SetValue = False: Exit Function
    IniObj(s).Value(k) = CStr(newVal)
    SetValue = True
End Function

Public Sub Save()
    For i = 0 To UBound(IniObj)
      If IniObj(i).name <> "" Then
        tmp = tmp & "[" & IniObj(i).name & "]" & vbCrLf
          For j = 0 To UBound(IniObj(i).Key)
             If IniObj(i).Key(j) <> "" Then
               tmp = tmp & IniObj(i).Key(j) & "=" & IniObj(i).Value(j) & vbCrLf
             End If
          Next
        tmp = tmp & vbCrLf
      End If
    Next
    Call writeFile(CStr(tmp))
    ReDim IniObj(0)
End Sub

Public Sub Release()
    ReDim IniObj(0) 'just to free up memory
End Sub

Private Function GetSectionIndex(Section) As Integer
    For i = 0 To UBound(IniObj)
       If LCase(IniObj(i).name) = LCase(Section) Then
          GetSectionIndex = CInt(i)
          Exit Function
       End If
    Next
    GetSectionIndex = -1
End Function

Private Function GetKeyIndex(SectionIndex, KeyName) As Integer
    For i = 0 To UBound(IniObj(SectionIndex).Key)
        If LCase(IniObj(SectionIndex).Key(i)) = LCase(KeyName) Then
            GetKeyIndex = CInt(i)
            Exit Function
        End If
    Next
    GetKeyIndex = -1
End Function

Private Function readFile(filename)
  f = FreeFile
  Temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     Temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   readFile = Temp
End Function

Private Sub writeFile(it As String)
    f = FreeFile
    Open IniFile For Output As #f
      Print #f, it
    Close f
End Sub











