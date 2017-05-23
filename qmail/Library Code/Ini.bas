Attribute VB_Name = "Ini"
'Info:     Manipulate and query INI files, Has extra functionality over
'           windows API so that you can enumerate sections and keys
'           Example usage Ini.load : <manipulate file>
'                         Ini.Save : Ini.Release
'
'Bug Warning: this will create a second key with the same name as another
'             if you tell it to !
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie

Private Type Sect
  name As String
  key() As String
  Value() As String
End Type

Private IniObj() As Sect
Private IniFile As String

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
      ReDim IniObj(i).key(UBound(subs))
      ReDim IniObj(i).Value(UBound(subs))
      For j = 1 To UBound(subs)
        If Trim(subs(j)) <> "" Then
          a = Split(subs(j), "=")
          IniObj(i).key(j) = a(0)
          IniObj(i).Value(j) = a(1)
        End If
      Next
    Next
End Sub

Public Function GetValue(Section, key) As String
    On Error GoTo out
    s = GetSectionIndex(Section)
    k = GetKeyIndex(s, key)
    GetValue = IniObj(s).Value(k)
     Exit Function
out: GetValue = Empty
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
    ReDim r(UBound(IniObj(s).key))
    For j = 0 To UBound(IniObj(s).key)
       r(j) = IniObj(s).key(j)
    Next
    EnumKeys = r()
End Function


Public Function AddKey(Section, key, Value) As Boolean
    s = GetSectionIndex(Section)
    If s = -1 Then AddKey = False: Exit Function
    ub = UBound(IniObj(s).key) + 1
    ReDim Preserve IniObj(s).key(ub)
    ReDim Preserve IniObj(s).Value(ub)
    IniObj(s).key(ub) = key
    IniObj(s).Value(ub) = Value
    AddKey = True
End Function

Public Function AddSection(SectionName) As Boolean
        If GetSectionIndex(SectionName) <> -1 Then AddSection = False: Exit Function
        ub = UBound(IniObj) + 1
        ReDim Preserve IniObj(ub)
        IniObj(ub).name = SectionName
        ReDim IniObj(ub).key(0)
        ReDim IniObj(ub).Value(0)
        AddSection = True
End Function

Public Function DeleteSection(Section) As Boolean
    s = GetSectionIndex(Section)
    If s = -1 Then DeleteSection = False: Exit Function
    ReDim IniObj(s).key(0)
    ReDim IniObj(s).Value(0)
    IniObj(s).name = ""
    DeleteSection = True
End Function

Public Function DeleteKey(Section, key) As Boolean
    s = GetSectionIndex(Section)
    k = GetKeyIndex(s, key)
    If s = -1 Then DeleteKey = False: Exit Function
    If k = -1 Then DeleteKey = False: Exit Function
    IniObj(s).key(k) = ""
    IniObj(s).Value(k) = ""
    DeleteKey = True
End Function

Public Function SetValue(Section, key, newVal) As Boolean
    s = GetSectionIndex(Section)
    k = GetKeyIndex(s, key)
    If s = -1 Then SetValue = False: Exit Function
    If k = -1 Then SetValue = False: Exit Function
    IniObj(s).Value(k) = CStr(newVal)
    SetValue = True
End Function

Public Sub Save()
    For i = 0 To UBound(IniObj)
      If IniObj(i).name <> "" Then
        tmp = tmp & "[" & IniObj(i).name & "]" & vbCrLf
          For j = 0 To UBound(IniObj(i).key)
             If IniObj(i).key(j) <> "" Then
               tmp = tmp & IniObj(i).key(j) & "=" & IniObj(i).Value(j) & vbCrLf
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
    For i = 0 To UBound(IniObj(SectionIndex).key)
        If LCase(IniObj(SectionIndex).key(i)) = LCase(KeyName) Then
            GetKeyIndex = CInt(i)
            Exit Function
        End If
    Next
    GetKeyIndex = -1
End Function

Private Function readFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   readFile = temp
End Function

Private Sub writeFile(it As String)
    f = FreeFile
    Open IniFile For Output As #f
      Print #f, it
    Close f
End Sub











