Attribute VB_Name = "fso"
'Info:     This is a wrapper for VB's built in file processes
'            making then easier to use.
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie

Public Type tFile
    baseName As String
    fullName As String
    fullpath As String
    filesize As Long
    extension As String
    parentFolder As String
    Attributes As VbFileAttribute
End Type
    
'should probably change these next to to use arrays would
'be more userfriendly cause collections really do suck
Public Type DirectoryContents
    sFolders As Collection
    sFiles As Collection
End Type

Public Function GetFileProps(filePath) As tFile
    Dim f As tFile
    tmp = Split(filePath, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       f.baseName = Mid(ub, 1, InStr(1, ub, ".") - 1)
       f.extension = Mid(ub, InStrRev(ub, "."), Len(ub))
    Else
       f.baseName = ub
       f.extension = ""
    End If
    f.fullName = ub
    f.fullpath = filePath
    If fso.FileExists(filePath) Then
        f.filesize = FileLen(filePath)
        f.Attributes = GetAttr(filePath)
    End If
    f.parentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
    GetFileProps = f
End Function

Public Function GetDirContents(path, Optional extension = ".*") As DirectoryContents
   Dim dc As DirectoryContents
   Set dc.sFiles = New Collection
   Set dc.sFolders = New Collection
   
   If Right(path, 1) <> "\" Then path = path & "\"
   If Left(extension, 1) = "*" Then extension = Mid(extension, 2, Len(extension))
   If Left(extension, 1) <> "." Then extension = "." & extension
   
   fs = Dir(path & "*" & extension, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then dc.sFiles.Add fs
     fs = Dir()
   Wend

   fd = Dir(path, vbDirectory)
   While fd <> ""
     If Left(fd, 1) <> "." Then
        If (GetAttr(path & fd) And vbDirectory) = vbDirectory Then
           dc.sFolders.Add fd
        End If
     End If
     fd = Dir()
   Wend
   
   GetDirContents = dc
End Function

Function GetSubFolders(folder, Optional retFullPath As Boolean = True) As String()
    Dim fnames() As String
    
    If Not FolderExists(folder) Then
        'returns empty array if fails
        GetSubFolders = fnames()
        Exit Function
    End If
    
   If Right(folder, 1) <> "\" Then folder = folder & "\"

   fd = Dir(folder, vbDirectory)
   While fd <> ""
     If Left(fd, 1) <> "." Then
        If (GetAttr(folder & fd) And vbDirectory) = vbDirectory Then
           push fnames(), IIf(retFullPath = True, folder & fd, fd)
        End If
     End If
     fd = Dir()
   Wend
   
   GetSubFolders = fnames()
End Function

Function FolderExists(ByVal path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function FileExists(ByVal path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Public Function GetParentFolder(path) As String
    Dim base As tFile
    base = GetFileProps(path)
    GetParentFolder = base.parentFolder
End Function

Public Function GetExtension(path) As String
    Dim base As tFile
    base = GetFileProps(path)
    GetExtension = base.extension
End Function

Public Function GetBaseName(path) As String
    Dim base As tFile
    base = GetFileProps(path)
    GetBaseName = base.baseName
End Function

Public Function GetFullName(path) As String
    Dim base As tFile
    base = GetFileProps(path)
    GetFullName = base.fullName
End Function

Public Sub CreateFolder(path)
   If FolderExists(path) Then Err.Raise 911, "CreateFolder", "Specified Folder Already Exists": Exit Sub
   MkDir path
End Sub

Function RandomNum()
    Randomize
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = tmp
End Function

Public Function ChangeExt(path, ext) As String
    ext = IIf(Left(ext, 1) = ".", ext, "." & ext)
    If fso.FileExists(path) Then
        Dim t As tFile
        t = fso.GetFileProps(path)
        ChangeExtension = t.parentFolder & "\" & t.baseName & ext
    Else
        'hack to just accept a file name might not be worth supporting
        bn = Mid(path, 1, InStr(1, path, ".") - 1)
        ChangeExtension = bn & ext
    End If
End Function

Public Function CreateTempFile(createIn, extension) As String
    
    If Not fso.FolderExists(createIn) Then Exit Function
    If Right(createIn, 1) <> "\" Then createIn = createIn & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
    Dim tmp As String
    Do
      tmp = createIn & RandomNum() & extension
    Loop Until Not fso.FileExists(tmp)
    
    f = FreeFile: Open tmp For Binary As f: Close f
    CreateTempFile = tmp
End Function

'------------------------------------------------------------
'--                     next two untested !                --
'------------------------------------------------------------

Function buildPath(folderpath) As Boolean
    On Error GoTo oops
    
    If FolderExists(folderpath) Then buildPath = True: Exit Function
    
    tmp = Split(folderpath, "\")
    build = tmp(0)
    For i = 1 To UBound(tmp)
        build = build & "\" & tmp(i)
        If InStr(tmp(i), ".") < 1 Then
            If Not FolderExists(build) Then CreateFolder (build)
        End If
    Next
    buildPath = True
    Exit Function
oops: buildPath = False
End Function

Function GetFolderFiles(folder, Optional filter = ".*") As String()
    Dim fnames() As String
    folder = IIf(Right(folder, 1) = "\", folder, folder & "\")
    If Not FolderExists(folder) Then Exit Function
    Dim dc As DirectoryContents
    dc = fso.GetDirContents(folder, filter)
    ReDim fnames(0 To dc.sFiles.count - 1)
    For i = 0 To dc.sFiles.count - 1
        fnames(i) = CStr(folder & dc.sFiles(i + 1))
    Next
    GetFolderFiles = fnames()
End Function

Function SafeFileName(proposed) As String
  badChars = ">,<,&,/,\,:,|,?,*,"""
  bad = Split(badChars, ",")
  For i = 0 To UBound(bad)
    proposed = Replace(proposed, bad(i), "")
  Next
  SafeFileName = CStr(proposed)
End Function

Public Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function WebFileNameFromPath(fullpath)
    If InStr(fullpath, "/") > 0 Then
        tmp = Split(fullpath, "/")
        WebFileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function readFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   readFile = temp
End Function

Public Sub writeFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Public Sub AppendFile(path, it)
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub


Sub Copy(fpath, toFolder)
   If FolderExists(toFolder) Then
       baseName = fso.FileNameFromPath(fpath)
       toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
       FileCopy fpath, toFolder & baseName
   Else 'assume tofolder is actually new desired file path
       FileCopy fpath, toFolder
   End If
End Sub

Sub Move(fpath, toFolder)
    Copy fpath, toFolder
    Kill fpath
End Sub

Sub Delete(fpath)
    Kill fpath
End Sub

Public Sub Rename(fullpath, newName)
  pf = fso.GetParentFolder(fullpath)
  Name fullpath As pf & "\" & newName
End Sub

Public Sub SetAttribute(fpath, it As VbFileAttribute)
   SetAttr fpath, it
End Sub

Public Sub CreateTextFile(fpath)
    f = FreeFile
    If fso.FileExists(fpath) Then Exit Sub
    Open fpath For Binary As f
    Close f
End Sub


'----------------------------------------------------------------------
'--                       Delete Folder Subs                         --
'----------------------------------------------------------------------
Public Sub DeleteFolder(folderpath, force As Boolean)
   Dim dc As DirectoryContents
   dc = fso.GetDirContents(folderpath)
   If dc.sFiles.count > 0 Or dc.sFolders.count > 0 And force = True Then Call deltre(CStr(folderpath), dc)
   Call RmDir(folderpath)
   path = Empty
End Sub


Private Function deltre(fp As String, fc As DirectoryContents)
'no error handling as safety (open files cause err)
   If fc.sFiles.count > 0 Then
      For i = 1 To fc.sFiles.count
        Kill fp & fc.sFiles.Item(i)
      Next
   End If
   
   If fc.sFolders.count > 0 Then
      For i = 1 To fc.sFolders.count
        Call redel(fp & "\" & fc.sFolders.Item(i))
      Next
   End If
End Function

Private Function redel(pt As String)
   Dim dd As DirectoryContents
   dd = fso.GetDirContents(pt)
   If dd.sFiles.count > 0 Or dd.sFolders.count > 0 Then Call deltre(pt, dd)
   Call RmDir(pt)
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Function aryIsEmpty(ary) As Boolean
  On Error GoTo oops
    X = UBound(ary)
    aryIsEmpty = False
  Exit Function
oops: aryIsEmpty = True
End Function

Function FolderName(folderpath) As String
    Dim ret As String
    tmp = Split(folderpath, "\")
    If Not aryIsEmpty(tmp) Then
        If Len(tmp(UBound(tmp))) <> 0 Then ret = tmp(UBound(tmp)) _
        Else ret = tmp(UBound(tmp) - 1)
    Else
        ret = CStr(folderpath)
    End If
    FolderName = ret
End Function

