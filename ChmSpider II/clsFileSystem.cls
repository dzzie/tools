VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Revision 3 <- Incompatiable with all previous..simplified & streamlined
'
'Info:     These are basically macros for VB's built in file processes
'            this should streamline your code quite a bit and hopefully
'            remove alot of redundant coding.
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie
    
Option Explicit

Function GetFolderFiles(folderpath As String, Optional filter As String = "*", Optional retFullPath As Boolean = True) As String()
   Dim fnames() As String
   Dim fs As String
   
   If Not FolderExists(folderpath) Then
        'returns empty array if fails
        GetFolderFiles = fnames()
        Exit Function
   End If
   
   folderpath = IIf(Right(folderpath, 1) = "\", folderpath, folderpath & "\")
   
   fs = Dir(folderpath & filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then push fnames(), IIf(retFullPath = True, folderpath & fs, fs)
     fs = Dir()
   Wend
   
   GetFolderFiles = fnames()
End Function

Function GetSubFolders(folderpath As String, Optional retFullPath As Boolean = True) As String()
    Dim fnames() As String
    Dim fd As String
    
    If Not FolderExists(folderpath) Then
        'returns empty array if fails
        GetSubFolders = fnames()
        Exit Function
    End If
    
   If Right(folderpath, 1) <> "\" Then folderpath = folderpath & "\"

   fd = Dir(folderpath, vbDirectory)
   While fd <> ""
     If Left(fd, 1) <> "." Then
        If (GetAttr(folderpath & fd) And vbDirectory) = vbDirectory Then
           push fnames(), IIf(retFullPath = True, folderpath & fd, fd)
        End If
     End If
     fd = Dir()
   Wend
   
   GetSubFolders = fnames()
End Function

Function FolderExists(path As String) As Boolean
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Function FileExists(path As String) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function GetParentFolder(path) As String
    Dim tmp() As String
    Dim ub As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

Sub CreateFolder(path As String)
   If FolderExists(path) Then Exit Sub
   MkDir path
End Sub

Function FileNameFromPath(fullpath As String) As String
    Dim tmp() As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    Else
        FileNameFromPath = fullpath
    End If
End Function

Function WebFileNameFromPath(fullpath As String)
    Dim tmp() As String
    If InStr(fullpath, "/") > 0 Then
        tmp = Split(fullpath, "/")
        WebFileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function DeleteFile(fpath As String) As Boolean
 On Error GoTo hadErr
    Kill fpath
    DeleteFile = True
 Exit Function
hadErr:
'MsgBox "DeleteFile Failed" & vbCrLf & vbCrLf & fpath
DeleteFile = False
End Function

Sub Rename(fullpath As String, newName As String)
  Dim pf As String
  pf = GetParentFolder(fullpath)
  Name fullpath As pf & "\" & newName
End Sub

Sub SetAttribute(fpath, it As VbFileAttribute)
   SetAttr fpath, it
End Sub

Function GetExtension(path) As String
    Dim tmp() As String
    Dim ub As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetExtension = Mid(ub, InStrRev(ub, "."), Len(ub))
    Else
       GetExtension = ""
    End If
End Function

Function GetBaseName(path As String) As String
    Dim tmp() As String
    Dim ub As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Function ChangeExt(path As String, ext As String)
    Dim bn As String
    ext = IIf(Left(ext, 1) = ".", ext, "." & ext)
    If FileExists(path) Then
        Rename path, GetBaseName(path) & ext
    Else
        'hack to just accept a file name might not be worth supporting
        bn = Mid(path, 1, InStr(1, path, ".") - 1)
        ChangeExt = bn & ext
    End If
End Function

Function SafeFileName(proposed As String) As String
  Dim badChars As String, bad() As String, i As Long
  badChars = ">,<,&,/,\,:,|,?,*,"""
  bad = Split(badChars, ",")
  For i = 0 To UBound(bad)
    proposed = Replace(proposed, bad(i), "")
  Next
  SafeFileName = CStr(proposed)
End Function

Function RandomNum() As Long
    Dim tmp As Long
    Randomize
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = tmp
End Function

Function GetFreeFileName(folder As String, Optional extension = ".txt") As String
    
    If Not FolderExists(folder) Then Exit Function
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
    Dim tmp As String
    Do
      tmp = folder & RandomNum() & extension
    Loop Until Not FileExists(tmp)
    
    GetFreeFileName = tmp
End Function

Function buildPath(folderpath As String) As Boolean
    On Error GoTo oops
    
    If FolderExists(folderpath) Then buildPath = True: Exit Function
    
    Dim tmp() As String, build As String, i As Long
    
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


Function ReadFile(filename) As Variant
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub writeFile(path As String, it As Variant)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Sub AppendFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub


Function Copy(fpath As String, toFolder As String)
   Dim baseName As String, newName As String
   If FolderExists(toFolder) Then
       baseName = FileNameFromPath(fpath)
       toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
       newName = toFolder & baseName
       FileCopy fpath, newName
       Copy = newName
   Else 'assume tofolder is actually new desired file path
       FileCopy fpath, toFolder
       Copy = toFolder
   End If
End Function

Function Move(fpath As String, toFolder As String)
    Dim fname As String
    fname = FileNameFromPath(fpath)
    toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
    Copy fpath, toFolder
    Kill fpath
    Move = toFolder & fname
End Function

Sub CreateFile(fpath As String)
    Dim f As Long
    f = FreeFile
    If FileExists(fpath) Then Exit Sub
    Open fpath For Binary As f
    Close f
End Sub


Function DeleteFolder(folderpath As String, Optional force As Boolean = True) As Boolean
 On Error GoTo failed
   Call delTree(folderpath, force)
   Call RmDir(folderpath)
   DeleteFolder = True
 Exit Function
failed:  DeleteFolder = False
End Function

Private Sub delTree(folderpath As String, Optional force As Boolean = True)
   Dim sfi() As String, sfo() As String, i As Long
   sfi() = GetFolderFiles(folderpath)
   sfo() = GetSubFolders(folderpath)
   If Not aryIsEmpty(sfi) And force = True Then
        For i = 0 To UBound(sfi)
            Kill sfi(i)
        Next
   End If
   
   If Not aryIsEmpty(sfo) And force = True Then
        For i = 0 To UBound(sfo)
            Call DeleteFolder(sfo(i), True)
        Next
   End If
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Private Function aryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim X As Long
    X = UBound(ary)
    aryIsEmpty = False
  Exit Function
oops: aryIsEmpty = True
End Function

Function FolderName(folderpath) As String
    Dim ret As String, tmp() As String
    tmp = Split(folderpath, "\")
    If Not aryIsEmpty(tmp) Then
        If Len(tmp(UBound(tmp))) <> 0 Then ret = tmp(UBound(tmp)) _
        Else ret = tmp(UBound(tmp) - 1)
    Else
        ret = CStr(folderpath)
    End If
    FolderName = ret
End Function


