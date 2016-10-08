Attribute VB_Name = "Module1"

Sub main()
    
    Dim cmd As String, ishex As Boolean
    
    cmd = Command
    
    If Right(cmd, 2) = "/x" Or Right(cmd, 2) = "-x" Then
        ishex = True
        cmd = Replace(cmd, "/x", Empty)
        cmd = Replace(cmd, "-x", Empty)
    End If
    
    cmd = Replace(cmd, """", Empty)
    cmd = Trim(cmd)
    
    If FileExists(cmd) Then
        tmp = FileLen(cmd)
        msg = "File"
    Else
        tmp = Len(Clipboard.GetText)
        msg = "Clipboard"
    End If
    
    If ishex Then tmp = "0x" & Hex(tmp)
    
    MsgBox "Length of data in " & msg & " is " & tmp & " bytes", vbInformation
    
End Sub

Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
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
