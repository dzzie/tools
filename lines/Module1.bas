Attribute VB_Name = "Module1"

Private Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long


Sub Main()
    
    Dim cmd As String, ishex As Boolean, tmp
    Dim fs As New CFileStream
    Dim sData As String, sMessage As String
    
    cmd = Command
    
    If Right(cmd, 2) = "/x" Or Right(cmd, 2) = "-x" Then
        ishex = True
        cmd = Replace(cmd, "/x", Empty)
        cmd = Replace(cmd, "-x", Empty)
    End If
    
    cmd = Replace(cmd, """", Empty)
    cmd = Trim(cmd)
    
    Con.Initialize
   
    If Con.Piped Then
      sData = Con.ReadStream()
      tmp = countOccurances(sData, vbLf)
      msg = "Piped Input"
    ElseIf FileExists(cmd) Then
        tmp = 0
        fs.Open_ cmd 'better for large file support
        While Not fs.eof
            fs.ReadLine
            tmp = tmp + 1
        Wend
        fs.Close_
        'tmp = countOccurances(ReadFile(cmd), vbLf)
        msg = "File"
    Else
        tmp = countOccurances(Clipboard.GetText, vbLf)
        msg = "Clipboard"
    End If
    
    If ishex Then tmp = "0x" & Hex(tmp)
    
    msg = msg & " contains " & tmp & " lines of text"
    
    If Len(sData) > 0 Then
        Con.WriteLine msg 'must switch PE header from GUI to console
    Else
        MsgBox msg, vbInformation
    End If
    
End Sub

Function countOccurances(x, find) As Long
    On Error Resume Next
    If InStr(x, find) < 1 Then Exit Function
    y = Split(x, find)
    countOccurances = UBound(y)
End Function

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
