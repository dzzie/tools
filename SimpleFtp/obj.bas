Attribute VB_Name = "obj"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Enum TransferModes
    cBinary = 0
    cASCII = 1
End Enum

Public Enum conModes
    cPort = 0
    cPasv = 1
End Enum

'##############################################
' Main Ftp object
Public Type ftpManagement
    'flags
    DataConnected As Boolean
    DataSendComplete As Boolean
    DataServerClosed As Boolean
    ResumingDownload As Boolean
    WaitingToReturn As Boolean
    ReceivingFile As Boolean
    ReceiveList As Boolean
    HadError As Boolean
    'data
    openFileNum As Integer
    ServerResponse As String
    ErrReason As String
    ListData As String
    'configs
    ConnectionMode As conModes
    timeOut As Long
    useNOOP As Boolean
End Type
'##############################################

Public Type ftpConnectionStats
    folder As String
    server As String
    port As Integer
    user As String
    pass As String
End Type

Public Enum ftpFileTypes 'these numbers corrospond to the file
    movie = 1            'icons held in an image list
    zip = 2
    fldr = 3
    exe = 4
    txt = 5
    img = 6
    unknown = 7
    'updir = 8
End Enum

Public Type ftpFile
    permissions As String
    ftype As ftpFileTypes
    fname As String
    byteSize As Long
End Type


Global oFtp As ftpManagement

Global f() As ftpFile 'global array of the stats on the
                      'current remote folders contents
Global dlDir As String         'download directory


'##############################################################
'                       Object Parsing
'##############################################################
Function parseFtpString(X) As ftpConnectionStats
    Dim c As ftpConnectionStats
    X = Replace(X, "ftp://", "")
    at = InStrRev(X, "@")
    If at > 0 Then 'is user:pass@server:port type login
       authinfo = Mid(X, 1, at - 1)
       serverinfo = Mid(X, at + 1)
       tmp = Split(authinfo, ":")
       c.user = tmp(0)
       c.pass = tmp(1)
       If InStr(serverinfo, ":") > 0 Then
           tmp = Split(serverinfo, ":")
           c.server = tmp(0)
           c.port = tmp(1)
       Else
           c.server = serverinfo
           c.port = 21
       End If
    Else 'is anonymous login type url
       c.user = "anonymous"
       c.pass = "someone@somewhere.com"
       slash = InStr(X, "/")
       sc = InStr(X, ":")
       If sc > 0 Then
          c.server = Mid(X, 1, sc - 1)
          If slash > 0 Then
                c.port = Mid(X, sc + 1, (slash - 1) - sc)
                c.folder = Mid(X, slash, Len(X))
          Else
                c.port = Mid(X, sc + 1, Len(X))
          End If
       Else
          c.port = 21
          If slash > 0 Then
                c.server = Mid(X, 1, slash - 1)
                c.folder = Mid(X, slash, Len(X))
          Else
                c.server = X
          End If
       End If
    End If
    
    If c.folder = "/" Then c.folder = Empty
    If Right(c.folder, 1) = "/" Then c.folder = Mid(c.folder, 1, Len(c.folder) - 1)
    
    'MsgBox "Port: " & c.port & vbCrLf & _
    '        "Server: " & c.server & vbCrLf & _
    '        "User: " & c.user & vbCrLf & _
    '        "Pass: " & c.pass & vbCrLf & _
    '        "Folder:" & c.folder
    
    parseFtpString = c
End Function

Function parseFtpFile(them() As String) As ftpFile()
  'some servers like wuftp send 1st element as "total xx" this
  'was removed in the remove total after the string was first
  'received format of data to parse:
  '-rw-r--r--   1 tilk     users          78 Jan 16 05:15 plus.gif
  
  'MsgBox Join(them, vbCrLf) & UBound(them)
  
  Dim f() As ftpFile
  ReDim f(UBound(them) - 1)
  
  f(1).fname = ".." 'will always be up dir
  For i = 2 To UBound(them) - 1
        On Error GoTo SKIP2_NEXT
        them(i) = shave(them(i))
        'MsgBox them(i)
        tmp = skinny(Split(them(i), " "))
        If UBound(tmp) < 4 Then GoTo SKIP2_NEXT
        
        If InStr(them(i), ":") > 0 Then
            'endsection = from seconds in timestamp to end of filename
            'have to do it this way to handle filenames with spaces
            endSection = Mid(them(i), InStrRev(them(i), ":"), Len(them(i)))
            f(i).fname = Trim(Mid(endSection, InStr(endSection, " ") + 1, Len(endSection)))
        Else
           f(i).fname = tmp(UBound(tmp))
        End If
        dot = InStrRev(f(i).fname, ".")
        
        f(i).permissions = tmp(0)
        If LCase(Left(tmp(0), 1)) = "d" Then f(i).ftype = fldr
        
        If f(i).ftype <> fldr Then
            If dot Then
                Dim t As ftpFileTypes
                ext = Mid(f(i).fname, dot, Len(f(i).fname))
                Select Case LCase(ext)
                  Case ".exe": t = exe
                  Case ".zip", ".tgz", ".z", ".gz", ".tar", ".rar": t = zip
                  Case ".txt", ".htm", ".html": t = txt
                  Case ".avi", ".mpg", ".asf", ".asx", ".wav", _
                       ".mp3", ".ram", ".mov", ".mpeg": t = movie
                  Case ".jpg", ".jpeg", ".gif", ".bmp": t = img
                  Case Else: t = unknown
                End Select
                f(i).ftype = t
             Else
               f(i).ftype = unknown
             End If
        End If
         
         f(i).byteSize = tmp(4)
SKIP2_NEXT:
     Next
    'MsgBox UBound(f)
    parseFtpFile = f
End Function












