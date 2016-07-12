Attribute VB_Name = "Globals"
Public Type HTTPRequest
     ip As String
     page As String
     rawData As String
     method As String
     arg() As String   'all querystring args in key=value format
     qryStr As String
     uAgent As String
     BasicAuth As String
End Type

Public Type Script
  path As String
  data As Variant
  RespCode As Integer
  extension As String
  httpHeader As String
  AuthedHeader As String   'the redir headerif it was a 401
  serveData As Boolean     'false if we are just redirecting
  AuthURL As String        'Accept only this URL
  probeServer As String    'Server to hand off Probe Requests :)
End Type

Public Type ip
  userAgent As String
  reqs As Integer          'number of Requests
  authed As Variant        'Authorization string provided default=false
  ip As String
  time As String           'last request from them
  PROBED As Variant        'last probed url req
End Type

Public Type config
    LogFile As String
    DynDns As String
    LastPath As String
End Type

Global cfg As config
Global Script As Script       'overused object of diff attributes
Global IPs() As ip            'access the data of the connectors
Global s As New Strings
Global ReadyToClose As Boolean


Function ParseRequest(X, ip) As HTTPRequest
    Dim h As HTTPRequest
    h.rawData = X
    s.Strng = Trim(X)
    
    h.ip = ip
    h.method = s.SubstringToChar(1, " ")
    
    fsp = s.IndexOf(" ") + 2         'first space
    ssp = s.NextIndexOf              'second space
    h.page = s.Substring(fsp, ssp)   'page request
    
    start = s.IndexOf("Basic:", , True)
    If start > 0 Then h.BasicAuth = s.ToEndOfLine
    
    start = s.IndexOf("User-Agent:", , True)
    If start > 0 Then h.uAgent = s.ToEndOfLine
    
    s.Strng = h.page
    qs = s.IndexOf("?")
       
    If qs > 0 Then
        h.page = s.SubstringToChar(1, "?")
        h.qryStr = s.ToEndOfStr(qs + 1)
        h.arg() = Split(h.qryStr, "&")
    End If
    
    ParseRequest = h
End Function

Sub resetScript()       'setScript Object to defaults
  With Script           'note serveData default is false
    .path = Empty
    .httpHeader = Empty
    .AuthedHeader = Empty
    .extension = Empty
    .serveData = True
    .AuthURL = Empty
    .probeServer = Empty
    .data = Empty
  End With
End Sub

Function buildHeader(ResponseCode, Connection, ContentType, Optional AUTH = False, Optional Location = False) As String
'Connection Types: 'Keep-Alive','Chunked','Close'
'ContentType: 'text/html','image/jpeg','image/gif'
'for redirect you must use ResponseCode of 301, or '302 Found'
'404 Not Found, 403 FORBIDDEN , 410 Gone, 401 Authorization Required

Script.RespCode = ResponseCode  'global object set

header = "HTTP/1.1 "
Select Case ResponseCode
   Case 200
     header = header & "200 OK" & vbCrLf
   Case 301
     header = header & "301 Moved Permanently" & vbCrLf
   Case 401
     header = header & "401 Authorization Required" & vbCrLf
End Select

header = header & "Server: Apache/1.3.11 (Unix)" & vbCrLf
header = header & "Pragma: no-cache" & vbCrLf
header = header & "Accept-Ranges: bytes" & vbCrLf
header = header & "Content-Length: " & FileLen(Script.path) & vbCrLf

If AUTH <> False Then
 header = header & "WWW-Authenticate: Basic realm=""" & AUTH & """" & vbCrLf
End If

header = header & "Connection: " & Connection & vbCrLf
header = header & "Content-Type: " & ContentType & vbCrLf

If Location <> False Then
  header = header & "Location: " & Location & vbCrLf
End If

header = header & vbCrLf 'signifies end of httpHeader
buildHeader = header     'return the value
End Function

Sub loadConfig()
    Call Ini.LoadFile(App.path & "\config.ini")
    frmMain.chkAuth = IIf(Ini.GetValue("main", "chkauth"), 1, 0)
    frmMain.chkLog = IIf(Ini.GetValue("main", "chkLog"), 1, 0)
    cfg.LogFile = Ini.GetValue("main", "LogFile")
    cfg.DynDns = Trim(Ini.GetValue("main", "DynDns"))
    cfg.LastPath = Ini.GetValue("main", "LastPath")
    frmMain.txtConfig(0) = Ini.GetValue("main", "Redirect")
End Sub

Sub saveConfig()
    Call Ini.LoadFile(App.path & "\config.ini")
    Ini.SetValue "main", "chkAuth", frmMain.chkAuth
    Ini.SetValue "main", "chkLog", frmMain.chkLog
    Ini.SetValue "main", "LogFile", cfg.LogFile
    Ini.SetValue "main", "DynDns", cfg.DynDns
    Ini.SetValue "main", "Redirect", frmMain.txtConfig(0)
    Ini.SetValue "main", "LastPath", getLastPath(frmMain.txtConfig(1))
    Ini.Save
End Sub

Function Escape(it)
    Dim f(): Dim c()
    n = Replace(it, "+", " ")
    If InStr(n, "%") > 0 Then
        t = Split(n, "%")
        For i = 0 To UBound(t)
            a = Left(t(i), 2)
            b = IsHex(a)
            If b <> Empty Then
                push f(), "%" & a
                push c(), b
            End If
        Next
        For i = 0 To UBound(f)
            n = Replace(n, f(i), c(i))
        Next
    End If
    Escape = n
End Function

Private Function IsHex(it)
    On Error GoTo out
      IsHex = Chr(Int("&H" & it))
    Exit Function
out:  IsHex = Empty
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function getLastPath(it)
    s.Strng = it
    If s.CountOccurancesOf("/") > 3 Then
        X = Split(it, "/")
        For i = 3 To UBound(X)
            tmp = tmp & X(i) & IIf(i < UBound(X), "/", "")
        Next
        getLastPath = tmp
    End If
End Function








