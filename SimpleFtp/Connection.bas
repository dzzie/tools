Attribute VB_Name = "ftp"
'connection.bas & frmFTP CodeBase created by NeoText as DLL
'       WebSight: http://www.neotextsoftware.com
'       v1 DLL Size: 57kb
'       Original Source Freely Available at:
'             http://www.pscode.com/xq/ASP/txtCodeId.11609/lngWId.1/qx/vb/scripts/ShowCode.htm
'
'Modified April 2001 - approx time 70hrs.
'       Modifications: dzzie@yahoo.com
'       Websight: http://www.geocities.com/dzzie/
'
'       Changelog:
'           Added in method of data transfer that could store data in memory
'               from data connection and not have to implicitly write to disk
'               this is used for obtaining directory listings and QuickView
'
'           Added support for: append file, resume download, send raw command
'                 keep alive timer, Quickview File and QuickUpload File

'           Added FtpFile Obj and parsing to pass back dir list obj array
'           Added FtpConnectionStats Obj and parsing for connection strings
'           Added progress bar user control in form
'           Re-coded all parsing of messages received from server
'           Changed method of err handling throughout
'           Programmed in default binary mode for transfers
'           General code cleanup/ optimization/ bug-fixes
'
'           Restructured code to compile in to main exe, ftp log
'               data now can show realtime and not after function return
'
'       Notes:
'               Neotext had an excellent flow layout of events for sending
'           commands and waiting for responses, much cleaner than alot
'           i have seen. I usally code everythign from scratch, but in
'           this case, his work was so close to how i would have tried to
'           do it, i couldnt justify the approximated 60-70hrs to recode
'           the winsock engine part from scratch. I have made substantial
'           changes, but his work still lies at its core, all credit for
'           design and layout is his.
'               My work here falls into the catagory of housecleaning,
'           and upgrading. Such is the power of open source we can all build
'           upon each others backs each doing what we can before our work
'           exceeds our projects needs. Then someone else can pick it up and
'           keep it going. I just hope that is what happens to my work anyway :)


Property Get ConnectionMode() As conModes
    ConnectionMode = oFtp.ConnectionMode
End Property

Property Let ConnectionMode(mode As conModes)
    oFtp.ConnectionMode = mode
End Property


Property Get timeOut() As Long
    timeOut = oFtp.timeOut
End Property

Property Let timeOut(NewVal As Long)
    oFtp.timeOut = NewVal
End Property

Property Get ErrReason() As String
    ErrReason = oFtp.ErrReason
End Property

Private Sub Initialize(Optional err As String = "")
    oFtp.ResumingDownload = False
    oFtp.DataSendComplete = False
    oFtp.DataServerClosed = False
    oFtp.WaitingToReturn = False
    oFtp.DataConnected = False
    oFtp.ReceiveList = False
    oFtp.HadError = False
    oFtp.openFileNum = 0
    oFtp.ErrReason = err
End Sub

'##############################################################
'          Ftp Command Wrappers & Err Handeling
'##############################################################
'  all subs will be set to just return boolean
'  simple subs will repress the error throwing and just check
'  if boolean, complex subs will let it raise an error, but will
'  catch the error internally and then themselves return a boolean
'  to the calling function and then set teh ftpErrReason variable
'  with err.description which is what the server responded with
'---------------------------------------------------------------
Private Function SendCommand(ftpCommand, Optional raiseErr As Boolean = True) As Boolean
   On Error GoTo out
    oFtp.WaitingToReturn = True
    frmFTP.wsControl.SendData ftpCommand & vbCrLf
    AddtoLog ftpCommand & vbCrLf
    frmFTP.tmrTimeout.Interval = ftpTimeOut
    frmFTP.tmrTimeout.Enabled = True
    
    Do Until oFtp.WaitingToReturn = False
        DoEvents: DoEvents: DoEvents
    Loop

    frmFTP.tmrTimeout.Enabled = False
    'MsgBox "done with command " & ftpCommand
    If checkforError(raiseErr) = True Then SendCommand = False _
    Else SendCommand = True
Exit Function
out: If raiseErr Then err.Raise 9, "SendCommand", "SendCommand <-" & err.Description _
     Else SendCommand = False
End Function

Private Function checkforError(Optional raiseErr As Boolean = True) As Boolean
    If oFtp.HadError = True Then
        checkforError = True
        oFtp.ErrReason = oFtp.ServerResponse
        If raiseErr Then err.Raise "80", "FtpClient", oFtp.ServerResponse
    Else
        checkforError = False
    End If
End Function

Private Function SendDataCommand(ftpCommand)
  On Error GoTo out
    oFtp.DataConnected = False
    
   With frmFTP.wsData
     Select Case oFtp.ConnectionMode
        Case cPort
            .Close
            .LocalPort = 0
            .Listen
            SendCommand "PORT " & Replace(.LocalIP, ".", ",") & "," & .LocalPort \ 256 & "," & .LocalPort Mod 256
            SendCommand ftpCommand
            WaitForDataOpen
        Case Else
            SendCommand "PASV"
            'need to fork this parsing off to a sub that can
            'wait until proper resp code = 227 is received
            sResp = oFtp.ServerResponse
            lPar = InStrRev(sResp, "(") + 1
            rPar = InStrRev(sResp, ")")
            info = Mid(sResp, lPar, rPar - lPar)
            tmp = Split(info, ",")
            pServer = Slice2Str(tmp, 0, 3, ".")
            pPort = (CLng(tmp(4)) * 256) + CLng(tmp(5))
            
            .Close
            .LocalPort = 0
            .RemotePort = pPort
            .RemoteHost = pServer
            .Connect
            
            WaitForDataOpen
            SendCommand ftpCommand
     End Select
  End With
Exit Function
out: err.Raise 80, , "Caught in SendDataCommand " & err.Description
End Function


'##############################################################
'                      PUBLIC MEMBERS
'##############################################################

Function Connect(ftpConString As String) As Boolean
 On Error GoTo out
    Call Initialize
    Dim f As ftpConnectionStats
    f = parseFtpString(ftpConString)
    If f.server = Empty Or f.port = Empty Then ftpErrReason = "Not Valid Connection String": Exit Function
    
    With frmFTP.wsControl
        .Close
        .RemoteHost = f.server
        .LocalPort = 0
        .RemotePort = f.port
        .Connect
    End With
    
    WaitForTCP
        
    SendCommand "USER " & f.user
    SendCommand "PASS " & f.pass
    
    If f.folder <> Empty Then ChangeDirectory f.folder
    If oFtp.useNOOP Then frmFTP.tmrKeepAlive.Enabled = True
    Connect = True

Exit Function
out:  ftpErrReason = err.Description
      frmFTP.tmrKeepAlive.Enabled = False
End Function

Function Disconnect()
    frmFTP.tmrKeepAlive.Enabled = False
    SendCommand "QUIT"
End Function

Function RenameFile(SourceFileName, DestFileName) As Boolean
    If SendCommand("RNFR " & SourceFileName, False) Then
       RenameFile = SendCommand("RNTO " & DestFileName, False)
    Else
       RenameFile = False
    End If
End Function

Function ChangeDirectory(ToFolder) As Boolean
        ChangeDirectory = SendCommand("CWD " & ToFolder, False)
        Call frmMain.ChangedCurrentFolder(ToFolder)
End Function

Function MakeDirectory(NewFolder) As Boolean
    MakeDirectory = SendCommand("MKD " & NewFolder, False)
End Function

Function RemoveFile(FileName) As Boolean
    RemoveFile = SendCommand("DELE " & FileName, False)
End Function

Function RemoveDirectory(FolderName) As Boolean
    RemoveDirectory = SendCommand("RMD " & FolderName, False)
End Function

Function SendRawCommand(rawCommand) As Boolean
    SendRawCommand = SendCommand(rawCommand, False)
End Function

Function TransferType(TransType As TransferModes)
    If TransType = cBinary Then SendCommand "TYPE I" _
    Else SendCommand "TYPE A"
End Function

Function QuickView(RemoteFileName, byteSize) As String
On Error GoTo out
    oFtp.ListData = Empty
    oFtp.ReceiveList = True
    If RemoteFileName = "" Then Exit Function
    TransferType cASCII
    Call frmFTP.init(byteSize)
    SendDataCommand "RETR " & RemoteFileName
    WaitForDataClose
    WaitForDataServerClosed
    frmFTP.HideForm
    QuickView = oFtp.ListData

Exit Function
out: oFtp.ReceiveList = False: oFtp.ErrReason = err.Description
End Function

Function QuickUpload(SaveAs, data As String) As Boolean
 On Error GoTo out
    If data = Empty Or SaveAs = Empty Then oFtp.ErrReason = "Empty Argument in function!": Exit Function
        
    frmFTP.init Len(data)
    Call TransferType(cASCII)
    SendDataCommand "STOR " & SaveAs
    oFtp.openFileNum = 1
    
    frmFTP.wsData.SendData data
    WaitForDataSend
        
    frmFTP.wsData.Close
    WaitForDataServerClosed
    
    frmFTP.pBar.SetPbPercent 99
    frmFTP.HideForm
    oFtp.openFileNum = 0
    QuickUpload = True
    
Exit Function
out: oFtp.ErrReason = err.Description: oFtp.openFileNum = 0
End Function

Function ListContents() As ftpFile() 'returns base 2 array !
 On Error GoTo out                       '(for adding to img list)
 
    oFtp.ListData = Empty
    oFtp.ReceiveList = True
    SendDataCommand "LIST"

    WaitForDataClose
    WaitForDataServerClosed

    oFtp.ReceiveList = False
    Dim s() As String
    s() = Standardize(RemoveTotal(oFtp.ListData))
    
    ListContents = parseFtpFile(s)

Exit Function
out: oFtp.ReceiveList = False
End Function

Function GetFile(RemoteFileName, DownloadDirectory, byteSize, Optional binarymode As Boolean = True, Optional raiseErr As Boolean = True) As Boolean
 On Error GoTo out
    If RemoteFileName = "" Then Exit Function
           
    fpath = endSlash(DownloadDirectory) & SafeFileName(RemoteFileName)
    
    If FileExists(fpath) Then
        msg1 = "Would you like to Resume the download of " & RemoteFileName & " ?"
        msg2 = "Confirm Delete of " & RemoteFileName
        If byteSize <= FileLen(fpath) Then
           MsgBox "You have already downloaded a file with this name & size..exiting": Exit Function
        ElseIf MsgBox(msg1, vbYesNo) = vbYes Then
            oFtp.ResumingDownload = True
            remainder = byteSize - FileLen(fpath)
            frmFTP.init remainder
            frmFTP.HasRecievedAnother remainder
        ElseIf MsgBox(msg2, vbExclamation + vbYesNo) = vbYes Then
            Kill fpath
        Else
            Exit Function
        End If
    End If
    
    Dim FileNum As Integer
    FileNum = FreeFile
    oFtp.ReceivingFile = True
    
    If binarymode Then Call TransferType(cBinary)
    
    If Not oFtp.ResumingDownload Then
        Call frmFTP.init(byteSize)
        Open fpath For Output As #FileNum
        Close #FileNum
    End If
    
    Open fpath For Binary Access Write As #FileNum
    oFtp.openFileNum = FileNum
    
    If oFtp.ResumingDownload Then
        SendDataCommand "REST " & FileLen(fpath)
    End If
    
    SendDataCommand "RETR " & RemoteFileName
    
    WaitForDataClose
    frmFTP.pBar.SetPbPercent 99
    WaitForDataServerClosed
    Close #FileNum
    
    oFtp.openFileNum = 0
    oFtp.ReceivingFile = False
    frmFTP.HideForm
    GetFile = True

Exit Function
out: Close #FileNum: oFtp.ErrReason = err.Description
     oFtp.ResumingDownload = False
     If raiseErr Then err.Raise err.Number, , err.Description
End Function

Function PutFile(LocalFileName, Optional binarymode As Boolean = True, Optional raiseErr As Boolean = True) As Boolean
 On Error GoTo out
    Const PacketSize = 3072
    Dim inFileData() As Byte
    Dim cnt As Long
    Dim bytesCount As Long
    Dim bytesLeft As Long

    tmp = Split(LocalFileName, "\")
    SaveAs = tmp(UBound(tmp))
    
    If binarymode Then Call TransferType(cBinary)
    SendDataCommand "STOR " & SaveAs
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Open LocalFileName For Binary Access Read As #FileNum
    oFtp.openFileNum = FileNum
    
    byteSize = LOF(FileNum)
    bytesCount = byteSize / PacketSize
    bytesLeft = byteSize Mod PacketSize
    
    Call frmFTP.init(byteSize)
    ReDim inFileData(1 To PacketSize) As Byte
    
    If bytesCount > 0 Then
        For cnt = 1 To bytesCount
            Get #FileNum, , inFileData()
            frmFTP.wsData.SendData inFileData()
            frmFTP.HasRecievedAnother PacketSize
            WaitForDataSend
        Next
    End If
    
    If bytesLeft > 0 Then
        ReDim inFileData(1 To bytesLeft) As Byte
        Get #FileNum, , inFileData()
        frmFTP.wsData.SendData inFileData()
        WaitForDataSend
    End If
    
    frmFTP.pBar.SetPbPercent 99
    frmFTP.wsData.Close
    WaitForDataServerClosed
    
    Close #FileNum
    oFtp.openFileNum = 0
    frmFTP.HideForm
    PutFile = True
    
Exit Function
out: Close #FileNum: oFtp.ErrReason = err.Description
If raiseErr Then err.Raise err.Number, , err.Description
End Function

Function AppendFile(RemoteFileName, LocalFileName, startAt As Long, Optional binarymode As Boolean = True, Optional raiseErr As Boolean = True) As Boolean
  On Error GoTo out
    Const PacketSize = 3072
    Dim inFileData() As Byte
    Dim advancePointer As Byte
    Dim cnt As Long
    Dim bytesCount As Long
    Dim bytesLeft As Long
    
    If binarymode Then Call TransferType(cBinary)
    SendDataCommand "APPE " & RemoteFileName
    
    Dim FileNum As Integer
    FileNum = FreeFile
    oFtp.openFileNum = FileNum
    
    Open LocalFileName For Binary Access Read As #FileNum
    
    byteSize = LOF(FileNum) - startAt
    bytesCount = byteSize / PacketSize
    bytesLeft = byteSize Mod PacketSize
    
    Call frmFTP.init(byteSize)
    ReDim inFileData(1 To PacketSize) As Byte
    
    If bytesCount > 0 Then
        'advance file pointer to right position
        Get #FileNum, startAt - 1, advancePointer
        For cnt = 1 To bytesCount
            Get #FileNum, , inFileData()
            frmFTP.wsData.SendData inFileData()
            frmFTP.HasRecievedAnother PacketSize
            WaitForDataSend
        Next
    End If
    
    If bytesLeft > 0 Then
        If bytesCount <= 0 Then
            'advance file pointer to right position
            Get #FileNum, startAt - 1, advancePointer
        End If
        ReDim inFileData(1 To bytesLeft) As Byte
        Get #FileNum, , inFileData()
        frmFTP.wsData.SendData inFileData()
        WaitForDataSend
    End If
    
    frmFTP.wsData.Close
    WaitForDataServerClosed
    
    Close #FileNum
    oFtp.openFileNum = 0
    frmFTP.HideForm
    AppendFile = True
    
Exit Function
out: oFtp.ErrReason = err.Description: Close #FileNum
If raiseErr Then err.Raise err.Number, , err.Description
End Function

'##############################################################
'                      Time Wait Subs
'##############################################################

Private Sub WaitForTCP()
    oFtp.WaitingToReturn = True
    frmFTP.tmrTimeout.Interval = oFtp.timeOut
    frmFTP.tmrTimeout.Enabled = True
    
    Do Until oFtp.WaitingToReturn = False
        DoEvents
    Loop
    
    frmFTP.tmrTimeout.Enabled = False
    checkforError
End Sub

Private Sub WaitForDataOpen()
    Do Until oFtp.DataConnected = True
        DoEvents
    Loop
    checkforError
End Sub

Private Sub WaitForDataClose()
    Do Until oFtp.DataConnected = False
        DoEvents
    Loop
    checkforError
End Sub

Private Sub WaitForDataSend()
    Do Until oFtp.DataSendComplete = True
        DoEvents
    Loop
    oFtp.DataSendComplete = False
    checkforError
End Sub

Private Sub WaitForDataServerClosed()
    Do Until oFtp.DataServerClosed = True
        DoEvents
    Loop
    oFtp.DataServerClosed = False
    checkforError
End Sub

