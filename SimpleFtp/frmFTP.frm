VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFTP 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transfer Progress..."
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3225
      TabIndex        =   0
      Top             =   60
      Width           =   930
   End
   Begin SimpleFtp.pBar pBar 
      Height          =   510
      Left            =   -60
      TabIndex        =   1
      Top             =   -45
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   900
   End
   Begin VB.Timer tmrKeepAlive 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   1470
      Top             =   870
   End
   Begin MSWinsockLib.Winsock wsData 
      Left            =   105
      Top             =   870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   990
      Top             =   870
   End
   Begin MSWinsockLib.Winsock wsControl 
      Left            =   555
      Top             =   870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totalBytes As Long
Dim curBytes As Long
Dim counter As Long

Private inFileData() As Byte

Private Sub cmdCancel_click()
    'On Error Resume Next
    wsData.Close 'still not right but works at least
    Close oFtp.openFileNum
    ftp.SendRawCommand "ABOR"
    Me.HideForm
End Sub

Private Sub wsControl_Close()
    oFtp.HadError = False
End Sub

Private Sub wsControl_DataArrival(ByVal bytesTotal As Long)
    Dim msg() As String
    Dim indata As String
    wsControl.GetData indata, vbString, bytesTotal
    
    AddtoLog indata
    msg() = Standardize(indata)
    
    For i = 0 To UBound(msg) - 1
        retcode = Left(msg(i), 3)
        oFtp.ServerResponse = Mid(msg(i), 4, Len(msg(i)))
        Select Case retcode
            Case 110, 202, 332, 421, 425, 426, 450, 451, _
                 452, 500, 501, 502, 503, 504, 530, 532, _
                 550, 551, 552, 553
                oFtp.HadError = True
                oFtp.WaitingToReturn = False
            Case 125, 150, 200, 211, 212, 213, 214, 215, 221, _
                 225, 226, 227, 257, 331, 350, 553, 220, 230, 250
                oFtp.HadError = False
                Select Case retcode
                    Case 225, 226: oFtp.DataServerClosed = True
                    Case 220, 230, 250: If Left(oFtp.ServerResponse, 1) <> "-" Then oFtp.WaitingToReturn = False
                    Case Else: oFtp.WaitingToReturn = False
                End Select
        End Select
    Next
End Sub

Private Sub wsData_Close()
    wsData.Close
    oFtp.DataConnected = False
End Sub

Private Sub wsData_Connect()
    oFtp.DataConnected = True
End Sub

Private Sub wsData_ConnectionRequest(ByVal requestID As Long)
    If oFtp.ConnectionMode = cPort Then
        wsData.Close
        wsData.Accept requestID
        oFtp.DataConnected = True
    End If
End Sub

Private Sub wsData_DataArrival(ByVal bytesTotal As Long)

    If oFtp.openFileNum > 0 And Not oFtp.ReceiveList Then
        ReDim inFileData(1 To bytesTotal) As Byte
        wsData.GetData inFileData(), , bytesTotal
        If oFtp.ResumingDownload Then
            Put #oFtp.openFileNum, LOF(oFtp.openFileNum) + 1, inFileData
        Else
            Put #oFtp.openFileNum, , inFileData
        End If
        HasRecievedAnother bytesTotal
    Else
        wsData.GetData tmp, vbString
        oFtp.ListData = oFtp.ListData & tmp
    End If

End Sub

Private Sub wsData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If oFtp.ConnectionMode = cPort Then
       oFtp.DataConnected = False
    Else 'PASV command connection to server failed
       oFtp.HadError = True
       oFtp.ServerResponse = Description
    End If
End Sub

Private Sub wsData_SendComplete()
    oFtp.DataSendComplete = True
End Sub

Private Sub tmrTimeout_Timer()
    oFtp.HadError = True
    oFtp.WaitingToReturn = False
End Sub

Private Sub tmrKeepAlive_Timer()
 On Error GoTo quit
    If Not oFtp.ReceivingFile And _
       Not oFtp.ReceiveList And _
       oFtp.openFileNum = 0 And oFtp.WaitingToReturn = False _
    Then
        If Not ftp.SendRawCommand("NOOP") Then err.Raise 40
    End If
 Exit Sub
quit: tmrKeepAlive.Enabled = False
End Sub


'####################################################
'                 Common functions
'####################################################

Function HasRecievedAnother(bytes)
    curBytes = curBytes + bytes
    counter = IIf(counter < 500, counter + 1, 0)
    If totalBytes > curBytes And counter Mod 20 = 0 Then
        pcent = Round(curBytes / totalBytes, 3) * 100
        sz = IIf(totalBytes > 100000, Round(totalBytes / 1000000, 2) & " Mb", totalBytes & " Bytes")
        pBar.SetPbPercent pcent, pcent & "% of " & sz
        Me.Caption = pcent & "% Transfered..."
        DoEvents
    End If
End Function

Public Sub init(FileSize)
    n = "frmProgress"
    Me.Left = GetSetting("FTPCLIENT", n, "MainLeft", 1000)
    Me.Top = GetSetting("FTPCLIENT", n, "MainTop", 1000)
    Me.Visible = True
    Me.Caption = "Starting Transfer..."
    
    pBar.ResetPb
    totalBytes = FileSize
    curBytes = 0
    counter = 0
End Sub

Public Sub HideForm()
    Me.Visible = False
    n = "frmProgress"
    SaveSetting "FTPCLIENT", n, "MainLeft", Me.Left
    SaveSetting "FTPCLIENT", n, "MainTop", Me.Top
End Sub


'for reference
'Public Enum RESPONSE_CODES
'    RESTATRT_MARKER_REPLY = 110
'    SERVICE_READY_IN_MINUTES = 120
'    DATA_CONNECTION_ALREADY_OPEN = 125
'    FILE_STATUS_OK = 150
'    COMMAND_OK = 200
'    COMMAND_NOT_IMPLEMENTED = 202
'    SYSTEM_HELP_REPLY = 211
'    DIRECTORY_STATUS = 212
'    FILE_STATUS = 213
'    HELP_MESSAGE = 214
'    NAME_SYSTEM_TYPE = 215
'    READY_FOR_NEW_USER = 220
'    CLOSING_CONTROL_CONNECTION = 221
'    DATA_CONNECTION_OPEN = 225
'    CLOSING_DATA_CONNECTION = 226
'    ENTERING_PASSIVE_MODE = 227
'    USER_LOGGED_IN = 230
'    FILE_ACTION_COMPLETED = 250
'    PATHNAME_CREATED = 257
'    USER_NAME_OK_NEED_PASSWORD = 331
'    NEED_ACCOUNT_FOR_LOGIN = 332
'    FILE_ACTION_PENDING_FURTHER_INFO = 350
'    SERVICE_NOT_AVAILABLE_CLOSING_CONTROL_CONNECTION = 421
'    CANNOT_OPEN_DATA_CONNECTION = 425
'    CONNECTION_CLOSED_TRANSFER_ABORTED = 426
'    FILE_ACTION_NOT_TAKEN = 450
'    ACTION_ABORTED = 451
'    ACTION_NOT_TAKEN = 452
'    COMMAND_UNRECOGNIZED = 500
'    ERROR_IN_PARAMETERS_OR_ARGUMENTS = 501
'    COMMAND_NOT_IMPLEMENTED = 502
'    BAD_SEQUENCE_OF_COMMANDS = 503
'    COMMAND_NOT_IMPLEMENTED_FOR_THAT_PARAMETER = 504
'    NOT_LOGGED_IN = 530
'    NEED_ACCOUNT_FOR_STORING_FILES = 532
'    ACTION_NOT_TAKEN_FILE_UNAVAILABLE = 550
'    ABORTED_PAGE_TYPE_UNKNOWN = 551
'    ABORTED_EXCEEDED_STORAGE_ALLOCATION = 552
'    FILE_NAME_NOT_ALLOWED = 553
'End Enum


