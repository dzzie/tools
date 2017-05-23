VERSION 5.00
Begin VB.Form frmQuickView 
   Caption         =   "QuickView"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7935
   LinkTopic       =   "Form2"
   ScaleHeight     =   3915
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMsg 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3630
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   150
      Width           =   7710
   End
   Begin VB.Menu mnuQuick 
      Caption         =   "QuickView"
      Visible         =   0   'False
      Begin VB.Menu mnuSelALL 
         Caption         =   "Select ALL"
      End
      Begin VB.Menu mnuUnix2Dos 
         Caption         =   "Unix -> Dos"
      End
      Begin VB.Menu mnuDos2Unix 
         Caption         =   "Dos -> Unix"
      End
      Begin VB.Menu mnuUploadChange 
         Caption         =   "Upload Changes"
      End
   End
End
Attribute VB_Name = "frmQuickView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public RemoteFileName As String


Private Sub Form_Load()
    txtMsg.Font.Name = courier
    txtMsg.Font.Size = 12
    txtMsg.ForeColor = 16711680
    txtMsg.BackColor = &HEFEFEF
    n = "frmQuickView"
    Me.Left = GetSetting(App.Title, n, "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, n, "MainTop", 1000)
    Me.Width = GetSetting(App.Title, n, "MainWidth", 9000)
    Me.Height = GetSetting(App.Title, n, "MainHeight", 6500)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtMsg.Width = Me.Width - 400
    txtMsg.Height = Me.Height - txtMsg.Top - 550
End Sub

Private Sub Form_Unload(Cancel As Integer)
    n = "frmQuickView"
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, n, "MainLeft", Me.Left
        SaveSetting App.Title, n, "MainTop", Me.Top
        SaveSetting App.Title, n, "MainWidth", Me.Width
        SaveSetting App.Title, n, "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuSelALL_Click()
   txtMsg.SelStart = 0
   txtMsg.SelLength = Len(txtMsg)
End Sub

Private Sub mnuUploadChange_Click()
   If Not ftp.QuickUpload(RemoteFileName, txtMsg) Then MsgBox oFtp.ErrReason, vbCritical _
   Else Me.Hide
End Sub

Private Sub txtMsg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
        LockWindowUpdate txtMsg.hWnd
        txtMsg.Enabled = False
        DoEvents
        PopupMenu mnuQuick
        txtMsg.Enabled = True
        LockWindowUpdate 0&
  End If
End Sub

Private Sub mnuUnix2Dos_Click()
     txtMsg = UnixToDos(txtMsg.Text)
End Sub

Private Sub mnuDos2Unix_Click()
    txtMsg = Replace(txtMsg, vbCrLf, vbLf)
End Sub

Private Function UnixToDos(it) As String
    If InStr(it, vbLf) > 0 Then
        tmp = Split(it, vbLf)
        For i = 0 To UBound(tmp)
            If InStr(tmp(i), vbCr) < 1 Then tmp(i) = tmp(i) & vbCr
        Next
        UnixToDos = Join(tmp, vbLf)
    Else
        UnixToDos = CStr(it)
    End If
End Function
