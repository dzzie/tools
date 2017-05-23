VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Options"
   Begin VB.Frame Frame 
      Caption         =   "System Paths"
      Height          =   2550
      Index           =   5
      Left            =   5205
      TabIndex        =   45
      Top             =   4470
      Visible         =   0   'False
      Width           =   4470
      Begin VB.TextBox txtfolders 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   825
         OLEDropMode     =   1  'Manual
         TabIndex        =   57
         Text            =   "txtFolders"
         Top             =   2025
         Width           =   3420
      End
      Begin VB.TextBox txtfolders 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   825
         OLEDropMode     =   1  'Manual
         TabIndex        =   52
         Text            =   "txtFolders"
         Top             =   1665
         Width           =   3420
      End
      Begin VB.TextBox txtfolders 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   51
         Text            =   "txtFolders"
         Top             =   1320
         Width           =   3405
      End
      Begin VB.TextBox txtfolders 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   50
         Text            =   "txtFolders"
         Top             =   975
         Width           =   3450
      End
      Begin VB.TextBox txtfolders 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   49
         Text            =   "txtFolders"
         Top             =   630
         Width           =   3420
      End
      Begin VB.TextBox txtfolders 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   46
         Text            =   "txtFolders"
         Top             =   270
         Width           =   3435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SaveTo :"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   56
         Top             =   2085
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Editor :"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   55
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trash : "
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   54
         Top             =   1035
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Browser :"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   53
         Top             =   1380
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inbox :"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   48
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Outbox :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   690
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   345
      Index           =   4
      Left            =   5175
      TabIndex        =   44
      Top             =   510
      Width           =   765
   End
   Begin VB.Frame Frame 
      Caption         =   "X-Mailer Headers"
      Height          =   2430
      Index           =   4
      Left            =   5175
      TabIndex        =   34
      Top             =   4200
      Visible         =   0   'False
      Width           =   4470
      Begin VB.TextBox txtX 
         Height          =   1155
         Index           =   2
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   1110
         Width           =   3015
      End
      Begin VB.TextBox txtX 
         Height          =   315
         Index           =   1
         Left            =   2445
         TabIndex        =   39
         Top             =   360
         Width           =   1920
      End
      Begin VB.ListBox List2 
         Height          =   1620
         Left            =   180
         TabIndex        =   35
         Top             =   645
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "X-Content"
         Height          =   195
         Index           =   3
         Left            =   1335
         TabIndex        =   38
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "X-Description"
         Height          =   195
         Index           =   2
         Left            =   1350
         TabIndex        =   37
         Top             =   420
         Width           =   945
      End
      Begin VB.Label Label5 
         Caption         =   "X-Headers"
         Height          =   240
         Index           =   0
         Left            =   285
         TabIndex        =   36
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Custom Signature"
      Height          =   2565
      Index           =   3
      Left            =   5130
      TabIndex        =   31
      Top             =   3615
      Visible         =   0   'False
      Width           =   4605
      Begin VB.TextBox txtSig 
         Height          =   1830
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   555
         Width           =   4350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Signature"
         Height          =   195
         Left            =   105
         TabIndex        =   32
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Send Mail Config (Default)"
      Height          =   2430
      Index           =   1
      Left            =   5070
      TabIndex        =   15
      Top             =   3315
      Visible         =   0   'False
      Width           =   4395
      Begin VB.TextBox txtSendConfig 
         Height          =   330
         Index           =   2
         Left            =   1170
         TabIndex        =   21
         Top             =   1800
         Width           =   3105
      End
      Begin VB.TextBox txtSendConfig 
         Height          =   330
         Index           =   1
         Left            =   1185
         TabIndex        =   20
         Top             =   1110
         Width           =   3105
      End
      Begin VB.TextBox txtSendConfig 
         Height          =   330
         Index           =   0
         Left            =   1305
         TabIndex        =   19
         Top             =   540
         Width           =   2955
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Reply to:"
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   18
         Top             =   1875
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mail From :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   1215
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SMTP Server"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   585
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Create"
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   5190
      TabIndex        =   14
      Top             =   930
      Width           =   765
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Delete"
      Height          =   330
      Index           =   2
      Left            =   5205
      TabIndex        =   13
      Top             =   1335
      Width           =   750
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Cancel"
      Height          =   345
      Index           =   1
      Left            =   5220
      TabIndex        =   12
      Top             =   1740
      Width           =   750
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "OK"
      Height          =   330
      Index           =   0
      Left            =   5220
      TabIndex        =   11
      Top             =   2160
      Width           =   765
   End
   Begin VB.Frame Frame 
      Caption         =   "Account Manager"
      Height          =   2430
      Index           =   2
      Left            =   5025
      TabIndex        =   10
      Top             =   3045
      Visible         =   0   'False
      Width           =   4395
      Begin VB.TextBox txtAccount 
         Height          =   300
         Index           =   3
         Left            =   2385
         TabIndex        =   30
         Top             =   1800
         Width           =   1680
      End
      Begin VB.TextBox txtAccount 
         Height          =   300
         Index           =   2
         Left            =   2400
         TabIndex        =   29
         Top             =   1290
         Width           =   1695
      End
      Begin VB.TextBox txtAccount 
         Height          =   300
         Index           =   1
         Left            =   2400
         TabIndex        =   28
         Top             =   735
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   195
         TabIndex        =   23
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   5
         Left            =   1350
         TabIndex        =   27
         Top             =   1575
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "POP Server"
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   26
         Top             =   1875
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   3
         Left            =   1545
         TabIndex        =   25
         Top             =   1335
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   195
         Index           =   2
         Left            =   1485
         TabIndex        =   24
         Top             =   765
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Accounts"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   22
         Top             =   405
         Width           =   675
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "General Settings"
      Height          =   2550
      Index           =   0
      Left            =   375
      TabIndex        =   7
      Top             =   600
      Width           =   4470
      Begin VB.TextBox txtChar 
         Height          =   285
         Left            =   195
         MaxLength       =   3
         TabIndex        =   60
         Top             =   1065
         Width           =   270
      End
      Begin VB.TextBox txtfolders 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   1320
         OLEDropMode     =   1  'Manual
         TabIndex        =   42
         Text            =   "txtfolders6"
         Top             =   2130
         Width           =   3045
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Delete mail to Trash"
         Height          =   195
         Index           =   2
         Left            =   255
         TabIndex        =   41
         Top             =   540
         Width           =   1995
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Original Email in Reply"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   795
         Width           =   2385
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Save Copy of sent messages"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   2505
      End
      Begin VB.Label Label6 
         Caption         =   "Note: This UI can not edit all options. To create accounts            you must edit qmail.ini manually"
         Height          =   420
         Left            =   165
         TabIndex        =   61
         Top             =   1395
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Character to use in Reply "
         Height          =   195
         Index           =   9
         Left            =   570
         TabIndex        =   59
         Top             =   1155
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tip : Drag and Drop Folders and Files into Textboxes"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   58
         Top             =   1815
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Attach Folder :"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   2190
         Width           =   1035
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   3180
      Left            =   195
      TabIndex        =   6
      Top             =   180
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   5609
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Settings"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sending"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Accounts"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Signature"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "X-Mailer"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Folders"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   5
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   4
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   2
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum tabname
    Settings = 0
    sending = 1
    Accounts = 2
    Signature = 3
    xmailer = 4
End Enum

Dim seltab As tabname

'just drop folders and files in best way
Private Sub txtfolders_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   txtfolders(Index) = Data.Files(1)
End Sub

Private Sub cmdButton_Click(Index As Integer)
  Select Case Index
    Case 0: Call saveConfig
    Case 1: Unload Me 'cancel
    Case 2: MsgBox "Delete"
    Case 3: MsgBox "Create"
    Case 4: MsgBox "update"
  End Select
End Sub

Private Sub Form_Load()
  For i = 1 To TabStrip.Tabs.Count - 1
    Frame(i).Top = Frame(0).Top
    Frame(i).Left = Frame(0).Left
    Frame(i).Width = Frame(0).Width
    Frame(i).Height = Frame(0).Height
  Next
  Me.Width = 6120: Me.Height = 3750
  Call loadOptions
End Sub


Private Sub List2_Click()
  For i = 0 To List2.ListCount
    If List2.Selected(i) Then
       txtX(1) = uc.xHeaders.xDesc(i + 1)
       txtX(2) = uc.xHeaders.xContent(i + 1)
       Exit Sub
    End If
  Next
End Sub

Private Sub List1_Click()
  For i = 0 To List1.ListCount
    If List1.Selected(i) Then
       txtAccount(1) = uc.Users(i).user
       txtAccount(2) = uc.Users(i).pass
       txtAccount(3) = uc.Users(i).Server
       Exit Sub
    End If
  Next
End Sub

Private Sub TabStrip_Click()
   j = TabStrip.SelectedItem.Index - 1
   For i = 0 To TabStrip.Tabs.Count - 1
     Frame(i).Visible = False
   Next
   Frame(j).Visible = True
   
   If j = 2 Or j = 4 Then
      cmdButton(2).Enabled = True
      cmdButton(3).Enabled = True
      cmdButton(4).Enabled = True
   Else
      cmdButton(2).Enabled = False
      cmdButton(3).Enabled = False
      cmdButton(4).Enabled = False
   End If
   
   seltab = j
End Sub

Private Sub loadOptions()
    'tab 0
    Check1(0).Value = IIf(uc.Prefs.saveSent, 1, 0)
    Check1(1).Value = IIf(uc.Prefs.MsgInReply, 1, 0)
    Check1(2).Value = IIf(uc.Prefs.useTrash, 1, 0)
    txtChar = uc.Prefs.ReplyChar
    txtfolders(6) = uc.folders.attach
    txtfolders(7) = "Bitfile Functionality Expanded Edit Ini manually!"
    
    'tab 1
    txtSendConfig(0) = uc.Send.Server
    txtSendConfig(1) = uc.Send.sender
    txtSendConfig(2) = uc.Send.replyTo

    'tab 2
    For i = 0 To UBound(uc.Users) - 1
       List1.AddItem uc.Users(i).user
    Next
    
    'tab 3
    If fso.FileExists(uc.folders.sigFile) Then
       txtSig = fso.readFile(uc.folders.sigFile)
    End If

    'tab 4
    For i = 1 To uc.xHeaders.xDesc.Count
       List2.AddItem uc.xHeaders.xDesc(i)
    Next

    'tab 5
    txtfolders(0) = uc.folders.inbox
    txtfolders(1) = uc.folders.oubox
    txtfolders(2) = uc.folders.trash
    txtfolders(3) = Replace(uc.folders.browser, """", "")
    txtfolders(4) = Replace(uc.folders.editor, """", "")
    txtfolders(5) = uc.folders.saveTo
End Sub

Private Sub saveConfig()
    
    If Not VerifyFolders() Then Exit Sub
    
    'tab 0
    s = "preferences"
    Ini.SetValue s, "savesent", Check1(0).Value
    Ini.SetValue s, "MsgInReply", Check1(1).Value
    Ini.SetValue s, "usetrash", Check1(2).Value
    Ini.SetValue s, "ReplyChar", txtChar
    Ini.SetValue "folders", "attach", txtfolders(6)
    'Ini.SetValue "folders", "bitfile", txtfolders(7)
    
    'tab 1
    Ini.SetValue "sending", "server", txtSendConfig(0)
    Ini.SetValue "sending", "from", txtSendConfig(1)
    Ini.SetValue "sending", "replyto", txtSendConfig(2)

    'tab 2
    Ini.SetValue "profile", "number", List1.ListCount
    For i = 0 To List1.ListCount
       Ini.SetValue "profile", "user" & i + 1, uc.Users(i).user
       Ini.SetValue "profile", "pass" & i + 1, uc.Users(i).pass
       Ini.SetValue "profile", "server" & i + 1, uc.Users(i).Server
       Ini.SetValue "profile", "port" & i + 1, uc.Users(i).port
    Next
    
    'tab 3
    If txtSig <> "" Then
      fso.writeFile uc.folders.sigFile, txtSig
    End If

    'tab 4
    Ini.SetValue "X-Headers", "number", List2.ListCount
    For i = 1 To List2.ListCount
       Ini.SetValue "profile", "desc" & i, uc.xHeaders.xDesc(i)
       Ini.SetValue "profile", "content" & i, uc.xHeaders.xContent(i)
    Next

    'tab 5
    Ini.SetValue "folders", txtfolders(0), uc.folders.inbox
    Ini.SetValue "folders", txtfolders(1), uc.folders.oubox
    Ini.SetValue "folders", txtfolders(2), uc.folders.trash
    Ini.SetValue "folders", txtfolders(3), uc.folders.browser
    Ini.SetValue "folders", txtfolders(4), uc.folders.editor
    Ini.SetValue "folders", txtfolders(5), uc.folders.saveTo
    
    Ini.Save
    Startup.ClearConfig
    Startup.loadConfig
    Unload Me
End Sub

Public Function VerifyFolders() As Boolean
    Dim bad
    For i = 0 To txtfolders.Count - 1
       Dim t As String
       t = txtfolders(i)
        Select Case i
          Case 0, 1, 2, 5, 6
              If Not fso.FolderExists(t) Then
                 bad = bad & "Folder : " & t & vbCrLf
              Else
                txtfolders(i) = IIf(Right(t, 1) = "\", t, t & "\")
              End If
          Case Else
              If Not fso.FileExists(t) Then bad = bad & "File : " & t & vbCrLf
        End Select
    Next
    If bad <> Empty Then
       Msg = Replace("As a safety precaution you are going to have\nto create the following files or folders before there\npaths will be saved into the ini file\n\n", "\n", vbCrLf)
       MsgBox Msg & bad
       VerifyFolders = False
    Else
       VerifyFolders = True
    End If
End Function


