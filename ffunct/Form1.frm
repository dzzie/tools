VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Folder Functions"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   105
      TabIndex        =   5
      Top             =   0
      Width           =   6975
      Begin VB.CheckBox Check1 
         Caption         =   "Recursive"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   9
         ToolTipText     =   "DO NOT USE THIS UNLES SYOU KNOW WHAT IT MEANS!!!"
         Top             =   630
         Width           =   1380
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   5220
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   1575
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Text            =   "Drop Folder to Process In here"
         Top             =   630
         Width           =   5235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Function :"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   105
         TabIndex        =   8
         Top             =   210
         Width           =   945
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   7200
      TabIndex        =   4
      Top             =   45
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   105
      TabIndex        =   0
      Top             =   1050
      Width           =   6930
      Begin VB.CommandButton Command1 
         Caption         =   "Process"
         Height          =   300
         Left            =   6000
         TabIndex        =   3
         Top             =   180
         Width           =   750
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1095
         TabIndex        =   2
         Top             =   180
         Width           =   4725
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         Caption         =   "lblMsg"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   225
         Width           =   720
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim shortcut() As String

Private Sub Combo1_Click()
    preprocess Combo1.ListIndex
End Sub

Private Sub Form_Load()
  
  vals = "Clear,->,html,txt,doc,exe,dll,ocx"
  shortcut = Split(vals, ",")
  For i = 0 To UBound(shortcut)
    List2.AddItem shortcut(i), i
  Next
  
  vals = "Extension Rename,Merge Files,Directory Listing,Html Index," & _
       "Remove Scripting,Remove HTML,Unzip,Zip,Set Hidden,Unset Hidden," & _
       "Set ReadOnly,Unset ReadOnly,Img Src Index,Sequential Rename," & _
       "Unix -> Dos,Dos -> Unix,Batch Replace,Prefix FileName,Strip Filename"
       
  s = Split(vals, ",")
  For i = 0 To UBound(s)
    Combo1.AddItem s(i), i
  Next
  
  Combo1.ListIndex = 0
  preprocess 2
  Text2 = GetSetting("Filefunc", "x", "lst", Text2)
End Sub


Private Sub preprocess(Index As Integer)
  Selopt = Index
  Select Case Index
     Case 0: lblMsg = "Rename ": Text2 = "html;htm->txt"
     Case 1: lblMsg = "Join ": Text2 = "txt;log"
     Case 2, 3: lblMsg = "Exclude ": Text2 = "dll;ocx;exe;ps"
     Case 4, 5: lblMsg = "Parse ": Text2 = "html;htm;xml;txt"
     Case 6: lblMsg = "Unzip ": Text2 = "zip": Call activate(Text2, False)
     Case 7: lblMsg = "Exclude ": Text2 = "log;tmp"
     Case 8, 9, 10, 11: lblMsg = "Exclude ": Text2 = "exe;dll;ocx"
     Case 12, 13, 17: lblMsg = "Include ": Text2 = "jpg;gif;bmp"
  End Select
  If Index <> 6 Then Call activate(Text2)
End Sub

Private Sub Command1_Click()
 'On Error GoTo warn
   SaveSetting "Filefunc", "x", "lst", Text2
   If Text1 = Empty Then MsgBox "You must drag and drop the files or folders you want to work with in the large white box :)": Exit Sub
   Select Case Selopt
      Case 6, 7, 13, 17, 16, 18: Call Prompt
   End Select
   Call Vdate(Selopt) 'reads in extension options
   
   Screen.MousePointer = 11
   Call ListEngine(Text1.ToolTipText)
   Screen.MousePointer = 0
   Me.Caption = "File Functions - complete: " & Now
Exit Sub
warn: MsgBox Err.Description
      Screen.MousePointer = 0
End Sub


Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Text1.ToolTipText = Data.files(1)
 Text1 = FolderName(Data.files(1))
End Sub

Private Sub Text2_Change()
  Text2.SelLength = Len(Text2.Text)
  Text2.ToolTipText = Text2
End Sub
Private Sub lblMsg_Click()
   If InStr(1, lblMsg.Caption, "Exclude") Then lblMsg.Caption = "Include " _
   Else If InStr(1, lblMsg.Caption, "Include") Then lblMsg.Caption = "Exclude "
End Sub
Private Sub lblMsg_Change()
    lblMsg.Left = Text2.Left - lblMsg.Width - 40
End Sub

Private Sub List2_Click() 'semi intelligent ruleset creator
 If Text2.Enabled = False Then Exit Sub
  seld = 0
  ar = InStr(1, Text2, ">")
  c = ";"
  For i = 0 To List2.ListCount
    If List2.Selected(i) Then seld = i: Exit For
  Next
  If seld = 0 Then Text2 = "": Exit Sub
  If seld = 1 And ar > 0 Then Exit Sub
  If Text2 = "" Or seld = 1 Or ar > 0 Then c = ""
  If ar > 0 And Len(Text2) - ar > 2 Then Text2 = Mid(Text2, 1, ar)
  Text2 = Text2 & c & shortcut(seld)
End Sub

Private Sub activate(txt As TextBox, Optional enable = True)
  If enable Then
    txt.Enabled = True
    txt.BackColor = vbWhite
  Else
    txt.Enabled = False
    txt.BackColor = &H8000000F
  End If
End Sub

Private Sub Prompt()
  
  Select Case Selopt
    Case 6: msg = "Unzip All Files Into Parent Directory?"
    Case 7:  msg = "Zip Entire Folder into one Zip?"
    Case 13: msg = "If you would like a character appended..enter it below."
    Case 17: msg = "Enter string to prepend to filename."
    Case 16: msg = "replace 'this'->'that' no quotes"
    Case 18: msg = "Enter string to strip from filename."
  End Select
  
  Select Case Selopt
    Case 6, 7
        ans = MsgBox(msg, vbYesNo)
        If ans = vbYes Then Together = True _
          Else: Together = False
    Case 13, 17, 18
        Seq.count = 0
        Seq.Appnd = InputBox(msg)
    Case 16
        Seq.Appnd = InputBox(msg)
  End Select
   
End Sub
