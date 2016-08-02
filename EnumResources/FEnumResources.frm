VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FEnumResources 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FEnumResources"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.TreeView TreeView1 
      Height          =   4275
      Left            =   3180
      TabIndex        =   3
      Top             =   180
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7541
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   4140
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   180
      TabIndex        =   1
      Top             =   2760
      Width           =   2835
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2835
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnumResources.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnumResources.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnumResources.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnumResources.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FEnumResources.frx":0C68
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FEnumResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©2000 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Implements IEnumResources

Private m_UserFile As String

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
   With File1
      If Right(.Path, 1) = "\" Then
         m_UserFile = .Path & .FileName
      Else
         m_UserFile = .Path & "\" & .FileName
      End If
   End With
   Call UpdateInfo(m_UserFile)
End Sub

Private Sub File1_PathChange()
   If File1.ListCount Then
      File1.ListIndex = 0
   Else
      m_UserFile = ""
      Call UpdateInfo(m_UserFile)
   End If
End Sub

Private Sub Form_Load()
   ' Just look at DLL/EXE files.
   File1.Pattern = "*.exe;*.dll;*.ocx"
   ' Set initial dirspec
   Drive1.Drive = Environ("windir")
   Dir1.Path = Environ("windir")
   ' Clean-up UI
   Set Me.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call UnloadChildren
End Sub

Private Sub TreeView1_Click()
   ' Easy way to look at node keys.
   Debug.Print TreeView1.SelectedItem.key
End Sub

Private Sub TreeView1_DblClick()
   Dim frm As FShowResource
   Dim ResType As Long
   Dim key As String
   
   key = TreeView1.SelectedItem.key
   If InStr(key, "#") = 1 Then
      ResType = Val(Mid(key, 2))
   Else
      ' Unknown resource type
      Exit Sub
   End If
   
   Select Case ResType
      Case RT_BITMAP
         Set frm = New FShowResource
         Load frm
         frm.DisplayBitmap m_UserFile, Mid$(key, InStr(key, "\") + 1)
      Case RT_ICON, RT_GROUP_ICON
         Set frm = New FShowResource
         frm.DisplayIcon m_UserFile, Mid$(key, InStr(key, "\") + 1)
      Case RT_STRING
         Set frm = New FShowResource
         frm.DisplayString m_UserFile, Mid$(key, InStr(key, "\") + 1)
   End Select
End Sub

' ***************************************************
'  Private methods
' ***************************************************
Private Sub UnloadChildren()
   Dim frm As Form
   For Each frm In Forms
      If Not (frm Is Me) Then
         Unload frm
      End If
   Next frm
End Sub

Private Sub UpdateInfo(ByVal FileSpec As String)
   Dim tvn As Node
   
   Call UnloadChildren
   
   If Len(FileSpec) Then
      Debug.Print FileSpec
      ' Call EnumResources(FileSpec)
      With TreeView1
         .Visible = False
         .Nodes.Clear
         Set tvn = .Nodes.Add(, , "Root", FileSpec, 4)
         tvn.Expanded = True
         If EnumResourcesEx(Me, FileSpec) = False Then
            Set tvn = .Nodes.Add("Root", tvwChild, "nil", "Couldn't load this module", 5)
         End If
         .Visible = True
      End With
      Me.Caption = "Resources: " & FileSpec
   Else
      ' No file passed, clear display info.
      TreeView1.Nodes.Clear
      Me.Caption = "EnumResources Demo"
   End If
End Sub

' ***************************************************
'  Implemented interface methods
' ***************************************************
Private Sub IEnumResources_EnumResourceSink(ByVal hModule As Long, ByVal ResName As String, ByVal ResType As String, Continue As Boolean)
   Dim tvn As Node
   Dim DispName As String
   
   With TreeView1.Nodes
      If Len(ResName) Then
         ' Add resource to proper type, disallowing dups.
         On Error Resume Next
         Set tvn = .Add(ResType, tvwChild, ResType & "\" & ResName, ResName, 3)
      Else
         ' New resource type
         If InStr(ResType, "#") = 1 Then
            DispName = ResTypeName(Val(Mid(ResType, 2)))
         Else
            DispName = Chr$(34) & ResType & Chr$(34)
         End If
         Set tvn = .Add("Root", tvwChild, ResType, DispName, 1)
         tvn.ExpandedImage = 2
         tvn.Expanded = True
      End If
   End With
End Sub

