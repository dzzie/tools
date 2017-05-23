VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Chm Spider II  - http://sandSprite.com"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   8760
   Icon            =   "CHM-SPIDER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   60
      TabIndex        =   1
      Top             =   2640
      Width           =   8655
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   15
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox txtBaseFolder 
         Height          =   330
         Left            =   1680
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   300
         Width           =   6180
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7440
         TabIndex        =   8
         Top             =   1140
         Width           =   1110
      End
      Begin VB.TextBox txtExtensions 
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
         Left            =   1680
         TabIndex        =   7
         ToolTipText     =   "Space delimited list of extensions to accept (wild cards permitted)"
         Top             =   1140
         Width           =   3975
      End
      Begin VB.CheckBox chkRecursive 
         Caption         =   "Recursive"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   6
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.TextBox txtOutPutFileName 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cboDefaultPage 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   4035
      End
      Begin VB.CommandButton cmdGenerateFile 
         Caption         =   "Generate Files"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   5760
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Base Folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Extension Filter:  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   13
         ToolTipText     =   "Click me to toggle action"
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Output File Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   12
         ToolTipText     =   "Click me to toggle action"
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Default Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Click me to toggle action"
         Top             =   1620
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Window Title"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   780
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHM-SPIDER.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHM-SPIDER.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHM-SPIDER.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CHM-SPIDER.frx":0C58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   4366
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   35
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu mnuFolderPopup 
      Caption         =   "mnuFolderPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRemFolder 
         Caption         =   "Remove Folder"
      End
   End
   Begin VB.Menu mnuFilePopup 
      Caption         =   "mnuFilePopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRemFile 
         Caption         =   "Remove File"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'file headers for chm script output
Dim hhpHeader As String
Dim hhcHeader As String

'filpaths for the different script files
Dim hhp As String
Dim hhc As String

'some formatting stuff needed to output scripts
Dim fldrNode As String
Dim fileNode As String

Const DEBUGMODE = False

Private curNode As Node

Sub d(Msg)
    If DEBUGMODE Then Debug.Print Msg
End Sub

Private Sub cmdBrowse_Click()
    Dim c As New clsCmnDlg
    txtBaseFolder = c.FolderDialog(, Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveMySetting "ext", txtExtensions.text
    RemoveSubClass
    On Error Resume Next
    Unload frmMain
End Sub

Sub Form_Load()
  Dim tmp() As String
  On Error GoTo hell
  
  clsTvm.SetTreeviewReference tv
  
  'txtBaseFolder = "C:\Documents and Settings\Administrator\Desktop\debug"
  
  SetMinWidthToCurrent Me, , 500
  SubClassLimitFormResize Me.hwnd
  
  txtExtensions = GetMySetting("ext", ".htm* .txt .doc .pdf")
  
  push tmp(), "<LI> <OBJECT type=\qtext/sitemap\q>"
  push tmp(), "\t <param name=\qName\q value=\q_____\q>"
  push tmp(), "\t </OBJECT>"
  push tmp(), " \t <UL>" & vbCrLf
  fldrNode = br(Join(tmp, vbCrLf))
  Erase tmp()
  
  push tmp(), "\t <LI> <OBJECT type=\qtext/sitemap\q>"
  push tmp(), "\t \t <param name=\qName\q value=\q*****\q>"
  push tmp(), "\t \t <param name=\qLocal\q value=\q_____\q>"
  push tmp(), " \t \t </OBJECT>" & vbCrLf
  fileNode = br(Join(tmp, vbCrLf))
  Erase tmp()
  
  push tmp(), "[Options]"
  push tmp(), "Compatibility = 1.1 Or later"
  push tmp(), "Compiled file = <filename>.chm"
  push tmp(), "Contents file= TOC.hhc"
  push tmp(), "Default topic = <deftopic>"
  push tmp(), "Display compile progress=Yes"
  push tmp(), "Enhanced decompilation=Yes"
  push tmp(), "Language=0x409 English (United States)"
  push tmp(), "Title=<title>"
  push tmp(), ""
  push tmp(), "[Files]"
  push tmp(), ""
  hhpHeader = Join(tmp, vbCrLf)
  Erase tmp()
  
  push tmp(), "<!DOCTYPE HTML PUBLIC \q-//IETF//DTD HTML//EN\q>"
  push tmp(), "<HTML><HEAD>"
  push tmp(), "<meta name=\qGENERATOR\q content=\qMicrosoft&reg; HTML Help Workshop 4.1\q>"
  push tmp(), "<!-- Sitemap 1.0 -->"
  push tmp(), "</HEAD><BODY>"
  push tmp(), "<UL>"
  push tmp(), ""
  hhcHeader = br(Join(tmp, vbCrLf))
  
Exit Sub
hell:
     
    MsgBox Err.Description, vbExclamation
    End
    
End Sub

Private Sub cmdGenerateFile_Click()

 Dim n As Node
 Dim fname As String
 Dim fpath As String
 Dim tmp As String
 'Dim nested As long
 'Dim fDeep As long
 
 Dim dirStack() As String
 Dim curfolderParent As String
 
    If tv.Nodes.Count < 2 Then
        MsgBox "You must first Scan in a directory, there are no files Selected", vbInformation
        Exit Sub
    End If

    If pStream.FileHandle > 0 Then pStream.fClose
    If cStream.FileHandle > 0 Then cStream.fClose

    pStream.fOpen hhp, otwriting
    cStream.fOpen hhc, otwriting

    tmp = hhpHeader
    tmp = Replace(tmp, "<filename>.chm", clsfso.GetBaseName(txtOutPutFileName) & ".chm")
    tmp = Replace(tmp, "<deftopic>", cboDefaultPage.text, 1, 1)
    tmp = Replace(tmp, "<title>", txtTitle, 1, 1)

    pStream.WriteLine tmp
    cStream.WriteLine hhcHeader

     For Each n In tv.Nodes
        If n.key = "topLevel" Then GoTo nextOne

        'fullpath minus top level node
        fpath = Replace(n.fullpath, tv.Nodes(1).fullpath & "\", "")

        If n.Image = 2 Then 'isFile
            pStream.WriteLine fpath

            fname = clsfso.GetBaseName(fpath)
            tmp = Replace(fileNode, "*****", fname)
            tmp = br(Replace(tmp, "_____", fpath) & "\n \t ")

            cStream.WriteLine tmp
        Else 'isFolder

            If Len(curfolderParent) = 0 Or Len(ub(dirStack)) = 0 Then
                curfolderParent = n.fullpath
                push dirStack, n.fullpath
                d "push dirstack " & n.fullpath
            Else
                curfolderParent = clsfso.GetParentFolder(n.fullpath)
            End If

            If curfolderParent = ub(dirStack) Then
                If n.fullpath <> ub(dirStack) Then
                    push dirStack, n.fullpath
                    d "push dirstack " & n.fullpath
                End If
            Else
               If InStr(n.fullpath, "bb") > 0 Then
                DoEvents:
               End If
               Do While curfolderParent <> ub(dirStack)
                    d "pop dirstack (" & ub(dirStack) & ")"
                    pop dirStack
                    cStream.WriteLine vbTab & "</UL>"
                    
                    If Len(ub(dirStack)) = 0 Then
                        push dirStack, n.fullpath
                        d "DirStack Closed Out"
                        d "push dirstack (" & n.fullpath & ")"
                        Exit Do
                    End If
                    
               Loop

            End If

            cStream.WriteLine Replace(fldrNode, "_____", clsfso.FolderName(fpath))
        End If

nextOne:
    Next

     Dim i As Long
     If Not aryIsEmpty(dirStack) Then
        For i = 0 To UBound(dirStack)
           cStream.WriteLine vbTab & "</UL>"
        Next
     End If
     
     pStream.WriteLine br("\n \n [INFOTYPES]\n \n ")
     cStream.WriteLine br("</UL>\n </BODY>\n </HTML>")

     pStream.fClose
     cStream.fClose

     ShellExecute Me.hwnd, vbNullString, hhp, vbNullString, 0, 1

End Sub

Sub cmdScan_Click()

    If txtBaseFolder = "" Then
        MsgBox "You Must Enter TopLevel Folder.", vbInformation
        Exit Sub
    End If
    
    If txtExtensions = Empty Then txtExtensions = "*"
    If txtOutPutFileName = Empty Then txtOutPutFileName = "project1"
    If txtTitle = Empty Then txtTitle = "Created with Chm Spider - http://sandsprite.com/"
        
    hhp = txtOutPutFileName
    hhp = txtBaseFolder & "\" & hhp & ".hhp"
    hhc = txtBaseFolder & "\TOC.hhc"
    
    Dim files() As String
    Dim folders() As String
    
    files = clsfso.GetFolderFiles(txtBaseFolder, , False)
    folders = clsfso.GetSubFolders(txtBaseFolder, True)
    
    tv.Nodes.Clear
    tv.Nodes.Add , , "topLevel", txtOutPutFileName, 4
    
    FilterArray files
    
    If Not aryIsEmpty(files) Then
         clsTvm.LoadArrayUnderNode files, "topLevel", 2
    End If
    
    If chkRecursive.value = 0 Then Exit Sub
    
    If Not aryIsEmpty(folders) Then
        Dim i As Long
        For i = 0 To UBound(folders)
            FolderEngine folders(i), "topLevel"
        Next
    End If
    
    tv.Nodes(1).Expanded = True
    
    Call fillComboWithFiles
    
End Sub

Sub fillComboWithFiles() 'and prune tree of empty folders
    Dim n As Node
    Dim pruneTree() As Node
    cboDefaultPage.Clear
    For Each n In tv.Nodes
        If n.Image = 2 Then 'is file
            cboDefaultPage.AddItem Replace(n.fullpath, tv.Nodes(1).text & "\", "", 1, 1)
        Else
            'its a folder, if no subfiles mark for deletion
            If n.Children = 0 Then
                If aryIsEmpty(pruneTree) Then
                    ReDim pruneTree(0)
                    Set pruneTree(0) = n
                Else
                    ReDim Preserve pruneTree(UBound(pruneTree))
                    Set pruneTree(UBound(pruneTree)) = n
                End If
            End If
        End If
    Next
    On Error Resume Next
    If Not aryIsEmpty(pruneTree) Then
        Dim i As Long
        For i = 0 To UBound(pruneTree)
            tv.Nodes.Remove pruneTree(i).Index
        Next
    End If
    cboDefaultPage.ListIndex = 1
End Sub



Sub FolderEngine(fldrpath As String, parentNodeId As String)

    Dim files() As String
    Dim folders() As String
    Dim tmp As String
    Dim myId As String
    Dim baseFldrName As String
    
    files = clsfso.GetFolderFiles(fldrpath, , False)
    folders = clsfso.GetSubFolders(fldrpath)
    
    FilterArray files
    
    baseFldrName = clsfso.FolderName(fldrpath)
    
    myId = clsTvm.AddNodeUnder(baseFldrName, parentNodeId, 1)
    
    If Not aryIsEmpty(files) Then
        clsTvm.LoadArrayUnderNode files, myId, 2
    End If
    
    If Not aryIsEmpty(folders) Then
        Dim i As Long
        For i = 0 To UBound(folders)
             FolderEngine folders(i), myId
        Next
    End If
    
End Sub



Private Sub mnuRemFile_Click()
    If Not curNode Is Nothing Then
        RemoveItemFromCombo cboDefaultPage, Replace(curNode.fullpath, tv.Nodes(1).text, "")
        If curNode.Children = 0 Then tv.Nodes.Remove curNode.Index
    End If
End Sub

Private Sub tv_Collapse(ByVal Node As MSComctlLib.Node)
    If Node.key <> "topLevel" Then
        Node.Image = 1
    End If
End Sub

Private Sub tv_Expand(ByVal Node As MSComctlLib.Node)
    If Node.key <> "topLevel" Then
        Node.Image = 3
    End If
End Sub

Private Sub tv_Mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button <> 2 Then Exit Sub
    
    If curNode Is Nothing Then Exit Sub
    
    If curNode.Image = 2 Then 'isFile
        PopupMenu mnuFilePopup
    Else
        PopupMenu mnuFolderPopup
    End If
    
End Sub

Private Sub mnuRemFolder_Click()
    If curNode Is Nothing Then Exit Sub
    tv.Nodes.Remove curNode.Index
    Call fillComboWithFiles
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    Set curNode = Node
End Sub

Private Sub txtBaseFolder_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    If clsfso.FolderExists(Data.files(1)) Then
         txtBaseFolder = Data.files(1) & "\"
   Else
          MsgBox "Only Drop Folders in here", vbInformation
    End If
End Sub


Private Sub Form_Resize()
    On Error Resume Next

    Frame1.Top = Me.Height - Frame1.Height - 450
    tv.Height = Frame1.Top - tv.Top
        
End Sub

