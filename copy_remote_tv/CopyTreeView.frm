VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CopyTreeView 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MathImagics LV Copy"
   ClientHeight    =   2895
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer myTimer 
      Left            =   3600
      Top             =   2040
   End
   Begin VB.CommandButton bCopy 
      BackColor       =   &H004040FF&
      Caption         =   "Locate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   -45
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1125
      Width           =   4095
   End
   Begin ComctlLib.TreeView myTreeView 
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   327682
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mPopup 
      Caption         =   "Popup"
      Begin VB.Menu mRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mDummy 
         Caption         =   "-"
      End
      Begin VB.Menu mLocate 
         Caption         =   "Locate"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "CopyAll"
      End
   End
End
Attribute VB_Name = "CopyTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=======================================================================
'
'  (c) 2002  MathImagical Software, Uki, NSW, Australia
'
'  This program was written by Jim White (aka Dr Memory), May 2002.
'
'  CopyTreeView  - demonstrates the dmFPDE_TreeView functions.
'  ============
'           This demo allows the user to locate a TreeView
'           in another applications, and produce a copy
'           in its own TreeView.
'
'========================================================================
'
' Some API functions are needed by the Demo form to help us identify target windows
' from the mouse position - when the cursor is over a TreeView we make the
' button green and enable it - just press return and we'll "clone" that TreeView.
'
' Notes (July 2002)
'
'    On startup, the form displays a red button labelled "Locate".
'    When you move the mouse over other a TreeView window in another
'    application, the label will turn green, and display that TreeView's
'    window handle, and the number of items in it.
'
'    When the button is green, click on it to create a copy. The copy will
'    be re-sized to match the target settings.
'
'    The detection of window details under a given cursor position is done in the
'    myTimer_Timer sub.
'
'=========================================================================

   Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
   Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
   Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
   Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  
   Private Type POINTAPI
      X        As Long
      y        As Long
      End Type

'========================================================================
   Dim TargetTreeview As dmTreeView    ' attributes of target window
'========================================================================
   Dim Splashed As Boolean

Dim tvEntries() As String

Private Sub Form_Load()
    mPopup.Visible = False
   dmCrashMode 0
   dmSetTreeViewColor myTreeView, &HF0FFFF
'   If Not Splashed Then
'      dmSplash.Show
'      DoEvents
'      While dmSplash.Visible: DoEvents: Wend
'      Unload dmSplash
'      End If
   myTimer.Interval = 500
   End Sub

Private Sub bCopy_Click()
   
   If TargetTreeview.hWnd = 0 Then Exit Sub
      
   If Not dmGetTreeviewInfo(TargetTreeview) Then
      MsgBox "TreeView handle " & TargetTreeview.hWnd & " is no longer valid", vbInformation
      If bCopy.Visible Then bCopy.BackColor = &H4040FF ' came here via popup
      Exit Sub
      End If
   
   If TargetTreeview.ItemCount < 1 Then
      MsgBox "No data items identified in target window", vbInformation
      Exit Sub
      End If
      
      
   Erase tvEntries
   myTimer.Enabled = False
   bCopy.Visible = False
   CopyTargetTreeview
   
   End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Private Sub CopyTargetTreeview()

   Dim item&, nItems&, iText$, iKey$
   
   Me.Caption = "Target handle = " & TargetTreeview.hWnd

   dmTreeviewScan TargetTreeview    ' take the snapshot of the target tree
   
   '===============================================================================================
   ' 1. the tables tvNext, tvPrev, tvChild, tvParent now correspond to Node.Next, Node.Prev etc
   ' 2. the first root-level node is item 1
   ' 3. the text for each item is in tvText(item)
   '
   ' To build a copy of the tree we traverse the tv-tables just as we would a TreeView.Nodes collection
   '===============================================================================================
   
   nItems = tvCount
   
   GoSub SetupMyTreeview
   item = 1
   While item > 0
      iText = tvText(item)
      If iText = "" Then iText = "<empty>"
      iKey = "N" & item
      myTreeView.Nodes.Add , , iKey, iText
      push tvEntries, iText
      CopySubtree item
      If tvExpanded(item) Then myTreeView.Nodes(iKey).Expanded = True
      item = tvNext(item)
      push tvEntries, Empty
      Wend
   Exit Sub
   
SetupMyTreeview:
   Dim mScalemode As Integer
   mScalemode = Me.ScaleMode
   
   myTreeView.Nodes.Clear
   '
   ' Adjust the form size so that Form_Resize will make our Treeview
   '    to be the same size as the source Treeview
   '
   With TargetTreeview
      MoveWindow Me.hWnd, ScaleX(Me.Left, mScalemode, vbPixels), ScaleY(Me.Top, mScalemode, vbPixels), _
                         ScaleX(Me.Width, mScalemode, vbPixels) + .Right - .Left - ScaleX(myTreeView.Width, mScalemode, vbPixels), _
                         ScaleY(Me.Height, mScalemode, vbPixels) + .Bottom - .Top - ScaleY(myTreeView.Height, mScalemode, vbPixels), 1
      End With
   Return
   
   End Sub
   
Private Sub CopySubtree(ByVal pItem As Long, Optional depth As Long = 1)
   
   Dim pKey As String, sKey As String   ' parent and sibling keys
   Dim firstchild As Boolean
   Dim item As Long, iKey As String, iText As String
   
   item = tvChild(pItem)
   If item = 0 Then Exit Sub  ' childless
   
   firstchild = True
   pKey = "N" & pItem
      
   While item <> 0
      iText = tvText(item): If iText = "" Then iText = "<empty>"
      iKey = "N" & item
      If firstchild Then
         myTreeView.Nodes.Add pKey, tvwChild, iKey, iText
         push tvEntries, String(depth, vbTab) & iText
      Else
         myTreeView.Nodes.Add sKey, tvwNext, iKey, iText
         push tvEntries, String(depth, vbTab) & iText
         End If
      firstchild = False
      sKey = iKey
      CopySubtree item, (depth + 1)
      If tvExpanded(item) Then myTreeView.Nodes(iKey).Expanded = True
      item = tvNext(item)
      Wend
   Exit Sub
   End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then Exit Sub
   myTreeView.Top = 0: myTreeView.Height = Me.ScaleHeight
   myTreeView.Left = 0: myTreeView.Width = Me.ScaleWidth
   End Sub

Private Sub mLocate_Click()
   Unload Me
   DoEvents
   Splashed = True
   Show
   Refresh
   End Sub

Private Sub mnuCopyAll_Click()
    Clipboard.Clear
    'Clipboard.SetText Join(tvText, vbCrLf)
    Clipboard.SetText Join(tvEntries, vbCrLf)
    MsgBox UBound(tvText) & " items copied to clipboard", vbInformation
End Sub

Private Sub mRefresh_Click()
   bCopy_Click
   End Sub

Private Sub myTreeView_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = 2 Then PopupMenu mPopup
   End Sub

Private Sub myTimer_Timer()
   Static lastWindow As Long
   Dim mousePos As POINTAPI
   Dim Cwindow As Long
   '
   '  If the mouse passes over a foreign-process TreeView
   '  give the copy button the green light
   '
   GetCursorPos mousePos
   Cwindow = WindowFromPoint(mousePos.X, mousePos.y)
   If Cwindow = lastWindow Then Exit Sub
   lastWindow = Cwindow
   
   If Cwindow = Me.hWnd Then Exit Sub
   If GetParent(Cwindow) = Me.hWnd Then Exit Sub
   
   Dim tvText As String
   Dim tvInfo As dmTreeView
   
   tvInfo.hWnd = Cwindow
   If dmGetTreeviewInfo(tvInfo) Then
      bCopy.Enabled = True
      bCopy.BackColor = &HA000&
      tvText = "Class = " & tvInfo.Class & vbLf & tvInfo.ItemCount & " items"
      bCopy.Caption = "hWnd = " & Cwindow & vbLf & tvText
      LSet TargetTreeview = tvInfo
      End If
   End Sub
