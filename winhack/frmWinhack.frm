VERSION 5.00
Begin VB.Form frmWinHack 
   Caption         =   "winHack  - http://sandsprite.com"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWinhack.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblCaption 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   900
      TabIndex        =   28
      Text            =   "0"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox lblClass 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   900
      TabIndex        =   27
      Text            =   "0"
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   6
      Left            =   3900
      Picture         =   "frmWinhack.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Toolbox"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   5
      Left            =   3540
      Picture         =   "frmWinhack.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Toolbox"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   4
      Left            =   3180
      Picture         =   "frmWinhack.frx":1D14
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Send Window Message"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   7
      Left            =   2820
      Picture         =   "frmWinhack.frx":2ADE
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "List Box tools"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hwnd On"
      Height          =   255
      Left            =   1680
      TabIndex        =   24
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3360
      TabIndex        =   21
      Top             =   420
      Width           =   915
      Begin VB.PictureBox picColor 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   735
         TabIndex        =   23
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtColor 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   60
         TabIndex        =   22
         Top             =   1800
         Width           =   795
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "frmWinhack.frx":3078
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   3
      Left            =   2460
      Picture         =   "frmWinhack.frx":3382
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Analyze Menus"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   2
      Left            =   2100
      Picture         =   "frmWinhack.frx":42EB
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Enumerate Children"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   1
      Left            =   1740
      Picture         =   "frmWinhack.frx":4FF7
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Disable Window"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   0
      Left            =   1380
      Picture         =   "frmWinhack.frx":59B7
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Enable Window"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtParentHwnd 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   900
      TabIndex        =   13
      Text            =   "0"
      Top             =   1920
      Width           =   1635
   End
   Begin VB.TextBox txthWnd 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   900
      OLEDragMode     =   1  'Automatic
      TabIndex        =   12
      Text            =   "0"
      Top             =   480
      Width           =   1635
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4560
      Top             =   2280
   End
   Begin VB.Label lblPath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2025
      Left            =   0
      TabIndex        =   26
      Top             =   2700
      Width           =   7995
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Parent Details"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label lblPHwndTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hWnd:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label lblPTextTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   2160
      Width           =   435
   End
   Begin VB.Label lblPClassTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label lblPCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   900
      TabIndex        =   7
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label lblPClass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   900
      TabIndex        =   6
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label lblParent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   900
      TabIndex        =   5
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label lblParentTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parent:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblClassTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   480
   End
   Begin VB.Label lblTextTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   435
   End
   Begin VB.Label lblCtrlDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Control Details"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   1200
   End
   Begin VB.Label lblHwndTItle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hWnd:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   525
   End
End
Attribute VB_Name = "frmWinHack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cnt As Long




Private Sub Check1_Click()
    Timer1.Enabled = IIf(Check1.value = 1, True, False)
    On Error Resume Next
    If Not Timer1.Enabled Then RemoveRectOutline CLng(txthWnd)
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error Resume Next
    Dim hwnd As Long
    hwnd = CLng(txthWnd)
    Select Case Index
        Case 0: If hwnd > 0 Then Call EnableWindow(hwnd, 1)
        Case 1: If hwnd > 0 Then Call EnableWindow(hwnd, 0)
        Case 2: frmChildren.Show
        Case 3: frmMenus.Show
        Case 4: frmMessages.Show
        Case 5: frmConv.Show
        'Case 6: frmMagnifier.Show
        Case 7: frmListSnag.Show
    End Select
End Sub

 

Private Sub Form_Resize()
    On Error Resume Next
    'If Me.Width < 4410 Then Me.Width = 4410
    'Me.Height = 3030
End Sub

Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    Dim f As Variant
    
    If IsNumeric(txthWnd) Then RemoveRectOutline CLng(txthWnd)
    
    For Each f In Forms
        Unload f
    Next
    Unload Me
    End
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = 99 'custom
    Screen.MouseIcon = LoadResPicture(101, vbResIcon)
    Timer1.Enabled = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cnt = 0
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = 0
    Timer1.Enabled = False
    On Error Resume Next
    RemoveRectOutline CLng(txthWnd)
End Sub


 

Private Sub txthWnd_Change()
    Dim hwnd As Long
    
    If txthWnd = Empty Then txthWnd = 0
    
    On Error Resume Next
    
    hwnd = CLng(txthWnd)
    If Err.Number <> 0 Then
        Err.Clear
        hwnd = CLng("&h" & txthWnd)
        If Err.Number <> 0 Then Exit Sub
    End If
    
    lblCaption = GetCaption(hwnd)
    lblClass = GetClass(hwnd)
    lblParent = GetParent(hwnd)
    lblPClass.Caption = GetClass(lblParent)
    lblPCaption.Caption = GetCaption(lblParent)
    txtParentHwnd = lblParent
    
    OutlineWindow hwnd
    
End Sub

Private Sub txthWnd_DragDrop(Source As Control, x As Single, y As Single)
    On Error Resume Next
    txthWnd = Source.Text
End Sub

Private Sub txthwnd_gotfocus()
    txthWnd.SelStart = 0
    txthWnd.SelLength = Len(txthWnd)
End Sub

Private Sub txthWnd_DblClick()
    Screen.MousePointer = vbHourglass
    Sleep 300
    Clipboard.Clear
    Clipboard.SetText txthWnd
    Screen.MousePointer = vbDefault
End Sub

Private Sub Timer1_Timer()
    Dim hwnd As Long, hdc As Long, pxColor As Double
    Dim pt As PointAPI
 
    GetCursorPos pt
    Me.Caption = "winHack   Last X: " & pt.x & "  Last Y:" & pt.y
    hwnd = WindowFromPoint(pt.x, pt.y)
    
    If Not IsNumeric(txthWnd) Then txthWnd = 0
    If txthWnd <> 0 And CLng(txthWnd) <> hwnd Then 'remove any rect we drew
       RemoveRectOutline CLng(txthWnd)
    End If
    
    txthWnd = hwnd
    
    hdc = GetDC(0)
    pxColor = GetPixel(hdc, pt.x, pt.y)
    ReleaseDC 0, hdc
    
    picColor.BackColor = pxColor
    txtColor = Hex(pxColor)
   
    lblPath.Caption = GetProcessPath(hwnd)
     
    'If MagnifierVisible Then frmMagnifier.Magnify pt.x, pt.y
End Sub


