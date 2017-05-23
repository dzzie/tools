VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMessages 
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form3"
   ScaleHeight     =   5760
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   5280
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMessages.frx":0000
   End
   Begin VB.CommandButton cmdRtfColor 
      Caption         =   "rtf selback color test"
      Height          =   375
      Left            =   3600
      TabIndex        =   30
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "MoveWindow( hwnd, x, y, w, h, repaint)"
      Height          =   915
      Left            =   60
      TabIndex        =   23
      Top             =   4200
      Width           =   5415
      Begin VB.CommandButton cmdMoveWindow 
         Caption         =   "Move Window"
         Height          =   375
         Left            =   3660
         TabIndex        =   29
         Top             =   420
         Width           =   1635
      End
      Begin VB.TextBox txtMovePos 
         Height          =   315
         Index           =   3
         Left            =   2640
         TabIndex        =   28
         Text            =   "0"
         Top             =   480
         Width           =   795
      End
      Begin VB.TextBox txtMovePos 
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   27
         Text            =   "0"
         Top             =   480
         Width           =   795
      End
      Begin VB.TextBox txtMovePos 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   26
         Text            =   "0"
         Top             =   480
         Width           =   795
      End
      Begin VB.TextBox txtMovePos 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Text            =   "0"
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "left             top              width           height"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "SetWindowPos( hwnd, hWndInsertAfter, x, y, cx, cy, wFlags)"
      Height          =   1275
      Left            =   60
      TabIndex        =   8
      Top             =   2880
      Width           =   5415
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmMessages.frx":008B
         Left            =   1920
         List            =   "frmMessages.frx":009E
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SetWindowPosition"
         Height          =   375
         Left            =   3660
         TabIndex        =   14
         Top             =   780
         Width           =   1635
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmMessages.frx":0124
         Left            =   120
         List            =   "frmMessages.frx":0134
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1755
      End
      Begin VB.TextBox txtPos 
         Height          =   315
         Index           =   4
         Left            =   2700
         TabIndex        =   12
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPos 
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   11
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPos 
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   10
         Text            =   "0"
         Top             =   840
         Width           =   795
      End
      Begin VB.TextBox txtPos 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Text            =   "0"
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "left             top              width           height"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ShowWindow( hwnd,  nCmdShow )"
      Height          =   735
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "ShowWindow"
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   240
         Width           =   1395
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMessages.frx":017F
         Left            =   240
         List            =   "frmMessages.frx":01A7
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   3420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SendMessage( hwnd, wMsg, wParam, lParam )"
      Height          =   1995
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   5475
      Begin VB.TextBox txtReturn 
         Height          =   285
         Left            =   4080
         TabIndex        =   22
         Top             =   1500
         Width           =   1155
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   780
         TabIndex        =   21
         Text            =   "Combo5"
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtMsgNum 
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Text            =   "0"
         Top             =   1500
         Width           =   1155
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SendMessage"
         Height          =   375
         Left            =   3960
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtMsgWParam 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Text            =   "0"
         Top             =   1500
         Width           =   1155
      End
      Begin VB.TextBox txtMsgLParam 
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Text            =   "0"
         Top             =   1500
         Width           =   1080
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMessages.frx":0258
         Left            =   780
         List            =   "frmMessages.frx":025A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   4500
      End
      Begin VB.Label Label2 
         Caption         =   "EM"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   780
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "WM"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "       wMsg          /     wParam       /     lParam"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   1140
         Width           =   3150
      End
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1

Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9
Private Const SW_SHOW = 5
Private Const SW_SHOWDEFAULT = 10
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOWNORMAL = 1

Private msgNum As Long
Private Const LF_FACESIZE = 32

Private Type CHARFORMAT
    cbSize As Integer                           '2
    wPad1 As Integer                            '4
    dwMask As Long                              '8
    dwEffects As Long                           '12
    yHeight As Long                             '16
    yOffset As Long                             '20
    crTextColor As Long                         '24
    bCharSet As Byte                            '25
    bPitchAndFamily As Byte                     '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte    '58
End Type


'https://msdn.microsoft.com/en-us/library/windows/desktop/bb787883(v=vs.85).aspx
Private Type CHARFORMAT2
    cbSize As Integer                           '2
    wPad1 As Integer                            '4
    dwMask As Long                              '8
    dwEffects As Long                           '12
    yHeight As Long                             '16
    yOffset As Long                             '20
    crTextColor As Long                         '24
    bCharSet As Byte                            '25
    bPitchAndFamily As Byte                     '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte    '58
    wPad2 As Integer                            '60

    ' Additional stuff supported by RICHEDIT20
    wWeight As Integer         'Font weight (LOGFONT value)
    sSpacing As Integer        'Amount to space between letters
    crBackColor As Long        'Background color
    lLCID As Long              'Locale ID
    dwReserved As Long         'Reserved. Must be 0
    sStyle As Integer          'Style handle
    wKerning As Integer        'Twip size above which to kern char pair
    bUnderlineType As Byte     'Underline type
    bAnimation As Byte         'Animated text like marching ants
    bRevAuthor As Byte         'Revision author index
    bReserved1 As Byte
End Type



Private Const EM_SETCHARFORMAT = (&H400 + 68)
Private Const CFM_BACKCOLOR = &H4000000


Private Sub cmdMoveWindow_Click()
    On Error GoTo hell
    MoveWindow CLng(frmWinHack.txthWnd), CLng(txtMovePos(0)), CLng(txtMovePos(1)), CLng(txtMovePos(2)), CLng(txtMovePos(3)), 1
    
    Exit Sub
hell: MsgBox Err.Description
End Sub

Private Sub cmdRtfColor_Click()

    Dim tCF2 As CHARFORMAT2
    tCF2.dwMask = CFM_BACKCOLOR
    tCF2.crBackColor = vbYellow
    tCF2.cbSize = Len(tCF2)
    SendMessage CLng(frmWinHack.txthWnd), EM_SETCHARFORMAT, 1, tCF2
    
End Sub

Private Sub Command1_Click()
    Dim a&
    
        Select Case Combo1.ListIndex
            Case 0: a = SW_HIDE
            Case 1: a = SW_MAXIMIZE
            Case 2: a = SW_MINIMIZE
            Case 3: a = SW_RESTORE
            Case 4: a = SW_SHOW
            Case 5: a = SW_SHOWDEFAULT
            Case 6: a = SW_SHOWMAXIMIZED
            Case 7: a = SW_SHOWMINIMIZED
            Case 8: a = SW_SHOWMINNOACTIVE
            Case 9: a = SW_SHOWNA
            Case 10: a = SW_SHOWNOACTIVATE
            Case 11: a = SW_SHOWNORMAL
        End Select
        
        ShowWindow CLng(frmWinHack.txthWnd), a
    
End Sub

Private Sub Combo2_Click()
  On Error GoTo shit
     l = Combo2.Text
     t = Split(Combo2.Text, " = ")
     txtMsgNum = Hex(t(1))
  Exit Sub
shit: MsgBox Err.Description
End Sub

Private Sub Combo5_Click()
  On Error GoTo shit
     l = Combo5.Text
     t = Split(Combo5.Text, " = ")
     txtMsgNum = Hex(t(1))
  Exit Sub
shit: MsgBox Err.Description
End Sub

Private Sub Command2_Click()
    On Error GoTo hell
    
    If txtMsgWParam = Empty Then txtMsgWParam = 0
    If txtMsgLParam = Empty Then txtMsgLParam = 0
    
    Screen.MousePointer = vbHourglass
    txtReturn = SendMessage(CLng(frmWinHack.txthWnd), CLng("&H" & txtMsgNum), CLng("&H" & txtMsgWParam), CLng("&H" & txtMsgLParam))
    Screen.MousePointer = vbDefault
        
    Exit Sub
hell: MsgBox Err.Description & vbCrLf & vbCrLf & "Are all your inputs in Hex?"
End Sub

Private Sub Command3_Click()
   Dim Zflags&, flags&
   
   Select Case Combo3.ListIndex
        Case 0: Zflags = HWND_BOTTOM    '= 1
        Case 1: Zflags = HWND_NOTOPMOST '= -2
        Case 2: Zflags = HWND_TOP       '= 0
        Case 3: Zflags = HWND_TOPMOST   '= -1
    End Select
   
   Select Case Combo4.ListIndex
        Case 0: flags = SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        Case 1: flags = SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        Case 2: flags = SWP_SHOWWINDOW
        Case 3: flags = SWP_HIDEWINDOW Or SWP_NOACTIVATE
        Case 4: flags = SWP_HIDEWINDOW Or SWP_NOREDRAW Or SWP_NOACTIVATE
    End Select
    
   SetWindowPos CLng(frmWinHack.txthWnd), Zflags, CLng(txtPos(1)), CLng(txtPos(2)), CLng(txtPos(3)), CLng(txtPos(4)), flags

End Sub



Private Sub Form_Load()

 Dim winMsg() As String
  
 push winMsg(), " WM_USER = &H400"
 push winMsg(), " WM_ACTIVATE = &H6"
 push winMsg(), " WM_ACTIVATEAPP = &H1C"
 push winMsg(), " WM_CHAR = &H102"
 push winMsg(), " WM_CHARTOITEM = &H2F"
 push winMsg(), " WM_CHILDACTIVATE = &H22"
 push winMsg(), " WM_CLEAR = &H303"
 push winMsg(), " WM_CLOSE = &H10"
 push winMsg(), " WM_COMMAND = &H111"
 push winMsg(), " WM_COPY = &H301"
 push winMsg(), " WM_CTLCOLORBTN = &H135"
 push winMsg(), " WM_CTLCOLORDLG = &H136"
 push winMsg(), " WM_CTLCOLOREDIT = &H133"
 push winMsg(), " WM_CTLCOLORLISTBOX = &H134"
 push winMsg(), " WM_CTLCOLORMSGBOX = &H132"
 push winMsg(), " WM_CTLCOLORSCROLLBAR = &H137"
 push winMsg(), " WM_CTLCOLORSTATIC = &H138"
 push winMsg(), " WM_CUT = &H300"
 push winMsg(), " WM_DEADCHAR = &H103"
 push winMsg(), " WM_DELETEITEM = &H2D"
 push winMsg(), " WM_DESTROY = &H2"
 push winMsg(), " WM_DESTROYCLIPBOARD = &H307"
 push winMsg(), " WM_ENABLE = &HA"
 push winMsg(), " WM_ERASEBKGND = &H14"
 push winMsg(), " WM_EXITMENULOOP = &H212"
 push winMsg(), " WM_FONTCHANGE = &H1D"
 push winMsg(), " WM_GETTEXT = &HD"
 push winMsg(), " WM_GETTEXTLENGTH = &HE"
 push winMsg(), " WM_HOTKEY = &H312"
 push winMsg(), " WM_HSCROLL = &H114"
 push winMsg(), " WM_HSCROLLCLIPBOARD = &H30E"
 push winMsg(), " WM_ICONERASEBKGND = &H27"
 push winMsg(), " WM_IME_KEYDOWN = &H290"
 push winMsg(), " WM_IME_KEYLAST = &H10F"
 push winMsg(), " WM_IME_KEYUP = &H291"
 push winMsg(), " WM_IME_NOTIFY = &H282"
 push winMsg(), " WM_IME_SELECT = &H285"
 push winMsg(), " WM_IME_SETCONTEXT = &H281"
 push winMsg(), " WM_IME_STARTCOMPOSITION = &H10D"
 push winMsg(), " WM_INITDIALOG = &H110"
 push winMsg(), " WM_INITMENU = &H116"
 push winMsg(), " WM_INITMENUPOPUP = &H117"
 push winMsg(), " WM_KEYDOWN = &H100"
 push winMsg(), " WM_KEYFIRST = &H100"
 push winMsg(), " WM_KEYLAST = &H108"
 push winMsg(), " WM_KEYUP = &H101"
 push winMsg(), " WM_KILLFOCUS = &H8"
 push winMsg(), " WM_LBUTTONDBLCLK = &H203"
 push winMsg(), " WM_LBUTTONDOWN = &H201"
 push winMsg(), " WM_LBUTTONUP = &H202"
 push winMsg(), " WM_MBUTTONDBLCLK = &H209"
 push winMsg(), " WM_MBUTTONDOWN = &H207"
 push winMsg(), " WM_MBUTTONUP = &H208"
 push winMsg(), " WM_MDIACTIVATE = &H222"
 push winMsg(), " WM_MDICASCADE = &H227"
 push winMsg(), " WM_MDICREATE = &H220"
 push winMsg(), " WM_MDIDESTROY = &H221"
 push winMsg(), " WM_MDIGETACTIVE = &H229"
 push winMsg(), " WM_MDIICONARRANGE = &H228"
 push winMsg(), " WM_MDIMAXIMIZE = &H225"
 push winMsg(), " WM_MDINEXT = &H224"
 push winMsg(), " WM_MDIREFRESHMENU = &H234"
 push winMsg(), " WM_MDIRESTORE = &H223"
 push winMsg(), " WM_MDISETMENU = &H230"
 push winMsg(), " WM_MDITILE = &H226"
 push winMsg(), " WM_MEASUREITEM = &H2C"
 push winMsg(), " WM_MENUCHAR = &H120"
 push winMsg(), " WM_MENUSELECT = &H11F"
 push winMsg(), " WM_MOUSEACTIVATE = &H21"
 push winMsg(), " WM_MOUSEFIRST = &H200"
 push winMsg(), " WM_MOUSELAST = &H209"
 push winMsg(), " WM_MOUSEMOVE = &H200"
 push winMsg(), " WM_MOVE = &H3"
 push winMsg(), " WM_NCACTIVATE = &H86"
 push winMsg(), " WM_NCCALCSIZE = &H83"
 push winMsg(), " WM_NCCREATE = &H81"
 push winMsg(), " WM_NCDESTROY = &H82"
 push winMsg(), " WM_NCHITTEST = &H84"
 push winMsg(), " WM_NCLBUTTONDBLCLK = &HA3"
 push winMsg(), " WM_NCLBUTTONDOWN = &HA1"
 push winMsg(), " WM_NCLBUTTONUP = &HA2"
 push winMsg(), " WM_NCMBUTTONDBLCLK = &HA9"
 push winMsg(), " WM_NCMBUTTONDOWN = &HA7"
 push winMsg(), " WM_NCMBUTTONUP = &HA8"
 push winMsg(), " WM_NCMOUSEMOVE = &HA0"
 push winMsg(), " WM_NCPAINT = &H85"
 push winMsg(), " WM_NCRBUTTONDBLCLK = &HA6"
 push winMsg(), " WM_NCRBUTTONDOWN = &HA4"
 push winMsg(), " WM_NCRBUTTONUP = &HA5"
 push winMsg(), " WM_NEXTDLGCTL = &H28"
 push winMsg(), " WM_NULL = &H0"
 push winMsg(), " WM_PAINT = &HF"
 push winMsg(), " WM_PAINTCLIPBOARD = &H309"
 push winMsg(), " WM_PAINTICON = &H26"
 push winMsg(), " WM_PALETTECHANGED = &H311"
 push winMsg(), " WM_PALETTEISCHANGING = &H310"
 push winMsg(), " WM_PARENTNOTIFY = &H210"
 push winMsg(), " WM_PASTE = &H302"
 push winMsg(), " WM_QUIT = &H12"
 push winMsg(), " WM_RBUTTONDBLCLK = &H206"
 push winMsg(), " WM_RBUTTONDOWN = &H204"
 push winMsg(), " WM_RBUTTONUP = &H205"
 push winMsg(), " WM_SETCURSOR = &H20"
 push winMsg(), " WM_SETFOCUS = &H7"
 push winMsg(), " WM_SETFONT = &H30"
 push winMsg(), " WM_SETHOTKEY = &H32"
 push winMsg(), " WM_SETREDRAW = &HB"
 push winMsg(), " WM_SETTEXT = &HC"
 push winMsg(), " WM_SHOWWINDOW = &H18"
 push winMsg(), " WM_SIZE = &H5"
 push winMsg(), " WM_SIZECLIPBOARD = &H30B"
 push winMsg(), " WM_SPOOLERSTATUS = &H2A"
 push winMsg(), " WM_SYSCHAR = &H106"
 push winMsg(), " WM_SYSCOLORCHANGE = &H15"
 push winMsg(), " WM_SYSCOMMAND = &H112"
 push winMsg(), " WM_SYSDEADCHAR = &H107"
 push winMsg(), " WM_SYSKEYDOWN = &H104"
 push winMsg(), " WM_SYSKEYUP = &H105"
 push winMsg(), " WM_TIMECHANGE = &H1E"
 push winMsg(), " WM_TIMER = &H113"
 push winMsg(), " WM_UNDO = &H304"
 push winMsg(), " WM_VKEYTOITEM = &H2E"
 push winMsg(), " WM_VSCROLL = &H115"
 push winMsg(), " WM_VSCROLLCLIPBOARD = &H30A"
 push winMsg(), " WM_WINDOWPOSCHANGED = &H47"
 push winMsg(), " WM_WINDOWPOSCHANGING = &H46"
  
 For i = 0 To UBound(winMsg)
    Combo2.AddItem winMsg(i)
 Next
 
 Erase winMsg()
 

push winMsg(), "EM_GETLINECOUNT = &HBA"
push winMsg(), "EM_LIMITTEXT = &HC5"
push winMsg(), "EM_LINESCROLL = &HB6"
push winMsg(), "EM_REPLACESEL = &HC2"
push winMsg(), "EM_SETPASSWORDCHAR = &HCC"
push winMsg(), "EM_SETREADONLY = &HCF"
push winMsg(), "EM_SETSEL = &HB1"

 
For i = 0 To UBound(winMsg)
    Combo5.AddItem winMsg(i)
 Next
 
 Combo1.ListIndex = 0
 Combo2.ListIndex = 0
 Combo3.ListIndex = 0
 Combo4.ListIndex = 0
 Combo5.ListIndex = 0
 
 Me.Caption = "Send Window Message for hWnd: " & frmWinHack.txthWnd
 
End Sub



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



'ESB_DISABLE_DOWN = &H2
'ESB_DISABLE_LEFT = &H1
'ESB_DISABLE_RIGHT = &H2
'ESB_DISABLE_UP = &H1
