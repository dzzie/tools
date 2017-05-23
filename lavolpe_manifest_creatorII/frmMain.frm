VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   6360
   Begin ComctlLib.TreeView tvItems 
      Height          =   6060
      Left            =   90
      TabIndex        =   0
      Top             =   375
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   10689
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   450
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList16"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtHelp 
      Height          =   915
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMain.frx":0000
      Top             =   6465
      Width           =   6176
   End
   Begin VB.Timer MenuAction 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   4905
      Top             =   30
   End
   Begin ComctlLib.ImageList ImageList24 
      Left            =   3720
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0006
            Key             =   "select"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0320
            Key             =   "unselect"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0A32
            Key             =   "child"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0D4C
            Key             =   "rtOS"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1066
            Key             =   "pc"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1380
            Key             =   "rtApp"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":169A
            Key             =   "rtDep"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":19B4
            Key             =   "rtWS"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1CCE
            Key             =   "rsFile"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList32 
      Left            =   5310
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1FE8
            Key             =   "select"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2302
            Key             =   "unselect"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2F54
            Key             =   "child"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":326E
            Key             =   "rtOS"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3588
            Key             =   "pc"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":38A2
            Key             =   "rtApp"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3BBC
            Key             =   "rtDep"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3ED6
            Key             =   "rtWS"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":41F0
            Key             =   "rsFile"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList16 
      Left            =   4335
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":450A
            Key             =   "select"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4824
            Key             =   "unselect"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":49FE
            Key             =   "child"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4D18
            Key             =   "rtOS"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5032
            Key             =   "pc"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":534C
            Key             =   "rtApp"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5666
            Key             =   "rtDep"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5980
            Key             =   "rtWS"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5C9A
            Key             =   "rsFile"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Only checked items are included in the manifest. Right click to edit."
      Height          =   270
      Left            =   135
      TabIndex        =   1
      Top             =   120
      Width           =   6030
   End
   Begin VB.Menu mnuMain 
      Caption         =   "The Manifest"
      Index           =   0
      Begin VB.Menu mnuCreate 
         Caption         =   "New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Load from File"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Load from Resource File (res)"
         Index           =   3
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Create from Project File (vbp)"
         Index           =   4
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Load from Clipboard"
         Index           =   6
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Append/Merge Manifest"
         Index           =   8
         Begin VB.Menu mnuAppend 
            Caption         =   "From File"
            Index           =   0
         End
         Begin VB.Menu mnuAppend 
            Caption         =   "From Clipboard"
            Index           =   1
         End
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Export Manifest"
         Index           =   10
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Miscellaneous"
      Index           =   2
      Begin VB.Menu mnuInsert 
         Caption         =   "Add Dependent Assembly Template to Tree"
         Index           =   0
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Clone Selected Dependent Assembly"
         Index           =   1
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Add File Dependency to Tree"
         Index           =   3
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Add Assembly File Element (New or Selected File) "
         Index           =   4
         Begin VB.Menu mnuAssmbly 
            Caption         =   "comClass Element"
            Index           =   0
         End
         Begin VB.Menu mnuAssmbly 
            Caption         =   "typeLib Element"
            Index           =   1
         End
         Begin VB.Menu mnuAssmbly 
            Caption         =   "comInterfaceProxyStub"
            Index           =   2
         End
         Begin VB.Menu mnuAssmbly 
            Caption         =   "windowClass"
            Index           =   3
         End
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Add Assembly NoInheritable Element"
         Index           =   5
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Add Assembly comInterfaceExternalProxyStub"
         Index           =   6
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Add Assembly windowClass Element"
         Index           =   7
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Add progID to Selected comClass Element"
         Index           =   8
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Help"
      Index           =   3
      Begin VB.Menu mnuHelp 
         Caption         =   "View Microsoft Related Site"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Microsoft: Assembly Manifest Site"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Copy Sub Main() to Clipboard"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Update 9 Apr 17
'   Located all proper window settings name spaces in the registry. Fixed all that were incorrect
'       fyi: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\SMI\WinSxS Settings
'       removed magicFutureSetting since I cannot find it registered in any WindowsSettings namespace
'       found an additional setting; but no documentation at all on web: forceFocusBasedMouseWheel
' Update 10 Apr 17
'   Fixed error where selected Qualified Name context menu did not exit menu routine
'   Added/updated some help prompts when updating assemblyIdentity attributes
'   Fixed logic flaw that would not allow more than 1 <file> element to be imported
' Update 13 Apr 17
'   Fixed error where adding <file> from Miscellaneous menu could result in duplicate key
'   Removed -1 as a valid attribute index in the cManifestEntryEx methods
'   Added shift+right click option to unhide attrs and/or override fixed values by allowing manual editing
' Update 15 Apr 17
'   Added subclassing to restrict window width/height. Having fixed size could result in window too large in 200% DPI
'   Customized resizing of treeview if app ran in 200% DPI or other DPI resulting in non-whole number TwipsPerPixel
' Upate 26 Apr 17 -- attempt to make project aware of assembly manifests, not just application manifests
'   Added support for all known assembly-manifest related elements (8 additional elements)
'   Ensured required valid empty/blank attributes are written to manifest when empty attributes are exported
'   Minor changes: prefixed element names no longer default in Output form
' Update 29 Apr 17 -- fluff & minor changes
'   Added frmMultiAttrValue from displayed for elements that can use comma-delmited value list
'   Added a few more help tips for assembly-manifest related items
'   Added support for assembly-manifest windowClass element
'   All manifest items added via the Manifest|Append/Merge Manifest menu can be deleted by user

'////////////////////////////////////////////               \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' The manifest returned by the cManifestEx class is guaranteed to have these elements & their known children
'   Only selected elements are exported to a manifest. Items listed are in order written to manifest
'   <assembly>              root node
'       <noInherit/>
'       <assemblyIdentity/>
'       <description/>
'       <dependency/>*      contains dependent assembly information
'       <file/>*            contains private file information
'       <trustInfo/>        contains UAC information
'       <compatibility/>    contains O/S compatibility information
'       <application/>      contains window settings information
' (*) = can have multiple instances; otherwise only one instance is allowed
' (*) = can be dynamically added & deleted. Else elements cannot be deleted, can be unselected

' If new entries are made to the base manifest (cManifestEntryEx updated) lots of work to do here:
'   If re-organizing the treeview, modify TheManifest.Begin, .ElementAdded, & .Finished as needed
'   modify the MenuAction_Timer() routine as needed (this is the right-click menu actions)
'   modify the pvSetHelpText() routine

' Tip: Check back with this URL every now & again to look for additions/changes
' New software versions and/or revisions may trigger changes to the Manifest schema
'   https://msdn.microsoft.com/en-us/library/windows/desktop/aa374191(v=vs.85).aspx
'   For example, within a week of rewriting this project, msdn added two new items to that page:
'       1) 2017/WindowsSettings: gdiScaling
'       2) an extra attribute value for dpiAwareness: permonitorv2
'   Update cManifestEx as needed


' Create JIT popup menus
Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Private Declare Function AppendMenu Lib "user32.dll" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal RECTD As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Const TPM_LEFTALIGN As Long = &H0&
Private Const TPM_NOANIMATION As Long = &H4000&
Private Const TPM_RETURNCMD As Long = &H100&
Private Const TPM_TOPALIGN As Long = &H0&
Private Const MF_STRING As Long = &H0&
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_CHECKED As Long = &H8&
Private Const MF_GRAYED As Long = &H1&
Private Const MF_POPUP As Long = &H10&
'Private Const MF_DEFAULT As Long = &H1000&
Private Const MF_DISABLED As Long = &H2& Or MF_GRAYED
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const TV_FIRST As Long = &H1100
Private Const TVGN_CARET As Long = &H9
Private Const TVM_GETITEMRECT As Long = (TV_FIRST + 4)
Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private WithEvents TheManifest As cManifestEx
Attribute TheManifest.VB_VarHelpID = -1
Private Enum ElementDisplayOptions
    edo_CanDelete = 256       ' element can be deleted
    edo_AttrsFixed = 128      ' attributes are hidden in tree view (can be displayed via shift+right click)
    edo_Export = 1024         ' include in the manifest
End Enum
Private Const DefaultURL As String = "https://msdn.microsoft.com/en-us/library/windows/desktop/aa374191(v=vs.85).aspx"


Private Sub Form_Load()
    
    Me.Caption = "Application Manifest Creator v" & App.Major & "." & App.Minor & "." & App.Revision
    
    ' enable DPI awareness for image list
    Select Case 1440 \ Screen.TwipsPerPixelX
    Case Is < 144: Set tvItems.ImageList = ImageList16  ' 16x16
    Case Is < 192: Set tvItems.ImageList = ImageList24  ' 24x24
    Case Else: Set tvItems.ImageList = ImageList32      ' 32x32
    End Select
    
    Set TheManifest = New cManifestEx
    TheManifest.CreateManifest Nothing, 0&
    
    ' subclass the form to restrict minimum height and maximum width; then initially set form size & position
    SetSizeRestrictions Me.hWnd, ScaleX(Me.Width, vbTwips, vbPixels), ScaleY(txtHelp.Height * 4!, vbTwips, vbPixels)
    Me.Move (Screen.Width - Me.Width) / 2!, (Screen.Height - Screen.Height * 0.7!) / 2!, Me.Width, Screen.Height * 0.7!
    
End Sub

Private Sub Form_Resize()
    On Error GoTo ExitRoutine
    If Not Me.WindowState = vbMinimized Then
        Dim sngCy As Single
        sngCy = Me.ScaleHeight - txtHelp.Height - tvItems.Top - Screen.TwipsPerPixelY * 5!
        With tvItems
            ' common controls don't always scale accurately in 200 DPI; address that
            If (1440! \ Screen.TwipsPerPixelX) = (1440! / Screen.TwipsPerPixelX) Then
                .Height = sngCy
            Else
                .Move .Left + Screen.TwipsPerPixelX, .Top + Screen.TwipsPerPixelY, txtHelp.Width, sngCy
                .Move .Left - Screen.TwipsPerPixelX, .Top - Screen.TwipsPerPixelY
            End If
        End With
        txtHelp.Top = tvItems.Top + sngCy + Screen.TwipsPerPixelY
    End If
ExitRoutine:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set TheManifest = Nothing
End Sub

Private Sub TheManifest_Begin(lParam As Long)
    
    ' called after cManifestEx successfully loaded a manifest xml & before sending that information
    
    If lParam < 1& Then
        tvItems.Nodes.Clear
        With tvItems.Nodes.Add(, , "rtApp", "Application", "rtApp")
            .Expanded = True
            .Sorted = True
        End With
        tvItems.Nodes.Add(, , "rtDep", "Dependencies (Windows XP and above)", "rtDep").Expanded = True
        tvItems.Nodes.Add "rtDep", tvwChild, "rtFile", "File(s)", "rsFile"
        With tvItems.Nodes.Add(, , "rtVista", "Windows Vista and above", "pc")
            .Expanded = True
            .Sorted = True
        End With
        With tvItems.Nodes.Add("rtVista", tvwChild, "rtWS", "Windows Settings", "rtWS")
            .Expanded = True
            .Sorted = True
        End With
        tvItems.Nodes.Add(, , "rtWin7", "Windows 7 and above", "pc").Expanded = True
        tvItems.Nodes.Add("rtWin7", tvwChild, "rtOS", "O/S Compatibility", "rtOS").Expanded = True
    End If
    
End Sub

Private Sub TheManifest_ElementAdded(Element As cManifestEntryEx, lParam As Long)

    Dim sCaption As String, sIndex As String, sName As String
    Dim a As Long, bNotSet As Boolean, tNode As ComctlLib.Node
    
    ' There are about two dozen default manifest elements created by the project, but any number
    '   of elements can be imported from an external manifest
    ' While displaying the elements, we are also attaching 'properties' to those
    '   elements for referencing, navigating, and editing as needed
    
    ' treeview node keys are rather simple:
    '   primary node uses the Key property of the passed Element
    '   attribute nodes append the primary Key with pipe & attr index, i.e., |0 |1 |2, etc
    
    ' default icon for the added element; can be changed herein
    If (Element.isActive And 1&) Then sIndex = "select" Else sIndex = "unselect"
    sName = Element.GetName()
    
    Select Case sName
    Case "noInherit", "noInheritable"
        If Element.isElementTopLevel = False Then   ' should always be a top-level node
            bNotSet = True
        Else
            tvItems.Nodes.Add "rtApp", tvwChild, Element.Key, "No " & Mid$(sName, 3) & ": True", sIndex
            If lParam <> 0& Then Element.Flags = Element.Flags Or edo_CanDelete
        End If
        
    Case "assemblyIdentity"
        If (Element.Flags And elem_Required) Then   ' Identity assembly vs dependent assembly
            Set tNode = tvItems.Nodes.Add(tvItems.Nodes("rtApp").Child.FirstSibling, tvwNext, "rtID", "Identity", "select")
            Element.isActive = Element.isActive Or 1&
        
        ElseIf Element.isElementDescendantOf("dependency", False) Then ' should be a dependent assembly
            If lParam = 0& Or Element.isActive = 0& Then
                sCaption = Element.GetValue("name", False) & ": True"
                Element.Flags = Element.Flags Or edo_AttrsFixed
            Else
                sCaption = Element.GetValue("name", False)
            End If
            Set tNode = tvItems.Nodes.Add("rtDep", tvwChild, Element.Key, sCaption, sIndex)
            tNode.Sorted = True
        Else
            bNotSet = True
        End If
        
    Case "comInterfaceExternalProxyStub"
        sCaption = sName & ": " & Element.GetValue("name")
        Set tNode = tvItems.Nodes.Add("rtDep", tvwChild, Element.Key, sCaption, sIndex)
        tNode.Sorted = True
        For a = 0& To Element.NumberAttributes - 1&
            tvItems.Nodes.Add tNode, tvwChild, Element.Key & "|" & CStr(a), Element.GetName(a) & ": " & Element.GetValue(a, False), "child"
        Next
        Element.Flags = Element.Flags Or edo_CanDelete
        tNode.Expanded = True: tNode.Selected = True
        
    Case "description"
        If Element.isElementTopLevel = False Then
            bNotSet = True
        Else
            sCaption = "Description"
            tvItems.Nodes.Add "rtApp", tvwChild, Element.Key, sCaption & ": " & Element.GetValue(), sIndex
        End If
        
    Case "supportedOS"
        If (Element.Flags("Id") And attr_ManualEntry) = 0& Then
            Select Case LCase$(Element.GetValue("Id"))
                Case "{e2011457-1546-43c5-a5fe-008deee3d3f0}": sCaption = "Windows Vista"
                Case "{35138b9a-5d96-4fbd-8e2d-a2440225f93a}": sCaption = "Windows 7"
                Case "{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}": sCaption = "Windows 8"
                Case "{1f676c76-80e1-4239-95bb-83d0f6d0da78}": sCaption = "Windows 8.1"
                Case "{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}": sCaption = "Windows 10"
            End Select
            tvItems.Nodes.Add "rtOS", tvwChild, Element.Key, sCaption & ": True", sIndex
        Else
            Set tNode = tvItems.Nodes.Add("rtOS", tvwChild, Element.Key, "Unknown O/S", sIndex)
            tvItems.Nodes.Add tNode, tvwChild, Element.Key & "|" & Element.GetAttributeIndex("Id"), "Id: " & Element.GetValue("Id", False), "child"
        End If
        Element.Flags = Element.Flags Or edo_AttrsFixed
        
    Case "requestedExecutionLevel"
        Set tNode = tvItems.Nodes.Add("rtVista", tvwChild, Element.Key, "Trust Info", "select")
        If Element.isActive Then tNode.Expanded = True
        For a = 0& To Element.NumberAttributes - 1
            tvItems.Nodes.Add tNode, tvwChild, Element.Key & "|" & CStr(a), Element.GetName(a) & ": " & Element.GetValue(a, False), "child"
        Next
        Element.isActive = Element.isActive Or 1&
    
    Case "file"
        sCaption = "File: " & Element.GetValue(a)
        Set tNode = tvItems.Nodes.Add("rtFile", tvwChild, Element.Key, sCaption, sIndex)
        If (Element.isActive And 1&) Then tNode.Parent.Expanded = True
        If lParam <> 0& Then            ' added this from the Insert menu; allow it to be deleted
            Element.Flags = Element.Flags Or edo_CanDelete
            Set tvItems.SelectedItem = tNode: tNode.Expanded = True
        End If
        For a = 0& To Element.NumberAttributes - 1
            tvItems.Nodes.Add tNode, tvwChild, Element.Key & "|" & CStr(a), Element.GetName(a) & ": " & Element.GetValue(a, False), "child"
        Next
    Case Else
        bNotSet = True
    End Select
    
    If bNotSet Then
        If InStr(1, Element.NameSpace, "WindowsSettings", vbTextCompare) Then
            sCaption = UCase$(Left$(sName, 1))
            For a = 2 To Len(sName)
                Select Case Asc(Mid$(sName, a, 1))
                Case 65 To 90: sCaption = sCaption & " " & UCase$(Mid$(sName, a, 1))
                Case Else: sCaption = sCaption & Mid$(sName, a, 1)
                End Select
            Next
            Set tNode = tvItems.Nodes.Add(pvGetWSnode(Element.NameSpace), tvwChild, Element.Key, sCaption & ": " & StrConv(Element.GetValue(, False), vbProperCase), sIndex)
        
        Else
            ' items not set above
            If Element.parentKey = vbNullString Then
                Set tNode = tvItems.Nodes.Add(pvGetNSnode(Element.NameSpace), tvwChild, Element.Key, Element.GetName(), sIndex)
            Else
                Set tNode = tvItems.Nodes.Add(Element.parentKey, tvwChild, Element.Key, Element.GetName(), sIndex)
            End If
            If (Element.Flags And elem_HasTextElement) Then
                tNode.Text = tNode.Text & ": " & Element.GetValue()
            ElseIf sName = "comInterfaceProxyStub" Then
                tNode.Text = tNode.Text & ": " & Element.GetValue("name")
            End If
            For a = 0& To Element.NumberAttributes - 1&
                tvItems.Nodes.Add tNode, tvwChild, Element.Key & "|" & CStr(a), Element.GetName(a) & ": " & Element.GetValue(a, False), "child"
            Next
            If lParam <> 0& Then            ' added this from the Insert menu; allow it to be deleted
                Element.Flags = Element.Flags Or edo_CanDelete
                If lParam > 0& Then tNode.EnsureVisible: tNode.Selected = True
            End If
        End If
    End If
    
    ' customize the properties of the assemblyIdentity element
    If sName = "assemblyIdentity" Then
        ' if attrs were hidden, then it is a default manifest assembly we don't want user's modifying
        If (Element.Flags And edo_AttrsFixed) = 0& Then
            tNode.Sorted = True
            For a = 0& To Element.NumberAttributes - 1&
                tvItems.Nodes.Add tNode, tvwChild, Element.Key & "|" & CStr(a), Element.GetName(a) & ": " & Element.GetValue(a, False), "child"
            Next
            If lParam <> 0& Then            ' added this from the Insert menu; allow it to be deleted
                Element.Flags = Element.Flags Or edo_CanDelete
                tNode.Expanded = True: Set tvItems.SelectedItem = tNode
            End If
        End If
    End If
    
End Sub

Private Sub TheManifest_Finish(lParam As Long)
    txtHelp.Text = "Right click on items for editing options. Only items that are checked will be written to a manifest. Shift + right click can unlock fixed items/attributes."
End Sub

Private Sub mnuCreate_Click(Index As Integer)
    
    If Index = 0 Then       ' default
        TheManifest.CreateManifest Nothing, 0&
    ElseIf Index = 10 Then   ' export
        frmOutput.SetManifest TheManifest
        frmOutput.Show vbModal
        Set frmOutput = Nothing
    ElseIf Index = 6 Then   ' copy to clipboard
        TheManifest.CreateManifest Clipboard, -1&
    ElseIf Index = 8 Then
        ' do nothing -- submenu
    Else ' res or external manifest file
        Dim cBrowser As UnicodeFileDialog
        Set cBrowser = New UnicodeFileDialog
        With cBrowser
            .DialogTitle = "Select Manifest Source Document"
            .Filter = "Manifest Files|*.manifest|VB Resource Files|*.res|VB Project Files|*.vbp|All Files|*.*"
            If Index > 2 Then .FilterIndex = Index - 1 Else .FilterIndex = Index
            .Flags = OFN_ENABLESIZING Or OFN_EXPLORER Or OFN_FILEMUSTEXIST
        End With
        If cBrowser.ShowOpen(Me.hWnd) = True Then
            If cBrowser.FilterIndex = 3& Then
                pvUploadVBPFile cBrowser.FileName
            Else
                TheManifest.CreateManifest cBrowser.FileName, -1&
            End If
        End If
        Set cBrowser = Nothing
    End If
    
End Sub

Private Sub mnuAppend_Click(Index As Integer)

    If Index = 1 Then   ' copy to clipboard
        TheManifest.CreateManifest Clipboard, -1&, True
    Else                ' external manifest file
        Dim cBrowser As UnicodeFileDialog
        Set cBrowser = New UnicodeFileDialog
        With cBrowser
            .DialogTitle = "Select Manifest Source Document"
            .Filter = "Manifest Files|*.manifest|All Files|*.*"
            .FilterIndex = 1
            .Flags = OFN_ENABLESIZING Or OFN_EXPLORER Or OFN_FILEMUSTEXIST
        End With
        If cBrowser.ShowOpen(Me.hWnd) = True Then
            TheManifest.CreateManifest cBrowser.FileName, -1&, True
        End If
        Set cBrowser = Nothing
    End If

End Sub

Private Sub mnuHelp_Click(Index As Integer)
    If Index < 2 Then
        Dim sURL As String
        If Index = 1 Then
            sURL = "https://msdn.microsoft.com/en-us/library/windows/desktop/aa374219(v=vs.85).aspx"
        Else
            sURL = mnuHelp(0).Tag
            If sURL = vbNullString Then sURL = DefaultURL
        End If
        On Error Resume Next
        ShellExecute Me.hWnd, "Open", sURL, vbNullString, vbNullString, vbNormalFocus
        If Err Then
            Clipboard.Clear: Clipboard.SetText sURL
            MsgBox Err.Description & vbCrLf & vbCrLf & "URL was placed on the clipboard", vbInformation + vbOKOnly, "Error"
            Err.Clear
        End If
    ElseIf Index = 3 Then
        ' fyi: this can be extracted manually from the res file by opening that file in NotePad
        Dim b() As Byte
        b() = LoadResData("SUBMAIN", "CUSTOM")
        Clipboard.Clear
        Clipboard.SetText StrConv(b(), vbUnicode)
        Erase b()
    End If
End Sub

Private Sub mnuInsert_Click(Index As Integer)
    
    Dim xmlElement As IXMLDOMElement, sPrefix As String
    
    If Index = 0 Then   ' add generic dependent assembly
        With TheManifest.xml.documentElement
            sPrefix = TheManifest.GetNameSpacePrefix((.namespaceURI))
            Set xmlElement = .selectSingleNode(sPrefix & "dependency") _
                             .cloneNode(True).selectSingleNode("*//" & sPrefix & "assemblyIdentity")
        End With
        With xmlElement
            .setAttribute "name", "Software.Title"
            .setAttribute "version", "1.0.0.0"
            .setAttribute "type", "win32"
            .setAttribute "processorArchitecture", "x86"
            .setAttribute "publicKeyToken", ""
            .setAttribute "language", ""
        End With
        TheManifest.AddItem xmlElement, 1&
        Set xmlElement = Nothing
        
    ElseIf Index = 1 Then ' clone specific dependent assembly
        Index = InStr(tvItems.SelectedItem.Key, "|")
        On Error Resume Next
        If Index = 0 Then
            If tvItems.SelectedItem.Parent.Key = "rtDep" Then
                Set xmlElement = TheManifest.Item(tvItems.SelectedItem.Key).ManifestElement
            End If
        ElseIf tvItems.SelectedItem.Parent.Parent.Key = "rtDep" Then
            Set xmlElement = TheManifest.Item(Left$(tvItems.SelectedItem.Key, Index - 1)).ManifestElement
        End If
        If Err Then Err.Clear
        On Error GoTo 0
        If xmlElement Is Nothing Then
            MsgBox "First select a dependent assembly entry to clone", vbInformation + vbOKOnly, "No Action Taken"
        Else
            sPrefix = TheManifest.GetNameSpacePrefix((xmlElement.namespaceURI))
            Set xmlElement = xmlElement.selectSingleNode("ancestor::" & sPrefix & "dependency").cloneNode(True)
            TheManifest.AddItem xmlElement, 1&
            Set xmlElement = Nothing
       End If
        
    ElseIf Index = 3 Then  ' add file element
        With TheManifest.xml.documentElement
            sPrefix = TheManifest.GetNameSpacePrefix((.namespaceURI))
            Set xmlElement = .selectSingleNode(sPrefix & "file").cloneNode(False)
            xmlElement.setAttribute "name", "FileName.ext"
        End With
        TheManifest.AddItem xmlElement, Index + 0&
        Set xmlElement = Nothing
    
    ElseIf Index = 5 Then ' nonInheritable
        With TheManifest.xml.documentElement
            sPrefix = TheManifest.GetNameSpacePrefix((.namespaceURI))
            Set xmlElement = .selectSingleNode(sPrefix & "noInheritable")
            If xmlElement Is Nothing Then
                Set xmlElement = .insertBefore(.appendChild(.ownerDocument.createNode(NODE_ELEMENT, sPrefix & "noInheritable", .namespaceURI)), .firstChild)
                TheManifest.AddSubItem xmlElement, , 1&
            Else
                MsgBox "Only one noInheritable element can exist in the manifest. Applies only to assembly-manifests", vbInformation + vbOKOnly, "No Action Taken"
            End If
        End With
        
    ElseIf Index = 6 Then
        With TheManifest.xml.documentElement
            sPrefix = TheManifest.GetNameSpacePrefix((.namespaceURI))
            Set xmlElement = .ownerDocument.createNode(NODE_ELEMENT, sPrefix & "comInterfaceExternalProxyStub", .namespaceURI)
        End With
        TheManifest.AddItem xmlElement, 1&
        
    ElseIf Index = 7 Then
        With TheManifest.xml.documentElement
            sPrefix = TheManifest.GetNameSpacePrefix((.namespaceURI))
            Set xmlElement = .ownerDocument.createNode(NODE_ELEMENT, sPrefix & "windowClass", .namespaceURI)
        End With
        xmlElement.Text = "Class Name"
        TheManifest.AddItem xmlElement, 1&
        
        
    ElseIf Index = 8 Then
        Dim Element As cManifestEntryEx
        On Error Resume Next
        Index = InStr(tvItems.SelectedItem.Key, "|")
        If Index Then
            Set Element = TheManifest.Item(Left$(tvItems.SelectedItem.Key, Index - 1&))
        Else
            Set Element = TheManifest.Item(tvItems.SelectedItem.Key)
        End If
        On Error GoTo 0
        If Not Element Is Nothing Then
            If Not Element.GetName() = "comClass" Then Set Element = Nothing
        End If
        If Element Is Nothing Then
            MsgBox "First select a comClass element, sub-element of File.", vbInformation + vbOKOnly, "No Action Taken"
        Else
            Set xmlElement = Element.ManifestElement
            With TheManifest.xml.documentElement
                sPrefix = TheManifest.GetNameSpacePrefix((.namespaceURI))
                Set xmlElement = xmlElement.appendChild(.ownerDocument.createNode(NODE_ELEMENT, sPrefix & "progid", .namespaceURI))
            End With
            TheManifest.AddSubItem xmlElement, Element.Key, 1&
        End If
    
    End If
End Sub

Private Sub mnuAssmbly_Click(Index As Integer)

    Dim Element As cManifestEntryEx, a As Long
    Dim sPrefix As String, xmlElement As IXMLDOMElement
    
    On Error Resume Next
    a = InStr(tvItems.SelectedItem.Key, "|")
    If a Then
        Set Element = TheManifest.Item(Left$(tvItems.SelectedItem.Key, a - 1&))
    Else
        Set Element = TheManifest.Item(tvItems.SelectedItem.Key)
    End If
    On Error GoTo 0
    If Not Element Is Nothing Then
        If Element.GetName() = "file" Then
            If Element.ManifestElement.selectSingleNode("*") Is Nothing Then
                Set Element = Nothing   ' don't append to the App-Manifest <file> elements
            End If
        Else
            If Element.parentKey = vbNullString Then
                Set Element = Nothing
            Else
                Do Until Element.parentKey = vbNullString
                    Set Element = TheManifest.Item(Element.parentKey)
                Loop
            End If
        End If
    End If
    
    sPrefix = TheManifest.GetNameSpacePrefix((TheManifest.xml.documentElement.namespaceURI))
    If Element Is Nothing Then
        With TheManifest.xml.documentElement
            Set xmlElement = .ownerDocument.createNode(NODE_ELEMENT, sPrefix & "file", .namespaceURI)
            xmlElement.setAttribute "name", "FileName.ext"
        End With
    Else
        Set xmlElement = Element.ManifestElement
    End If
    
    Select Case Index
    Case 0: Set xmlElement = xmlElement.appendChild(xmlElement.ownerDocument.createNode(NODE_ELEMENT, sPrefix & "comClass", xmlElement.namespaceURI))
    Case 1: Set xmlElement = xmlElement.appendChild(xmlElement.ownerDocument.createNode(NODE_ELEMENT, sPrefix & "typelib", xmlElement.namespaceURI))
    Case 2: Set xmlElement = xmlElement.appendChild(xmlElement.ownerDocument.createNode(NODE_ELEMENT, sPrefix & "comInterfaceProxyStub", xmlElement.namespaceURI))
    Case 3:
        Set xmlElement = xmlElement.appendChild(xmlElement.ownerDocument.createNode(NODE_ELEMENT, sPrefix & "windowClass", xmlElement.namespaceURI))
        xmlElement.Text = "Class Name"
    End Select
    If xmlElement.parentNode.parentNode Is Nothing Then
        TheManifest.AddItem xmlElement, 1&
    Else
        TheManifest.AddSubItem xmlElement, Element.Key, 1&
    End If
    Set Element = Nothing
    Set xmlElement = Nothing

End Sub

Private Sub MenuAction_Timer()

    ' note: timer only used to allow mouse release & node selection to finish before showing menu
    MenuAction.Enabled = False
    
    ' editing a selected item
    Dim sValue As String, sCaption As String, sKey As String
    Dim hMenu As Long, lSubMenu As Long, lReturn As Long, lMask As Long
    Dim lIndex As Long, vAttr As Variant, lFlags As Long
    Dim f As frmExtract, ptMouse As POINTAPI
    Dim mElement As cManifestEntryEx, tDoc As DOMDocument60
    
    GetCursorPos ptMouse
    sKey = tvItems.SelectedItem.Key
    Select Case tvItems.SelectedItem.Image
    Case "select", "unselect", "child"
        If sKey = "rtID" Then sKey = tvItems.SelectedItem.Child.FirstSibling.Key
        lIndex = InStr(sKey, "|")
        If lIndex = 0& Then
            Set mElement = TheManifest.Item(sKey)
        Else
            If Not tvItems.SelectedItem.Key = "rtID" Then
                vAttr = CLng(Mid$(sKey, lIndex + 1&))
            End If
            Set mElement = TheManifest.Item(Left$(sKey, lIndex - 1&))
        End If
        lFlags = mElement.Flags(vAttr)  ' ElementPropEnum & AttrPropEnum have same values for what we need
    Case Else
        ' main tree node/branch; not a editable item
    End Select
    
    ' lIndex contains these settings; also is order of submenu item creation
    '   1=edit options (if any)
    '   2=element text node (does not apply to attrs)
    '   4=select/unselect node
    '   8=can delete    << this is a form-level only flag; not a cManifestEntryEx enum flag
    '  16=select/unselect all nodes
    '  32=expand/collapse
    '  64=can display xml segment
    ' 128=can display qualified name
    
    If mElement Is Nothing Then ' not on an editable node
        lIndex = 48&    ' can collapse/expand or select/unselect all
    Else
        lIndex = 192&            ' can display qualified name, xml extract
        If (lFlags And elem_Fixed) = 0& Then lIndex = lIndex Or 1& ' can be edited
        If (lFlags And edo_CanDelete) Then lIndex = lIndex Or 8& ' can be deleted
        If IsEmpty(vAttr) Then  ' non-attr options
            If (lFlags And elem_HasTextElement) Then lIndex = lIndex Or 2& ' has a text element
            If (lFlags And elem_Required) = 0& Then lIndex = lIndex Or 4& ' can select/unselect
            If tvItems.SelectedItem.Children Then lIndex = lIndex Or 32& ' can expand/collapse
        End If
        ' allow shift+right click to unlock locked values
        If Mid$(tvItems.Tag, 2, 1) = CStr(vbShiftMask) Then
            If IsEmpty(vAttr) = False Or (lIndex And 2&) = 2& Then
                lFlags = lFlags Or elem_ManualEntry
                lIndex = lIndex Or 1&
            End If
        End If
    End If
    
    hMenu = CreatePopupMenu
    
    '/// create any display-friendly submenu items to select from when changing a value
    ' menu items 95-999 are reserved for customizing values. Id 99 & 98 are for manual entry & erasing, respectively
    If Not mElement Is Nothing Then
        If (lFlags And elem_HasValueList) Then  ' elem_Fixed never applies if a value list exists
            lSubMenu = CreatePopupMenu
            AppendMenu hMenu, MF_POPUP, lSubMenu, ByVal "Value"
            For lMask = 0& To mElement.GetValueListCount(vAttr) - 1&
                AppendMenu lSubMenu, MF_STRING, 100 + lMask, mElement.GetValueListItem(lMask, vAttr, False)
            Next
            If (lFlags And elem_ManualEntry) Then AppendMenu lSubMenu, MF_STRING, 99, "New Value..."
            If (lFlags And elem_CanBeBlank) Then AppendMenu hMenu, MF_STRING, 98, "Erase Value"
        ElseIf (lFlags And 2) = 2 Or (lFlags And edo_AttrsFixed) = 0 Then
            If (lFlags And edo_AttrsFixed) = 0 Then
                If (lIndex And 1) = 0& Then ' can't be edited?
                    If (lIndex And 2) Then sCaption = "Set Text Value" Else sCaption = "Change Value"
                    AppendMenu hMenu, MF_STRING Or MF_DISABLED, 99, sCaption
                Else
                    If (lIndex And 2) Then sCaption = "Set Text Value" Else sCaption = "Change Value"
                    AppendMenu hMenu, MF_STRING, 99, sCaption
                    If (lFlags And elem_CanBeBlank) Then AppendMenu hMenu, MF_STRING, 98, "Erase Value"
                End If
            End If
        End If
    End If
    
    '/// create the other submenu items, as applicable
    '   Use IDs less than 95 if no action is required on menu selection; else use 1000+
    lMask = 4&
    Do Until lMask > lIndex
        Select Case (lIndex And lMask)
        Case 4
            If (mElement.isActive And 1) Then sValue = "Unselect from Manifest" Else sValue = "Select to the Manifiest"
            AppendMenu hMenu, MF_STRING, 1000, sValue
            If (lFlags And edo_AttrsFixed) Then
                ' allow shift right click to unlock attributes
                If Mid$(tvItems.Tag, 2, 1) = CStr(vbShiftMask) Then AppendMenu hMenu, MF_STRING, 1040, "Show Attributes"
            End If
        Case 8
            AppendMenu hMenu, MF_SEPARATOR Or MF_DISABLED, 75, ByVal 0&
            AppendMenu hMenu, MF_STRING, 1030, "Delete from Tree"
        Case 16
            AppendMenu hMenu, MF_STRING, 1010, "Select All"
            AppendMenu hMenu, MF_STRING, 1011, "Unselect All"
        Case 32
            If (lIndex And 16&) Then AppendMenu hMenu, MF_SEPARATOR Or MF_DISABLED, 75, ByVal 0&
            If tvItems.SelectedItem.Expanded Then sCaption = "Collapse" Else sCaption = "Expand"
            AppendMenu hMenu, MF_STRING, 1020, sCaption
        Case 64
            AppendMenu hMenu, MF_SEPARATOR Or MF_DISABLED, 75, ByVal 0&
            lSubMenu = CreatePopupMenu
            AppendMenu hMenu, MF_POPUP, lSubMenu, ByVal "Qualified Name"
            AppendMenu lSubMenu, MF_STRING, 90, "BaseName: " & mElement.GetName(vAttr)
            sCaption = mElement.NameSpace(vAttr)
            If sCaption = vbNullString Then sCaption = "{none assigned}"
            AppendMenu lSubMenu, MF_STRING, 91, "NameSpace: " & sCaption
        Case 128
            AppendMenu hMenu, MF_SEPARATOR Or MF_DISABLED, 75, ByVal 0&
            AppendMenu hMenu, MF_STRING, edo_Export, "Show me the XML"
        End Select
        lMask = lMask + lMask
    Loop
    lReturn = TrackPopupMenu(hMenu, TPM_LEFTALIGN Or TPM_NOANIMATION Or TPM_RETURNCMD Or TPM_TOPALIGN, ptMouse.x, ptMouse.y, vbDefault, tvItems.hWnd, 0&)
    DestroyMenu hMenu

    Select Case lReturn
    Case Is < 95   ' do nothing
    Case 98, 99     ' erase/manual edit
        If lReturn = 98& Then       ' erase; set to null string
            sValue = vbNullString
        Else
            sCaption = vbNullString
            ' allow shift right click to unlocked fixed value list
            If (mElement.Flags(vAttr) And (elem_HasValueList Or elem_ManualEntry)) = (elem_HasValueList Or elem_ManualEntry) Then
                ' provide a list of known good values along with the prompt
                ' currently: dpiAwareness & comClass miscXXX elements are only such elements
                '   where it has a list & can accept non-list values
                With frmMultiAttrValue
                    .SetValueList mElement, vAttr
                    .lblPrompt.Caption = mElement.GetName(vAttr) & vbCrLf & "Select one or more items:"
                    .Show vbModal, Me
                    sValue = .Value
                End With
                Unload frmMultiAttrValue: Set frmMultiAttrValue = Nothing
            Else
                Select Case mElement.GetName()  ' add any specific help/prompt for the element/attribute being manually edited
                    Case "assemblyIdentity"
                        Select Case mElement.GetName(vAttr)
                        Case "name"
                            sCaption = "Use the following format for the name: Organization.Division.Name. For example Microsoft.Windows.mysampleApp"
                        Case "language"
                            If (mElement.Flags() And elem_Required) Then  ' asterisk authorized for ID element only for assembly manifests
                                sCaption = "Enter the DHTML language code or leave blank"
                            Else
                                sCaption = "Enter the DHTML language code or asterisk (for neutral language)"
                            End If
                        Case "version"
                            sCaption = "Use the four-part version format: major.minor.build.revision. Each of the parts separated by periods can be 0-65535 inclusive"
                        Case "publicKeyToken"
                            sCaption = "Enter 16-character hexadecimal string representing the last 8 bytes of the SHA-1 hash of the public key under which the application or assembly is signed."
                        End Select
                    Case "file"
                        Select Case mElement.GetName(vAttr)
                        Case "name": sCaption = "Enter the file name and extension"
                        Case "hashalg": sCaption = "Enter algorithm used to create a hash of the file, i.e., SHA1"
                        Case "hash": sCaption = "Enter a hexadecimal string of length depending on the hash algorithm"
                        End Select
                    Case Else
                        If Not sCaption = vbNullString Then
                            sCaption = "Enter a manifest-valid value or one of the following:" & vbCrLf & sCaption
                        End If
                End Select
                If sCaption = vbNullString Then sCaption = "Enter a manifest-valid value. Select the help menu if needed."
                sValue = InputBox(sCaption, "Manual Entry", mElement.GetValue(vAttr))
            End If
            If StrPtr(sValue) = 0 Then GoTo ExitRoutine
            
            If (lFlags And elem_CanBeBlank) = 0& Then
                If Trim$(sValue) = vbNullString Then
                    MsgBox "Blank entry is invalid", vbInformation + vbOKOnly, "No Action Taken"
                    GoTo ExitRoutine
                End If
            End If
        End If
        mElement.SetValue sValue, vAttr: sValue = mElement.GetValue(vAttr, False)
        lIndex = InStr(tvItems.SelectedItem.Text, ":")
        tvItems.SelectedItem.Text = Left$(tvItems.SelectedItem.Text, lIndex) & " " & sValue
        ' if shift+right click changed the value, remove the elem_Fixed property for subsequent edits
        If (lFlags And elem_Fixed) Then mElement.Flags(vAttr) = mElement.Flags(vAttr) Xor elem_Fixed
        
        ' pretty up elements where we display a key attr value in its tree node caption
        If Not (IsEmpty(vAttr) = True Or tvItems.SelectedItem.Key = "rtID") Then
            If mElement.GetName(vAttr) = "name" Then
                Select Case mElement.GetName()
                    Case "file", "comInterfaceExternalProxyStub", "comInterfaceProxyStub"
                        sCaption = tvItems.SelectedItem.Parent.Text
                        lIndex = InStr(sCaption, ":")
                        tvItems.SelectedItem.Parent.Text = Left$(sCaption, lIndex) & " " & sValue
                    Case "assemblyIdentity"
                        tvItems.SelectedItem.Parent.Text = sValue
                End Select
            End If
        End If
    
    Case Is < 1000  ' value list entries
        mElement.SetValue mElement.GetValueListItem(lReturn - 100&, vAttr), vAttr
        sValue = mElement.GetValue(vAttr, False)
        lIndex = InStr(tvItems.SelectedItem.Text, ":")
        tvItems.SelectedItem.Text = Left$(tvItems.SelectedItem.Text, lIndex) & " " & sValue
        
    Case 1000       ' select/unselect
        If tvItems.SelectedItem.Image = "select" Then
            tvItems.SelectedItem.Image = "unselect"
        Else
            tvItems.SelectedItem.Image = "select"
        End If
        mElement.isActive = mElement.isActive Xor 1&
        
    Case 1010, 1011 ' select/unselect all; calls recursive routine
        pvSelectSubItems tvItems.SelectedItem.Child.FirstSibling, (lReturn = 1010), False
    
    Case 1020       ' expand/collapse
        tvItems.SelectedItem.Expanded = Not tvItems.SelectedItem.Expanded
        
    Case 1030       ' delete tree item
        lIndex = 0
        With tvItems.SelectedItem.Parent
            If Left$(.Key, 3) = "ns." Then
                If .Image = "pc" Then lIndex = .Index
            End If
        End With
        If MsgBox("Are you sure you want to permanently remove this from the tree?", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
            pvSelectSubItems tvItems.SelectedItem, False, True
            If lIndex > 0 Then
                If tvItems.Nodes(lIndex).Children = 0 Then tvItems.Nodes.Remove lIndex
            End If
        End If
        
    Case 1040       ' unhide the attrs for default manifest dependencies (common controls & GDI+)
        With tvItems.SelectedItem
            lIndex = InStr(.Text, ":")
            If lIndex Then .Text = Left$(.Text, lIndex - 1&)
            For lIndex = 0& To mElement.NumberAttributes - 1&
                tvItems.Nodes.Add .Key, tvwChild, .Key & "|" & CStr(lIndex), mElement.GetName(lIndex) & ": " & mElement.GetValue(lIndex, False), "child"
            Next
            .Expanded = True
        End With
        mElement.Flags = mElement.Flags Xor edo_AttrsFixed ' remove the flag
        Set tvItems.SelectedItem = tvItems.SelectedItem.Child.FirstSibling
        
    Case edo_Export       ' extract item's xml
        mElement.isActive = mElement.isActive Or edo_Export
        Set tDoc = TheManifest.ExportXML(edo_Export, True, True, False, False)
        Set f = New frmExtract
        f.txtExtract.Text = tDoc.xml
        mElement.isActive = mElement.isActive Xor edo_Export
        Set tDoc = Nothing
        f.Show , Me
    
    End Select
    
ExitRoutine:
    Set mElement = Nothing
    
End Sub

Private Sub tvItems_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    tvItems.Tag = CStr(Button And 7) & CStr(Shift And vbShiftMask) & CStr(x)
End Sub

Private Sub tvItems_NodeClick(ByVal Node As ComctlLib.Node)

    Call pvSetHelpText
    
    If Left$(tvItems.Tag, 1) = "1" Then        ' left button click vs. right button click
        If Node.Key = "rtID" Then Exit Sub      ' can't unselect this node
        If tvItems.ImageList.ListImages(Node.Image).Index > 2& Then Exit Sub
        
        Dim tRect As RECT, iIndex As Long, sKey As String
        ' version 5 of the common controls doesn't have a "Checked" property; do it the hard way
        tRect.Left = SendMessage(tvItems.hWnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
        If Not tRect.Left = 0& Then
            If SendMessage(tvItems.hWnd, TVM_GETITEMRECT, 1&, tRect) Then
                tRect.Right = ScaleX(CLng(Mid$(tvItems.Tag, 3)), Me.ScaleMode, vbPixels)
                If tRect.Right < tRect.Left Then
                    sKey = Node.Key: iIndex = InStr(sKey, "|")
                    If iIndex > 0 Then sKey = Left$(sKey, iIndex - 1)
                    If Node.Image = "select" Then Node.Image = "unselect" Else Node.Image = "select"
                    TheManifest.Item(sKey).isActive = TheManifest.Item(sKey).isActive Xor 1&
                End If
            End If
        End If
    ElseIf Left$(tvItems.Tag, 1) = "2" Then ' right click node
        MenuAction.Enabled = True
    End If
End Sub

Private Function pvGetNSnode(ns As String) As ComctlLib.Node

    ' only called if a namespace exists that is not expected

    Dim tNode As ComctlLib.Node
    Dim sPrefix As String
    
    On Error Resume Next
    sPrefix = TheManifest.GetNameSpacePrefix((ns))
    If Right$(sPrefix, 1) = ":" Then
        sPrefix = "ns." & Left$(sPrefix, Len(sPrefix) - 1&)
    Else
        sPrefix = "ns."
    End If
    Set tNode = tvItems.Nodes(sPrefix)
    If Err Then
        Err.Clear
        If ns = vbNullString Then ns = "{null namespace}"
        Set tNode = tvItems.Nodes.Add(, , sPrefix, "NameSpace " & ns, "pc")
        tNode.Expanded = True: tNode.Sorted = True
    End If
    Set pvGetNSnode = tNode
        
End Function

Private Function pvGetWSnode(ns As String) As ComctlLib.Node

    ' called to dynamically create WindowSettings nodes
    
    Dim tNode As ComctlLib.Node
    Dim sCaption As String
    Dim i As Long, j As Long
    
    On Error Resume Next    ' parse out the year from the URL
    i = InStr(1, ns, "/WindowsSettings", vbTextCompare)
    If i = 0 Then
        sCaption = ns
    Else
        j = InStrRev(ns, "/", i - 1)
        If j = 0 Then sCaption = ns Else sCaption = Mid$(ns, j + 1, i - j - 1)
    End If
    Set tNode = tvItems.Nodes("ws." & sCaption)
    If Err Then
        Err.Clear
        Set tNode = tvItems.Nodes.Add("rtWS", tvwChild, "ws." & sCaption, sCaption & " NameSpace", "rtWS")
        tNode.Expanded = True: tNode.Sorted = True
    End If
    On Error GoTo 0
    Set pvGetWSnode = tNode

End Function

Private Sub pvSetHelpText()

    ' Provide some meaningful help for the treeview node clicked on
    ' If there is a specific URL the user can go to, relating to the clicked node, provide that too
    '   If no URL is available, leave blank and the DefaultURL constant will be used

    Dim sURL As String, sName As String, a As Long, sKey As String
    
    sURL = vbNullString
    Select Case tvItems.SelectedItem.Key
    Case "rtApp": txtHelp.Text = "Application related manifest elements"
    Case "rtID": txtHelp.Text = "The 'identity' element that describes your application"
    Case "rtDep": txtHelp.Text = "Dependent assemblies/DLLs to be used with your application"
    Case "rtVista": txtHelp.Text = "Options that apply from Windows Vista and newer"
    Case "rtWin7": txtHelp.Text = "Options that apply from Windows 7 and newer"
    Case "rtWin10": txtHelp.Text = "Options that apply from Windows 8 and newer"
    Case "rtOS", "rtComp": txtHelp.Text = "Compatibility options introduced with Windows 7"
        sURL = "http://msdn.microsoft.com/en-us/library/windows/desktop/hh848036%28v=vs.85%29.aspx"
    Case Else
        If Left$(tvItems.SelectedItem.Key, 3) = "ns." Then
            txtHelp.Text = "Elements that are assigned to the selected name space"
        ElseIf Left$(tvItems.SelectedItem.Key, 3) = "ws." Or tvItems.SelectedItem.Key = "rtWS" Then
            txtHelp.Text = "Various WindowSettings namespace settings"
        ElseIf tvItems.SelectedItem.Key = "rtFile" Then
            txtHelp.Text = "Specifies files that are private to the application"
        Else
            a = InStr(tvItems.SelectedItem.Key, "|")
            If a = 0& Then sKey = tvItems.SelectedItem.Key Else sKey = Left$(tvItems.SelectedItem.Key, a - 1)
            sName = TheManifest.Item(sKey).GetName()
            
            Select Case sName
            Case "noInherit": txtHelp.Text = "Include this element in an application manifest to set the activation contexts generated from the manifest with the 'no inherit' flag. When this flag is not set in an activation context, and the activation context is active, it is inherited by new threads in the same process, windows, window procedures, and Asynchronous Procedure Calls. Setting this flag prevents the new object from inheriting the active context"
            Case "assemblyIdentity"
                Select Case LCase$(TheManifest.Item(sKey).GetValue("name"))
                Case "microsoft.windows.common-controls": txtHelp.Text = "Adding the common controls to your manifests requires you to include the Sub Main to your project and start your project with Sub Main(). Click the help menu above to copy a typical Sub Main() to memory."
                    sURL = "https://msdn.microsoft.com/en-us/library/windows/desktop/bb773175(v=vs.85).aspx"
                Case "microsoft.windows.gdiplus": txtHelp.Text = "Allows application to access and use the updated GDI+ v1.1 library. Do not include if application will be used in Windows XP."
                Case Else: txtHelp.Text = "App Title, Version and Type are required. All others are optional"
                End Select
            Case "description": txtHelp.Text = "Optional element that describes your application"
            Case "requestedExecutionLevel": txtHelp.Text = "To understand why you would want to add a security section to your manifest, click the help menu above."
                sURL = "http://msdn.microsoft.com/en-us/library/bb756929.aspx"
            Case "supportedOS": txtHelp.Text = "To view why you would want to add versioning to your manifest click the help menu above. If selecting any, select all that apply."
                sURL = "http://msdn.microsoft.com/en-us/library/windows/desktop/hh848036%28v=vs.85%29.aspx"
            Case "autoElevate": txtHelp.Text = "You should research this before adding to a manifest. May be applicable only to embedded manifests and digitally signed applications"
            Case "disableTheming": txtHelp.Text = "Specifies whether giving UI elements a theme is disabled"
            Case "disableWindowFiltering": txtHelp.Text = "Specifies whether to disable window filtering. TRUE disables window filtering so you can enumerate immersive windows from the desktop"
            Case "printerDriverIsolation": txtHelp.Text = "Printer driver isolation improves the reliability of the Windows print service by enabling printer drivers to run in processes that are separate from the process in which the print spooler runs"
            Case "dpiAware": txtHelp.Text = "DPI-Awareness prevents DPI virtualization (stretching/scaling) when your application is displayed on a larger DPI screens."
                sURL = "http://msdn.microsoft.com/en-us/library/windows/desktop/dd756693%28v=vs.85%29.aspx"
            Case "dpiAwareness": txtHelp.Text = "Applies to Win10 v1607. This setting overrides the older <dpiAware> setting. DPI-Awareness prevents DPI virtualization (stretching/scaling) when your application is displayed on a larger DPI screens."
                sURL = "http://msdn.microsoft.com/en-us/library/windows/desktop/dd756693%28v=vs.85%29.aspx"
            Case "longPathAware": txtHelp.Text = "Applies to Win10 v1607. Supposedly allows application to use wide/unicode file/directory APIs without the 260 character folder length restriction."
                sURL = "https://msdn.microsoft.com/en-us/library/windows/desktop/aa365247(v=vs.85).aspx"
            Case "gdiScaling"
                txtHelp.Text = "Applies to Win10 v1703. The GDI (graphics device interface) framework can apply DPI scaling to primitives and text on a per-monitor basis without updates to the application itself. This can be useful for GDI applications no longer being actively updated."
            Case "file": txtHelp.Text = "Specifies files that are private to the application"
            
            ' assembly manifests
            Case "noInheritable": txtHelp.Text = "The noInheritable element is required in the assembly manifest if the assembly is used by any application manifests that include the noInherit element"
            Case "comInterfaceExternalProxyStub": txtHelp.Text = "A subelement of an assembly element and is used for automation interfaces. For example, IDispatch and its derived interfaces"
            Case "comInterfaceProxyStub": txtHelp.Text = "If a file in the assembly implements a proxy stub, the corresponding file tag must include a comInterfaceProxyStub subelement having attributes that are identical to a comInterfaceProxyStub element. Marshaling interfaces between processes and threads may not work as expected if you omit some of the comInterfaceProxyStub dependencies for your component."
            Case "windowClass": txtHelp.Text = "The name of a windows class that is to be versioned"
            
            Case Else
                If Left$(tvItems.SelectedItem.Parent.Key, 3) = "ws." Then
                    txtHelp.Text = "You should research this before adding to a manifest. There isn't a lot of information available when this application was created"
                Else
                    txtHelp.Text = "Assembly-Manifest or unrecognized element or namespace. Could be a custom element, misspelled or case-sensitive mismatch. You should research this before adding to a manifest."
                End If
            End Select
        End If
    End Select
    mnuHelp(0).Tag = sURL
End Sub

Private Sub pvUploadVBPFile(FileName As String)

    ' routine fills in the app title, description and version from a .vbp file

    Dim hFile As Long, sLines() As String, bData() As Byte
    Dim lLine As Long, iPos As Long, sMsg As String, sKey As String
    Dim sVersion As String, bSubMain As Boolean, bResFile As Boolean
    Dim sName As String, sDescription As String
    Dim tNode As ComctlLib.Node
    
    hFile = CreateTheFile(FileName, True, IsUnicodeSystem())
    If hFile = INVALID_HANDLE_VALUE Or hFile = 0& Then
        MsgBox "Failed to open the manifest file. Ensure proper permissions exist", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    lLine = GetFileSize(hFile, ByVal 0&)
    If lLine < 1& Then
        CloseHandle hFile
        MsgBox "File is not in expected format", vbExclamation + vbOKOnly, "Error"
    Else
        ReDim bData(0 To lLine - 1)
        ReadFile hFile, bData(0), lLine, lLine, ByVal 0&
        CloseHandle hFile
        If lLine > UBound(bData) Then
            sLines() = Split(StrConv(bData, vbUnicode), vbCrLf)
            Erase bData()
            sVersion = "M.m.0.R"
            For lLine = 0 To UBound(sLines)
                iPos = InStr(sLines(lLine), "=")
                If iPos Then
                    Select Case LCase$(Left$(sLines(lLine), iPos - 1))
                    Case "name"
                        If sName = vbNullString Then sName = Replace$(Mid$(sLines(lLine), iPos + 1), Chr$(34), vbNullString)
                    Case "title": sName = Replace$(Mid$(sLines(lLine), iPos + 1), Chr$(34), vbNullString)
                    Case "majorver": sVersion = Replace$(sVersion, "M", Mid$(sLines(lLine), iPos + 1))
                    Case "minorver": sVersion = Replace$(sVersion, "m", Mid$(sLines(lLine), iPos + 1))
                    Case "revisionver": sVersion = Replace$(sVersion, "R", Mid$(sLines(lLine), iPos + 1))
                    Case "description": sDescription = Replace$(Mid$(sLines(lLine), iPos + 1), Chr$(34), vbNullString)
                    Case "resfile32": bResFile = True
                    Case "startup"
                        If StrComp(Mid$(sLines(lLine), iPos + 1), """Sub Main""", vbTextCompare) = 0 Then bSubMain = True
                    End Select
                End If
            Next
            Erase sLines()
            If InStr(sVersion, "M") Then sVersion = Replace$(sVersion, "M", "1")
            If InStr(sVersion, "m") Then sVersion = Replace$(sVersion, "m", "0")
            If InStr(sVersion, "R") Then sVersion = Replace$(sVersion, "R", "0")
            
            TheManifest.CreateManifest Nothing, 0&
            sMsg = tvItems.Nodes("rtID").Child.FirstSibling.Key
            With TheManifest.Item(Left$(sMsg, InStr(sMsg, "|") - 1))
                If Not sName = vbNullString Then
                    lLine = .GetAttributeIndex("name")
                    sMsg = tvItems.Nodes(.Key & "|" & CStr(lLine)).Text
                    .SetValue sName, lLine
                    iPos = InStr(sMsg, ":")
                    If iPos = 0 Then
                        tvItems.Nodes(.Key & "|" & CStr(lLine)).Text = sMsg & ": " & sName
                    Else
                        tvItems.Nodes(.Key & "|" & CStr(lLine)).Text = Left$(sMsg, iPos + 1) & sName
                    End If
                End If
                If Not sVersion = vbNullString Then
                    lLine = .GetAttributeIndex("version")
                    If lLine > -1& Then
                        sMsg = tvItems.Nodes(.Key & "|" & CStr(lLine)).Text
                        .SetValue sVersion, lLine
                        iPos = InStr(sMsg, ":")
                        If iPos = 0 Then
                            tvItems.Nodes(.Key & "|" & CStr(lLine)).Text = sMsg & ": " & sVersion
                        Else
                            tvItems.Nodes(.Key & "|" & CStr(lLine)).Text = Left$(sMsg, iPos + 1) & sVersion
                        End If
                    End If
                End If
            End With
            If Not sDescription = vbNullString Then
                Set tNode = tvItems.Nodes("rtApp").Child.FirstSibling
                Do Until tNode Is Nothing
                    If Left$(tNode.Text, 11) = "Description" Then
                        sMsg = tNode.Text
                        TheManifest.Item(tNode.Key).SetValue sDescription
                        iPos = InStr(sMsg, ":")
                        If iPos = 0 Then
                            tNode.Text = sMsg & ": " & sDescription
                        Else
                            tNode.Text = Left$(sMsg, iPos + 1) & sDescription
                        End If
                        Exit Do
                    End If
                    Set tNode = tNode.Next
                Loop
            End If
            
            sMsg = "The Identity element has been populated with the data from your VBP file." & vbCrLf & vbCrLf & "The following is also provided..." & vbCrLf
            If bResFile Then
                sMsg = sMsg & "A resource file is referenced in that project. You may want to review/replace its manifest if it has any" & vbCrLf
            Else
                sMsg = sMsg & "No resource file is referenced in that project" & vbCrLf
            End If
            If bSubMain = False Then
                sMsg = sMsg & "If you will be using a manifest file, you should create a Sub Main and start your application from that"
            End If
            MsgBox sMsg, vbInformation + vbOKOnly
            
        Else
            MsgBox "Failed to read the manifest file. Ensure proper permissions exist", vbExclamation + vbOKOnly, "Error"
        End If
    End If
    
End Sub

Private Sub pvSelectSubItems(Node As ComctlLib.Node, bSelect As Boolean, bDelete As Boolean)
    
    ' recursive routine to select/unselect all sub-nodes that can be checked/unchecked
    
    If bDelete Then
        Do Until Node.Children = 0
            pvSelectSubItems Node.Child.FirstSibling, False, bDelete
        Loop
        If Node.Image = "select" Or Node.Image = "unselect" Then
            TheManifest.RemoveItem Node.Key
        End If
        tvItems.Nodes.Remove Node.Index
    Else
        Do Until Node Is Nothing
            If Node.Image = "select" And Node.Key <> "rtID" Then
                If bSelect = False Then
                    Node.Image = "unselect"
                    TheManifest.Item(Node.Key).isActive = TheManifest.Item(Node.Key).isActive Xor 1&
                End If
            ElseIf Node.Image = "unselect" Then
                If bSelect = True Then
                    Node.Image = "select"
                    TheManifest.Item(Node.Key).isActive = TheManifest.Item(Node.Key).isActive Or 1&
                End If
            End If
            If Node.Children Then pvSelectSubItems Node.Child.FirstSibling, bSelect, bDelete
            Set Node = Node.Next
        Loop
    End If
    
End Sub
