VERSION 5.00
Begin VB.Form frmAnalyze 
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2220
   ScaleWidth      =   5025
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Menu mnuF 
      Caption         =   "Extract"
      Begin VB.Menu mnuAutoExtract 
         Caption         =   "Extract Boot Track"
         Index           =   0
      End
      Begin VB.Menu mnuAutoExtract 
         Caption         =   "Extract Fat 1"
         Index           =   1
      End
      Begin VB.Menu mnuAutoExtract 
         Caption         =   "Extract Fat 2"
         Index           =   2
      End
      Begin VB.Menu mnuAutoExtract 
         Caption         =   "Extract Root Directory"
         Index           =   3
      End
      Begin VB.Menu mnuExtractClusters 
         Caption         =   "Extract Cluster Chain"
      End
   End
   Begin VB.Menu spacer1 
      Caption         =   "Parse"
      Index           =   0
      Begin VB.Menu mnuParseFAT 
         Caption         =   "Parse Fat Entries"
      End
      Begin VB.Menu mnuParseDirectory 
         Caption         =   "Parse Directory Entry"
      End
   End
   Begin VB.Menu spacer2 
      Caption         =   "Tools"
      Begin VB.Menu mnuHexedit 
         Caption         =   "HexEdit File"
      End
      Begin VB.Menu mnuDumpRange 
         Caption         =   "HexDump Range"
      End
      Begin VB.Menu mnuConversion 
         Caption         =   "Conversion Tools"
      End
      Begin VB.Menu mnuTxtSearch 
         Caption         =   "Search Text"
      End
      Begin VB.Menu mnuByteSearch 
         Caption         =   "Search Bytes"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnufunctions 
         Caption         =   "Copy "
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnufunctions 
         Caption         =   "Hex Only"
         Index           =   1
      End
      Begin VB.Menu mnufunctions 
         Caption         =   "Text Save"
         Index           =   2
      End
      Begin VB.Menu mnufunctions 
         Caption         =   "Binary Save"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmAnalyze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ImgFile As String
Private myrole As roles

Enum roles
    boot = 0
    fat = 1
    direc = 2
    reprt = 3
    other = 4
End Enum

Sub SetmyRole(r As roles)
    myrole = r
    
    If myrole = reprt Then
        mnufunctions(1).Visible = False
        mnufunctions(3).Visible = False
    Else
        mnufunctions(1).Visible = True
        mnufunctions(3).Visible = True
    End If
End Sub


Private Sub Form_Load()
   Me.Caption = "Examining Binary Image: " & ImgFile
   SetmyRole other
End Sub

Sub Initalize(imageFile)
    On Error Resume Next
    ImgFile = imageFile
    Me.Visible = True
    Text1 = "Image File Loaded Successfully..." & vbCrLf & _
            "Use Menu macros to examine image..." & vbCrLf & vbCrLf & _
            IIf(InStr(ImgFile, ".tmp") < 1 And _
                FileLen(ImgFile) <> FLOP_IMG_SIZE, _
                "Warning ! File Length does not match that of a valid floppy image!", Empty _
                )
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Width = Me.Width - 200
    Text1.Height = Me.Height - 400
End Sub

Private Sub mnuAutoExtract_Click(Index As Integer)
    Dim tmp() As String, s As Long, l As Long
    
    Select Case Index
        Case 0: s = 0: l = &H1FF:       SetmyRole boot
        Case 1: s = &H200: l = &H1200:  SetmyRole fat
        Case 2: s = &H1400: l = &H1200: SetmyRole fat
        Case 3: s = &H2600: l = &H1800: SetmyRole direc
    End Select
    
    tmp() = bin.HexReadSegment(ImgFile, s, l)
    Text1 = bin.HexDump(tmp, s)
End Sub

Private Sub mnuConversion_Click()
    Load frmConv: frmConv.Show
End Sub

Private Sub mnuDumpRange_Click()
    On Error GoTo shit
    t = InputBox("Format: Hex Offset to start at,Hex Length to read" & vbCrLf & vbCrLf & "Note: Offset will be force to be on lowest mod 16 boundry")
    If InStr(t, ",") < 1 Then MsgBox "No comma delimiter exiting": Exit Sub
    v = Split(t, ",")
    o = CInt("&h" & v(0))
    l = CInt("&h" & v(1))
    m = o Mod 16
    If m <> 0 Then: o = o - m: l = l + m
    Dim tmp() As String
    tmp() = bin.HexReadSegment(ImgFile, o, l)
    Text1 = bin.HexDump(tmp, o)
    SetmyRole other
Exit Sub
shit: MsgBox Err.Description
End Sub

Private Sub mnuExtractClusters_Click()
    Dim tmp() As String, toLong As Boolean, f As String, o As Long
    
    a = frmInput.GetInput("Enter Comma delimited list of clusters to extract. Cluster numbers have to be in hex.\n\nRember- data clusters start at cluster 2 which is at 4200h", Text1.SelText)
    If a = Empty Then Exit Sub
    t = Split(a, ",")
    
    SetmyRole other
    If UBound(t) > 5 Then
        toLong = True
        ans = MsgBox("I am working on better support for long results, for now...these results are to long We have to save them directly to file.", vbOKCancel + vbInformation)
        If ans = vbCancel Then Exit Sub
        f = MDIForm1.CmnDlg.ShowSave(App.path, AllFiles, "Save File as")
        If f = Empty Then Exit Sub Else CreateFile f
        SetmyRole reprt
    End If
    
    For i = 0 To UBound(t)
        If t(i) <> Empty Then
            o = (cHex(t(i)) - 2) * 512
            o = &H4200 + o
            tmp() = bin.HexReadSegment(ImgFile, o, 512)
            If toLong Then bin.binAppend tmp(), f _
            Else ret = ret & bin.HexDump(tmp, o) & vbCrLf
        End If
    Next
     
    Text1 = IIf(toLong, "Binary File written: " & f, ret)
   
End Sub

Private Sub mnufunctions_Click(Index As Integer)
    Dim f As String
    If Index > 1 Then
        f = MDIForm1.CmnDlg.ShowSave(App.path, AllFiles, "Save File as")
        If f = Empty Then Exit Sub
    End If
    
    Select Case Index
        Case 0: Clipboard.Clear: Clipboard.SetText IIf(Text1.SelLength > 0, Text1.SelText, Text1)
        Case 1: Text1 = ExtractHexFromDump(Text1)
        Case 2: WriteFile f, IIf(Text1.SelLength > 0, Text1.SelText, Text1)
        Case 3: BinarySaveHexDump Text1, f
    End Select
End Sub

Private Sub mnuHexedit_Click()
    frmHexEdit.loadfile ImgFile
End Sub

Private Sub mnuParseDirectory_Click()
    Dim entry()
    If myrole <> direc Then mnuAutoExtract_Click 3
    SetmyRole reprt
    
    t = bin.ExtractHexFromDump(Text1)
    If Right(t, 2) = vbCrLf Then t = Mid(t, 1, Len(t) - 2)
    'divide up block every 2 lines (32 bytes) = one short file name entry
    tmp = Split(t, vbCrLf)
    For i = 0 To UBound(tmp) Step 2
        push entry(), tmp(i) & " " & tmp(i + 1)
    Next

    'now divide each entry into array based on hex byte value
    Dim LFNTank(), report()
    For i = 0 To UBound(entry)
        dat = Split(entry(i), " ")
        If dat(11) = "0F" Then
            push LFNTank(), entry(i)
        Else
            If Left(entry(i), 8) <> "00 00 00" Then
                push report(), parseDirectoryEntry(CStr(entry(i)), LFNTank())
                Erase LFNTank()
            End If
        End If
    Next
        
    Text1 = Replace(Join(report, vbCrLf), Chr(0), "[Chr(0)]")
    

End Sub

Private Sub mnuParseFAT_Click()
    If myrole <> fat Then mnuAutoExtract_Click 1
    SetmyRole reprt
    
    t = ExtractHexFromDump(Text1)
    Text1 = ParseFatEntries(t)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then ShowRtClkMenu Me, Text1, mnuPopup
End Sub

Sub BinarySaveHexDump(txt, f As String)
        t = Replace(ExtractHexFromDump(txt), vbCrLf, " ")
        t = Replace(t, "  ", " ")
        ary = Split(t, " ")
        For i = 0 To UBound(ary)
            ary(i) = Chr(cHex(ary(i)))
        Next
        binWrite ary, f
End Sub
