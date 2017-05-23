VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7245
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin floppyForensics.CmnDlg CmnDlg 
      Left            =   6960
      Top             =   0
      _ExtentX        =   582
      _ExtentY        =   503
   End
   Begin VB.Menu mnuF 
      Caption         =   "File"
      Begin VB.Menu mnuImageFlp 
         Caption         =   "Image Floppy"
      End
      Begin VB.Menu mnuMountFlp 
         Caption         =   "Mount Floppy Image"
      End
      Begin VB.Menu mnuMountRawFlop 
         Caption         =   "Mount Raw Floppy"
      End
      Begin VB.Menu mnuImage2Floppy 
         Caption         =   "Burn Image to Floppy"
      End
      Begin VB.Menu mnuHexedit 
         Caption         =   "Hexedit File"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
    'frmHexEdit.Visible = True
    'Me.Visible = False
   If Not IsNT() Then
        Me.Caption = "Imaging not enabled for non-NT systems"
        TurnOff mnuImageFlp
        TurnOff mnuMountRawFlop
        TurnOff mnuImage2Floppy
   End If
End Sub

Sub TurnOff(m As Menu)
    m.Enabled = False
    m.Caption = m.Caption & " (NT Only)"
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim d As New frmAnalyze
    d.Initalize Data.files(1)
End Sub

Private Sub mnuHexedit_Click()
    f = CmnDlg.ShowOpen(App.path, AllFiles, "Select Image File")
    If f = Empty Then Exit Sub
    frmHexEdit.loadfile f
End Sub

Private Sub mnuImage2Floppy_Click()
    f = CmnDlg.ShowOpen(App.path, AllFiles, "Select Image File")
    If f = Empty Then Exit Sub
    If FileLen(f) <> FLOP_IMG_SIZE Then MsgBox "File length is wrong exiting", vbCritical: Exit Sub
    SaveImageToFloppy f
End Sub

Private Sub mnuImageFlp_Click()
    ImageFloppy CmnDlg.ShowSave(App.path, AllFiles, "Save Image As", True)
End Sub

Private Sub mnuMountFlp_Click()
    f = CmnDlg.ShowOpen(App.path, AllFiles, "Select Image To Open")
    If FileExists(f) And f <> Empty Then
        Dim d As New frmAnalyze
        d.Initalize f
    End If
End Sub

Private Sub mnuMountRawFlop_Click()
    General.ReadTrack 1, 1
    Dim d As New frmAnalyze
    d.Initalize App.path & "\flp.tmp"
    d.Text1 = d.Text1 & vbCrLf & vbCrLf & "Note: only first track has been imaged." & vbCrLf & vbTab & "To Access data beyond 8FFF requires full image"
End Sub
