VERSION 5.00
Begin VB.UserControl CmnDlg 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   285
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FFFFFF&
   Picture         =   "CmnDlg.ctx":0000
   ScaleHeight     =   270
   ScaleWidth      =   285
   ToolboxBitmap   =   "CmnDlg.ctx":1365
End
Attribute VB_Name = "CmnDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Enum FilterTypes
    textFiles = 0
    htmlFiles = 1
    exeFiles = 2
    zipFiles = 3
    AllFiles = 4
End Enum

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Dim o As OPENFILENAME
Private filters(5)
Private extensions(5)

Private Sub UserControl_Initialize()
    UserControl.Width = 325
    UserControl.Height = 285
    
    filters(0) = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(1) = "Html Files (*.htm*)" + Chr$(0) + "*.htm*" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(2) = "Exe Files (*.exe)" + Chr$(0) + "*.exe" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(3) = "Zip Files (*.zip)" + Chr$(0) + "*.zip" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(4) = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    
    extensions(0) = "txt"
    extensions(1) = "html"
    extensions(2) = "exe"
    extensions(3) = "zip"
    extensions(4) = "bin"
End Sub

Function ShowOpen(initDir, Optional filt As FilterTypes = AllFiles, Optional title = Empty) As String
    o.lStructSize = Len(o)
    o.hwndOwner = GetForegroundWindow()
    o.hInstance = App.hInstance
    o.lpstrFilter = filters(filt)
    o.lpstrFile = Space$(254)
    o.nMaxFile = 255
    o.lpstrFileTitle = Space$(254)
    o.nMaxFileTitle = 255
    o.lpstrInitialDir = initDir
    o.lpstrTitle = title
    o.flags = 0

    ShowOpen = IIf(GetOpenFileName(o), Trim$(o.lpstrFile), Empty)
End Function

Function ShowSave(initDir, Optional filt As FilterTypes = AllFiles, Optional title = Empty, Optional ConfirmOvewrite As Boolean = True) As String
    o.lStructSize = Len(o)
    o.hwndOwner = GetForegroundWindow()
    o.hInstance = App.hInstance
    o.lpstrFilter = filters(filt)
    o.lpstrFile = Space$(254)
    o.nMaxFile = 255
    o.lpstrFileTitle = Space$(254)
    o.nMaxFileTitle = 255
    o.lpstrInitialDir = initDir
    o.lpstrTitle = title
    o.lpstrDefExt = extensions(filt)
    o.flags = 0

    Dim tmp As String
    tmp = IIf(GetSaveFileName(o), Trim$(o.lpstrFile), Empty)
    If ConfirmOvewrite And tmp <> Empty Then
        If FileExists(tmp) Then
            If MsgBox("File Already Exists" & vbCrLf & vbCrLf & "Are you sure you wish to overwrite existing file?", vbYesNo + vbExclamation, "Confirm Overwrite") = vbYes Then ShowSave = tmp
        Else
            ShowSave = tmp
        End If
    Else
        ShowSave = tmp
    End If
End Function

Private Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Private Sub UserControl_Resize()
    UserControl.Width = 325
    UserControl.Height = 285
End Sub



