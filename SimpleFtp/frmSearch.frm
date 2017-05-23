VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   300
      Left            =   3360
      TabIndex        =   4
      Top             =   1575
      Width           =   1275
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   285
      Left            =   3375
      TabIndex        =   3
      Top             =   1230
      Width           =   1260
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   330
      Left            =   3420
      TabIndex        =   2
      Top             =   75
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   60
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      ToolTipText     =   "Right Click on me to manage List items"
      Top             =   450
      Width           =   3240
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3255
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveSel 
         Caption         =   "Remove Selected"
      End
      Begin VB.Menu mnuRemUnSel 
         Caption         =   "Remove Un-Selected"
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub PredefinedSearch(search)
    If frmMain.lv.ListItems.Count < 2 Then Exit Sub
    txtSearch = search
    Me.Show
    cmdSearch_Click
End Sub

Private Sub cmdSearch_Click()
    If txtSearch = Empty Then Exit Sub
    List1.Clear
    
    Dim tmp()
    push tmp, Empty
    For i = 0 To UBound(f)
        If f(i).ftype <> fldr Then
            If acceptExpr(f(i).fname, txtSearch) Then push tmp, i & " - " & f(i).fname
        End If
    Next
     
    If UBound(tmp) > 0 Then
        For i = 1 To UBound(tmp): List1.AddItem tmp(i): Next
    Else
        List1.AddItem "- Sorry No Matchs Found -"
        List1.AddItem " "
        List1.AddItem "did you know you can use wildcards?"
        List1.AddItem " "
        List1.AddItem "Example to find all .gif files:"
        List1.AddItem "*.gif , *.g* , *.g*f"
    End If
End Sub

Private Sub cmdDelete_Click()
    If List1.ListCount < 1 Then Exit Sub
    Dim tmp()
    tmp() = GetListBoxContents
    msg = br("Confirm Delete of the following REMOTE FILES ?\n\n") & Join(tmp, vbCrLf)
    If MsgBox(msg, vbYesNo + vbCritical) = vbYes Then
        Dim t As String
        For i = 1 To UBound(tmp)
            If Not ftp.RemoveFile(tmp(i)) Then t = t & tmp(i) & vbCrLf
        Next
        If t <> Empty Then MsgBox br("Error deleting following files:\n") & t, vbCritical
        Call ResetForm
        frmMain.showFileList
    End If
End Sub

Private Sub cmdDownload_Click()
    If List1.ListCount < 1 Then Exit Sub
    Dim tmp()
    tmp() = GetListBoxContents(False)
    msg = br("Are you sure you want to download all of the following files?\n\n") & Join(GetListBoxContents, vbCrLf)
    If MsgBox(msg, vbYesNo) = vbYes Then
        For i = 1 To UBound(tmp)
            If tmp(i) <> Empty Then
                X = CInt(tmp(i))
                ftp.GetFile f(X).fname, dlDir, f(X).byteSize
            End If
        Next
    End If
    Call ResetForm
End Sub

Private Function GetListBoxContents(Optional names As Boolean = True)
    Dim tmp()
    push tmp, Empty 'keep 1 based
    For i = 0 To List1.ListCount - 1
        it = List1.List(i)
        Index = Mid(it, 1, InStr(it, " "))
        fname = Mid(it, Len(Index) + 3, Len(it))
        If names Then push tmp, fname Else push tmp, Index
    Next
    GetListBoxContents = tmp()
End Function

'accepts * style wildcards for parsing strings...kewl
Function acceptExpr(ByVal strTest, find) As Boolean
    
    'these next three lines make it understand diff between
    'appearing in beginning and end of strings should i use it ?
    strTest = Chr(5) & strTest & Chr(5)
    If Left(find, 1) <> "*" Then find = Chr(5) & find
    If Right(find, 1) <> "*" Then find = find & Chr(5)
    
    If InStr(find, "*") > 0 Then tmp = Split(find, "*") _
    Else push tmp, find
    
    Dim t As Integer, f As Integer, p As Integer
    p = 1
    For i = 0 To UBound(tmp)
        If tmp(i) <> Empty Then
              t = t + 1
              X = InStr(p, strTest, tmp(i), vbTextCompare)
              If X > 0 Then p = X: f = f + 1
        End If
    Next
    
    If (t = f And t > 0) Or find = "*" Then acceptExpr = True
End Function


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu Me.mnuPopup
End Sub

Private Sub ResetForm()
    Me.Hide
    List1.Clear
    txtSearch = Empty
End Sub

Private Sub mnuRemoveSel_Click()
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            List1.RemoveItem i
            mnuRemoveSel_Click
            Exit Sub
        End If
    Next
End Sub

Private Sub mnuRemUnSel_Click()
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = False Then
            List1.RemoveItem i
            mnuRemUnSel_Click
            Exit Sub
        End If
    Next
End Sub
