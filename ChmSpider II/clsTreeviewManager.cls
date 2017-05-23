VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreeviewManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public nodeKeys As New clsCollection
'Private tv As treeview 'a dll shouldnt need that shit should be as stand alone as possible
Private tv As Object
Private Const tvwChild = 4

Sub Reset()
    tv.Nodes.Clear
    nodeKeys.ResetCollection
End Sub

Sub SetTreeviewReference(t As Object)
    Set tv = t
    nodeKeys.Name = "nodeKeys for Tv.tag: " & tv.Tag
End Sub

Property Let DebugOutput(b As Boolean)
    nodeKeys.DebugOutput = b
End Property

Function AddNodeUnder(text, pNodeId, Optional ico As Long = 0) As String
    'returns key of child node added
    'rember all ids have to be logged to nodekeys
    'because duplicate key raises error! damn treeviews all to heck!
    'GenUniqueID should not fail now...think all debugged
    On Error GoTo oghhShit
    Dim uid As String
    uid = nodeKeys.GenUniqueKey()
    nodeKeys.Add uid, text
    tv.Nodes.Add pNodeId, tvwChild, uid, text, IIf(ico > 0, ico, "")
    AddNodeUnder = uid
    Exit Function
oghhShit: MsgBox Err.Description & vbCrLf & vbCrLf & "pid='" & pNodeId & "' uid='" & uid & "' Text=" & text
End Function

Sub LoadArrayUnderNode(ary, pNodeId, Optional ico As Long = 0)
    Dim i As Long
    If aryIsEmpty(ary) Then Exit Sub
    For i = 0 To UBound(ary)
        AddNodeUnder ary(i), pNodeId, ico
    Next
End Sub

Sub BuildTreePath(folderpath)
    Dim ret(), i As Long, uid As String, Child As String, pid As String
    If Right(folderpath, 1) = "/" Then folderpath = folderpath & "<default>"
    ret() = GetPathsInStep(folderpath)
    For i = 0 To UBound(ret)
        If Not TreePathExists(ret(i)) Then
            If i = 0 Then
                uid = nodeKeys.GenUniqueKey()
                tv.Nodes.Add , , uid, ret(0), 1
                nodeKeys.Add uid, ret(0)
            Else
                Child = ParentObjName(ret(i))
                pid = nodeKeys.keyFromValue(ret(i - 1))
                uid = nodeKeys.GenUniqueKey()
                tv.Nodes.Add pid, tvwChild, uid, Child, IIf(i = UBound(ret), 3, 2)
                nodeKeys.Add uid, ret(i)
            End If
        End If
    Next
End Sub

Function ParentObjName(it) As String
    If InStr(it, "/") > 0 Then
        ParentObjName = Mid(it, InStrRev(it, "/") + 1)
    Else
        ParentObjName = CStr(it)
    End If
End Function

Function TreePathExists(treePath) As Boolean
    Dim i As Long, tmp As String
    For i = 1 To tv.Nodes.Count
        tmp = tv.Nodes(i).fullpath
        If tmp = treePath Then TreePathExists = True: Exit For
    Next
    'Debug.Print "Looking for " & treePath & " - " & IIf(TreePathExists, "Found", "NotFound")
End Function

Function GetPathsInStep(fullpath) As Variant()
    Dim ret(), it, tmp() As String, i As Long
    'fullpath = var/www/htdocs/
    'ret(0) = var/
    'ret(1) = var/www/
    'ret(2) = var/www/htdocs/
    tmp = Split(fullpath, "/")
     
    For i = 0 To UBound(tmp)
        it = it & IIf(it = "", "", "/") & tmp(i)
        push ret(), it
    Next
    GetPathsInStep = ret()
End Function

