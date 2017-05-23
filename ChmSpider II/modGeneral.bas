Attribute VB_Name = "modGeneral"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global clsfso As New clsFileSystem
Global clsTvm As New clsTreeviewManager
Global clsDlg As New clsCmnDlg
Global pStream As New clsFileStream
Global cStream  As New clsFileStream

Global Const webSite = "http://sandsprite.com/chmspider/"

Sub GraceFullTearDown()
    RemoveSubClass
    End
End Sub

Function br(it) As String
 Dim tmp As String
 tmp = Replace(it, "\n ", vbCrLf) 'notice teh space...necessary ..now will
 tmp = Replace(tmp, "\t ", vbTab)  '  only bug up with a file name in format
 br = Replace(tmp, "\q", """")    '  "c:\folder\n ame_of_file"
End Function

Function aryIsEmpty(ary) As Boolean
    On Error GoTo yup
     Dim l As Long
     l = UBound(ary)
     If l = -1 Then Err.Raise 1
     aryIsEmpty = False
    Exit Function
yup: aryIsEmpty = True
End Function

Function accept(fpath) As Boolean
    Dim strTest As String
    Dim tmp() As String
    Dim i As Long
    
    strTest = clsfso.GetExtension(fpath)
    tmp = Split(frmMain.txtExtensions, " ")
    
    For i = 0 To UBound(tmp)
        If strTest Like tmp(i) Then
           accept = True
           Exit For
        End If
    Next
    
End Function

Sub FilterArray(ary() As String)
    'returns filtered array of just file names
    Dim ret() As String, i As Long
    
    If aryIsEmpty(ary) Then Exit Sub
        
    For i = 0 To UBound(ary)
        If accept(ary(i)) Then
            push ret(), clsfso.FileNameFromPath(CStr(ary(i)))
        End If
    Next
    
    ary = ret()
    
End Sub

Sub SaveMySetting(key, val)
    SaveSetting "ChmSpider", "Config", key, val
End Sub
Function GetMySetting(key, Optional default) As String
    GetMySetting = GetSetting("ChmSpider", "Config", key, default)
End Function

Sub RemoveItemFromCombo(cbo As ComboBox, itemText)
    Dim i As Long
    On Error Resume Next
    For i = 0 To cbo.ListCount
        If cbo.List(i) = itemText Then cbo.RemoveItem i
    Next
End Sub

Function ub(ary) As String
    If aryIsEmpty(ary) Then Exit Function
    ub = CStr(ary(UBound(ary)))
End Function

Sub pop(ary)
    If aryIsEmpty(ary) Then Exit Sub
    If UBound(ary) = 0 Then
        Erase ary
    Else
        ReDim Preserve ary(UBound(ary) - 1)
    End If
End Sub


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function CountDepth(fpath As String) As Long
    If Len(fpath) = 0 Then
        CountDepth = 0: Exit Function
    End If
    
    Dim tmp() As String
    tmp() = Split(fpath, "\")
    CountDepth = UBound(tmp)
    
End Function



'Private Sub cmdGenerateFile_Click()
'
' Dim n As Node
' Dim fname As String
' Dim fPath As String
' Dim tmp As String
' Dim nested As long
' Dim dirStack() As String
' Dim curfolderParent As String
'
' Dim fDeep As long
'
'    If pStream.FileHandle > 0 Then pStream.fClose
'    If cStream.FileHandle > 0 Then cStream.fClose
'
'    pStream.fOpen hhp, otwriting
'    cStream.fOpen hhc, otwriting
'
'    tmp = hhpHeader
'    tmp = Replace(tmp, "<filename>.chm", clsfso.GetBaseName(txtOutPutFileName) & ".chm")
'    tmp = Replace(tmp, "<deftopic>", cboDefaultPage.text, 1, 1)
'    tmp = Replace(tmp, "<title>", txtTitle, 1, 1)
'
'    pStream.WriteLine tmp
'    cStream.WriteLine hhcHeader
'
'     For Each n In tv.Nodes
'        If n.key = "topLevel" Then GoTo nextOne
'
'        fPath = Replace(n.fullpath, tv.Nodes(1).fullpath & "\", "")
'
'        If n.Image = 2 Then 'isFile
'            pStream.WriteLine fPath
'
'            fname = clsfso.GetBaseName(fPath)
'            tmp = Replace(fileNode, "*****", fname)
'            tmp = br(Replace(tmp, "_____", fPath) & "\n \t ")
'
'            cStream.WriteLine tmp
'        Else 'isFolder
'
'            If Len(curfolderParent) = 0 Then
'                curfolderParent = n.fullpath
'                push dirStack, n.fullpath
'                Debug.Print "Cur Parent=0 pushing " & n.fullpath
'            Else
'                curfolderParent = clsfso.GetParentFolder(n.fullpath)
'                'curfolderParent = n.fullpath
'                Debug.Print "Resetting CurFolderparent:" & n.fullpath & "->" & curfolderParent
'            End If
'
'            If curfolderParent = ub(dirStack) Then
'                push dirStack, n.fullpath
'                Debug.Print "Pushing dirstack of child: " & curfolderParent
'            Else
'               Do While curfolderParent <> ub(dirStack)
'
'                    Debug.Print "Reducing curfolderParent=" & curfolderParent
'                    Debug.Print "Reducing ub(dirStack)=" & ub(dirStack)
'
'                    pop dirStack
'
'
'                    If Len(ub(dirStack)) = 0 Then
'                        Debug.Print "DirStack 0 pushing " & n.fullpath
'                        push dirStack, n.fullpath
'                        Exit Do
'                    End If
'
'                    Debug.Print "Closing list tag"
'                    cStream.WriteLine vbTab & "</UL>"
'
'               Loop
'
'            End If
'
'            cStream.WriteLine Replace(fldrNode, "_____", clsfso.FolderName(fPath))
'        End If
'
'nextOne:
'    Next
'
'     pStream.WriteLine br("\n \n [INFOTYPES]\n \n")
'     cStream.WriteLine br("</UL>\n </BODY>\n </HTML>")
'
'     pStream.fClose
'     cStream.fClose
'
'
'
'     ShellExecute Me.hwnd, vbNullString, hhp, vbNullString, 0, 1
'
'
'End Sub
