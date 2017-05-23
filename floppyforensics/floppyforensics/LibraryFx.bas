Attribute VB_Name = "LibraryFx"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private binhex(15) As String
Private Initalized As Boolean


Sub Initalize()
    binhex(0) = "0000":   binhex(8) = "1000"
    binhex(1) = "0001":   binhex(9) = "1001"
    binhex(2) = "0010":   binhex(10) = "1010"
    binhex(3) = "0011":   binhex(11) = "1011"
    binhex(4) = "0100":   binhex(12) = "1100":
    binhex(5) = "0101":   binhex(13) = "1101"
    binhex(6) = "0110":   binhex(14) = "1110":
    binhex(7) = "0111":   binhex(15) = "1111"
    Initalized = True
End Sub

Function Hex2Bin(it As String) As String
  If Not Initalized Then Initalize
  Dim tmp As String  'it = 2 char hex string
  If Len(it) = 1 Then it = "0" & it 'need 01 not 1 for val=1
   For i = 1 To 2
      ch = Mid(it, i, 1)
      If IsNumeric(ch) Then
        tmp = tmp & binhex(ch)
      Else
        tmp = tmp & binhex((Asc(ch) - 65 + 10))
      End If      'chr A--> asc65 -->hex chr 10
    Next
  Hex2Bin = tmp
End Function

Sub ShowRtClkMenu(f As Form, t As TextBox, m As Menu)
        LockWindowUpdate t.hwnd
        t.Enabled = False
        DoEvents
        f.PopupMenu m
        t.Enabled = True
        LockWindowUpdate 0&
End Sub

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function cHex(v) As Long
    On Error Resume Next
    cHex = CLng("&h" & v)
End Function

Function removeLast(it, X) As String
    On Error Resume Next
    removeLast = CStr(Mid(it, 1, Len(it) - X))
End Function

Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function FolderExists(path) As Boolean
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function GetParentFolder(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

Sub WriteFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Sub AppendFile(path, it)
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub

Sub CreateFile(path)
    f = FreeFile
    Open path For Random As f
    Close f
End Sub

Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function



Public Function Slice(ary, lbnd, ubnd, Optional joinChr As String = " ")
    If lbnd > ubnd Then Exit Function
    Dim tmp()
    ReDim tmp(ubnd - lbnd)
    For i = 0 To UBound(tmp)
        tmp(i) = ary(lbnd + i)
    Next
    Slice = Join(tmp, joinChr)
End Function

Function byteswap(s)
    hi = Left(s, 2)
    lo = Right(s, 2)
    byteswap = lo & hi
End Function

Function bytesToAscii(it)
    t = Split(Trim(it), " ")
    For i = 0 To UBound(t)
        r = r & Chr("&h" & t(i))
    Next
    bytesToAscii = CStr(r)
End Function

Function IsNT() As Boolean
   Dim myVer As OSVERSIONINFO
   myVer.dwOSVersionInfoSize = 148
   Call GetVersionEx&(myVer)
   If myVer.dwPlatformId = 2 Then IsNT = True
End Function

'ok so this one is misnamed and my math is stinkey poo poo :(
Function RoundUp(s, step)
    Dim r As Long
    r = CLng(s)
    If r Mod step <> 0 Then
        If r < step Then
            r = 0
        Else
            r = r - step
            While r Mod step <> 0: r = r + 1: Wend
        End If
    End If
    RoundUp = r
End Function

