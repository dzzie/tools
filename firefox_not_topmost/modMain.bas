Attribute VB_Name = "modMain"
Public ChildWindows As Collection

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public classFilter As String

Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim c As New CWindow
    c.hWnd = hWnd
    If Not IsObject(ChildWindows) Then Set ChildWindows = New Collection
    If Len(classFilter) > 0 Then
        If InStr(1, c.className, classFilter, vbTextCompare) > 0 Then ChildWindows.Add c 'module level collection object...
    Else
        ChildWindows.Add c 'module level collection object...
    End If
    EnumChildProc = 1  'continue enum
End Function


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = IIf(IsEmpty(value), "", value)
End Sub

Function ColToStr(c As Collection) As String
    Dim tmp() As String
    
    For Each x In c
        push tmp, x
    Next
    
    ColToStr = Join(tmp, vbCrLf)
End Function

Function WindowUnderCursor() As Long
    Dim p As POINTAPI
    GetCursorPos p
    WindowUnderCursor = WindowFromPoint(p.x, p.Y)
End Function
