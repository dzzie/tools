Attribute VB_Name = "winApi"
Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Const GWL_WNDPROC = (-4)
Const WM_COPYDATA = &H4A
Dim PrevWndProc As Long
Dim PrevhWnd As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lngParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

'Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
'Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Declare Function GetParent& Lib "user32" (ByVal hWnd As Long)

'declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ShowWindowAsync& Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long)
    
'Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal _
'    hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx _
'    As Long, ByVal cy As Long, ByVal wFlags As Long)
     
'Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long

Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Enum windowMessages
'    WM_CLOSE = &H10
'    SW_HIDE = 0
'    SW_MAXIMIZE = 3
'    SW_SHOW = 5
'    SW_MINIMIZE = 6
'    HWND_TOPMOST = -1
'    HWND_NOTOPMOST = -2
    SW_RESTORE = 9
End Enum

'Private Const SWP_SHOWWINDOW = &H40

'Private sPattern As String
'Private hFind As Long

'wild card find window code Arkadiy Olovyannikov
'http://www.pscode.com/xq/ASP/txtCodeId.10259/lngWId.1/qx/vb/scripts/ShowCode.htm
'Public Function EnumWinProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
'    Dim k As Long, sName As String
'    If IsWindowVisible(hWnd) And GetParent(hWnd) = 0 Then
'        sName = Space$(128)
'        k = GetWindowText(hWnd, sName, 128)
'        If k > 0 Then
'            sName = Left$(sName, k)
'            If lParam = 0 Then sName = UCase(sName)
'            If sName Like sPattern Then
 '               hFind = hWnd
 ''               EnumWinProc = 0
 '               Exit Function
 '           End If
 '       End If
 '   End If
'
'EnumWinProc = 1
'End Function


'Public Function FindWindowWild(sWild As String, Optional bMatchCase As Boolean = True) As Long
'    sPattern = sWild
'    If Not bMatchCase Then sPattern = UCase(sPattern)
'    EnumWindows AddressOf EnumWinProc, bMatchCase
'    FindWindowWild = hFind
'End Function

'Sub SendWindowMessage(hwin, msg As windowMessages)
'
'      If msg = WM_CLOSE Then
'         SendMessage hwin, msg, 0, 0
'      ElseIf msg = HWND_NOTOPMOST Or msg = HWND_TOPMOST Then
'         SetWindowPos hwin, msg, 100, 100, 500, 700, SWP_SHOWWINDOW
'      Else
'         ShowWindow hwin, msg
'      End If
'
'End Sub

Sub SendMsgToPreviousInstance(msg As String)
    Dim h As Long
    
    h = CLng(GetSetting(App.Title, "frmMessage", "hWnd", 0))
    If h <> 0 Then SendData h, msg

End Sub

Sub ShowPreviousInstance()
    Dim h As Long
    h = CLng(GetSetting(App.Title, "frmMessage", "hWnd", 0))
    If h <> 0 Then
        If IsIconic(h) Then
            ShowWindowAsync h, CLng(SW_RESTORE)
            BringWindowToTop h
        Else
            BringWindowToTop h
            SetForegroundWindow h
        End If
    End If
End Sub

'below fx from Keith Weimer
'http://pscode.com/xq/ASP/txtCodeId.28394/lngWId.1/qx/vb/scripts/ShowCode.htm
'
'Sends previous instance of app the command line received by this instance
'by grabbing its window handle and sending a message to that instances
'window handling proceedure which we altered when the other instance started up
'to be able to react to these messages..pretty slick *nods* boy do i need a good
'book on API :-\

Private Sub SendData(hWnd As Long, Data As String)
    'On Error Resume Next
    
    Dim Buffer(1 To 2048) As Byte
    Dim CopyData As COPYDATASTRUCT
    
    CopyMemory Buffer(1), ByVal Data, Len(Data)
    CopyData.dwData = 3
    CopyData.cbData = Len(Data) + 1
    CopyData.lpData = VarPtr(Buffer(1))
    SendMessage hWnd, WM_COPYDATA, hWnd, CopyData
End Sub

Sub Hook(hWnd As Long)
    'On Error Resume Next
    
    PrevhWnd = hWnd
    PrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Sub Unhook()
    On Error Resume Next
    
    SetWindowLong PrevhWnd, GWL_WNDPROC, PrevWndProc
End Sub

Private Sub InterProcessComms(lParam As Long)
    'On Error Resume Next
    
    Dim CopyData As COPYDATASTRUCT
    Dim Buffer(1 To 2048) As Byte
    Dim Temp As String
    
    CopyMemory CopyData, ByVal lParam, Len(CopyData)
    
    If CopyData.dwData = 3 Then
        CopyMemory Buffer(1), ByVal CopyData.lpData, CopyData.cbData
        Temp = StrConv(Buffer, vbUnicode)
        Temp = Left$(Temp, InStr(1, Temp, Chr$(0)) - 1)
        'heres where we work with the intercepted message
        Library.ParseCommandLine Temp
    End If
End Sub

Private Function WindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    
    If wMsg = WM_COPYDATA Then InterProcessComms lParam
    WindowProc = CallWindowProc(PrevWndProc, hWnd, wMsg, wParam, lParam)
End Function

