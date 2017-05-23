Attribute VB_Name = "modLimitResize"
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Type POINTAPI
     X As Long
     y As Long
End Type

Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

' min/max form sizes
Private MIN_WIDTH As Long
Private MAX_WIDTH As Long
Private MIN_HEIGHT As Long
Private MAX_HEIGHT As Long

' private consts
Private Const GWL_WNDPROC = (-4)
Private Const WM_GETMINMAXINFO = &H24

Private OldWindowProc As Long
Private myHwnd As Long

Sub SetMinWidthToCurrent(frm As Form, Optional maxWidPixels As Long, Optional maxHeightPixels As Long)
    MIN_WIDTH = frm.Width \ Screen.TwipsPerPixelX
    MIN_HEIGHT = frm.Height \ Screen.TwipsPerPixelY
    MAX_WIDTH = IIf(maxWidPixels = 0, MIN_WIDTH, maxWidPixels)
    MAX_HEIGHT = IIf(maxHeightPixels = 0, MIN_HEIGHT, maxHeightPixels)
End Sub

Sub SubClassLimitFormResize(hwnd As Long)
    If MIN_WIDTH = 0 Then Err.Raise 1, , "Set size contraints first"
    If myHwnd <> 0 Then Err.Raise 1, , "Already Subclassing"
    If hwnd < 1 Then Err.Raise 1, , "Invalid Window Handle"
    
    myHwnd = hwnd
    OldWindowProc = GetWindowLong(hwnd, GWL_WNDPROC)
    Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Sub RemoveSubClass()
     If myHwnd = 0 Then Exit Sub
     Call SetWindowLong(myHwnd, GWL_WNDPROC, OldWindowProc)
End Sub

Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        
        CopyMemory MinMax, ByVal lParam, Len(MinMax)

        MinMax.ptMinTrackSize.X = MIN_WIDTH
        MinMax.ptMinTrackSize.y = MIN_HEIGHT
        MinMax.ptMaxTrackSize.X = MAX_WIDTH
        MinMax.ptMaxTrackSize.y = MAX_HEIGHT
       
        CopyMemory ByVal lParam, MinMax, Len(MinMax)

        WindowProc = 1
        Exit Function
    End If
   
    'If Msg = WM_SYSCOMMAND Then
    '      If wParam = MyMenuID Then
    '         Call SystemMenuHandler
    '      End If
    'End If
    
   WindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
   
End Function



