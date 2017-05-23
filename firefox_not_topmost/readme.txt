firefox has been getting stuck topmost on me lately so annoying!
this will constantly scan for open firefox windows and unset it

Dim ws As New CWindowsSystem

Private Sub Timer1_Timer()
    
    Dim c As Collection
    Dim w As CWindow
    
    List1.Clear
    
    Set c = ws.ChildWindows(, "MozillaWindowClass")
    If c.count = 0 Then
        List1.AddItem "Nothing to do " & Now
    Else
        For Each w In c
            w.TopMost = False
            List1.AddItem Now & " Setting not top most: " & w.hWnd
        Next
    End If
    
End Sub

update: 3.29.17 - so apparentlty other windows are now following suit
so we will enum all, see if its topmost and remove that if it is..must have
been a shit windows update.

M$: so do you want your free upgrade to Win10 now? - annddd how about now??


Private Sub Timer1_Timer()
    
    Dim c As Collection
    Dim w As CWindow
    
    List1.Clear
    
    'Set c = ws.ChildWindows(, "MozillaWindowClass")
    Set c = ws.ChildWindows()
        
    For Each w In c
        If w.TopMost And w.Visible Then
            List1.AddItem "Unsettings topmost for 0x" & Hex(w.hwnd) & " - " & w.className & " - " & w.Caption
            w.TopMost = False
        End If
    Next
     
End Sub

