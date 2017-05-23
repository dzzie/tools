VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Timer Timer1 
      Interval        =   3500
      Left            =   4200
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As New CWindowsSystem

Private Sub Form_Resize()
    List1.Width = Me.Width
    List1.Height = Me.Height
End Sub

Private Sub Timer1_Timer()
    
    Dim c As Collection
    Dim w As CWindow
    
    List1.Clear
    Set c = ws.ChildWindows()
        
    For Each w In c
        If w.TopMost And w.Visible Then
            List1.AddItem "Unsettings topmost for 0x" & Hex(w.hwnd) & " - " & w.className & " - " & w.Caption
            w.TopMost = False
        End If
    Next
     
End Sub
