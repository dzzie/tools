VERSION 5.00
Begin VB.UserControl longTimer 
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   InvisibleAtRuntime=   -1  'True
   Picture         =   "longTimer.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   450
   ToolboxBitmap   =   "longTimer.ctx":0920
   Begin VB.Timer Timer1 
      Left            =   570
      Top             =   15
   End
End
Attribute VB_Name = "longTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ticks As Long
Private triggerOn As Long

Public Event Activate()
Public Event Tick()

Public Property Get Enabled() As Boolean
    Enabled = Timer1.Enabled
End Property

Public Property Let Enabled(ByVal vEnabled As Boolean)
    Timer1.Enabled = vEnabled
    ticks = 0
End Property

Public Property Let Interval(mSecs As Long)
    If mSecs > 0 Then Timer1.Interval = mSecs: triggerOn = 1 _
     Else Enabled = False
End Property

Public Property Get Interval() As Long
    Interval = Timer1.Interval
End Property

Public Property Get TicksTillTrigger() As Long
    TicksTillTrigger = triggerOn - ticks
End Property

Private Sub Class_Initialize()
   Enabled = False
   Interval = 1000
End Sub

Private Sub Class_Terminate()
   Enabled = False
End Sub

Public Sub SetMinutes(min As Long)
    triggerOn = min
    Timer1.Interval = 60000
End Sub

Public Sub SetSeconds(secs As Long)
    triggerOn = secs
    Timer1.Interval = 1000
End Sub

Private Sub timer1_timer()
    
    ticks = ticks + 1
    
    If Timer1.Interval >= 1000 Then RaiseEvent Tick
    
    If ticks = triggerOn Then
        RaiseEvent Activate
        ticks = 0
    End If

End Sub
