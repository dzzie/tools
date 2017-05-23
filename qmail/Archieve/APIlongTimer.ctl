VERSION 5.00
Begin VB.UserControl longTimer 
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   390
   InvisibleAtRuntime=   -1  'True
   Picture         =   "longTimer.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   390
   ToolboxBitmap   =   "longTimer.ctx":0920
End
Attribute VB_Name = "longTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mEnabled As Boolean
Private mInterval As Long
Private mStart As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SleepEx Lib "kernel32" _
   (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
   
Public Event Activate()


Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal vEnabled As Boolean)
    mEnabled = vEnabled
    mStart = 0
    If vEnabled = True Then Call Running
End Property

Public Property Let Interval(mSecs As Long)
    If mSecs > 0 Then
      mInterval = mSecs
      Call Running
    Else
      Enabled = False
    End If
End Property

Public Property Get Interval() As Long
    Interval = mInterval
End Property

Private Sub Class_Initialize()
    mEnabled = False
    mInterval = 1000
End Sub

Private Sub Class_Terminate()
    mEnabled = False
End Sub

Public Sub SetMinutes(min As Integer)
    Interval = min * 1000 * 60
End Sub

Public Sub SetSeconds(secs As Long)
    Interval = secs * 1000
End Sub

Private Sub Running()
    Dim Elapsed As Long
    
    Do While mEnabled
         If mStart = 0 Then mStart = GetTickCount
        
         Elapsed = GetTickCount
         
         If (Elapsed - mStart) >= mInterval Then
              mStart = GetTickCount
              RaiseEvent Activate
         End If
         
         DoEvents
         Call SleepEx(200, True)
    Loop
    
End Sub

