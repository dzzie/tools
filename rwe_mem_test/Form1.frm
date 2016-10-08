VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)

Private Const PAGE_RWX      As Long = &H40
Private Const MEM_COMMIT    As Long = &H1000

Private base As Long

Private Sub form_load()
    
    Dim b() As Byte
    
    base = VirtualAlloc(ByVal 0&, &H1000, MEM_COMMIT, PAGE_RWX)
     
    ReDim b(&H1000)
    For i = 0 To UBound(b)
        b(i) = &H41
    Next
    
    RtlMoveMemory base, VarPtr(b(0)), UBound(b)
    
    Me.Caption = Hex(base)
     
    
End Sub

