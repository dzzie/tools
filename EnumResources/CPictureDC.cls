VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPictureDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************
'  Copyright (C)1997, Karl E. Peterson
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' *************************************************************
Option Explicit
'
' Win32 API Declarations, Structures, and Constants
'
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObj Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'
' Bitmap Header Definition
'
Private Type BITMAP '14 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
'
' Member variables
'
Dim m_bmp As BITMAP
Dim m_hDC As Long
Dim m_hBmp As Long
Dim m_hBmpTmp As Long
Dim m_pict As Picture

' ****************************************************
'  Initialization and Termination
' ****************************************************
Private Sub Class_Initialize()
   'Dim hDCTmp As Long
   Dim hWndScn As Long
   Dim hDCScn As Long
   '
   ' Get desktop DC, and create compatable DC.
   '
   hWndScn = GetDesktopWindow()
   hDCScn = GetDC(hWndScn)
   m_hDC = CreateCompatibleDC(hDCScn)
   Call ReleaseDC(hWndScn, hDCScn)
End Sub

Private Sub Class_Terminate()
   '
   ' Clean up resources
   '
   If m_hBmp Then
      Call SelectObject(m_hDC, m_hBmpTmp)
   End If
   Call DeleteDC(m_hDC)
End Sub

' ****************************************************
'  Public Properties
' ****************************************************
Public Property Let hBitmap(NewVal As Long)
   Static PropSet As Boolean
   '
   ' Write-once handle to bitmap
   '
   If PropSet = False Then
      m_hBmp = NewVal
      m_hBmpTmp = SelectObject(m_hDC, m_hBmp)
      Call GetObj(m_hBmp, Len(m_bmp), m_bmp)
      PropSet = True
   End If
End Property

Public Property Get hDC() As Long
   hDC = m_hDC
End Property

Public Property Get bmType() As Long
   bmType = m_bmp.bmType
End Property

Public Property Get Width() As Long
   Width = m_bmp.bmWidth
End Property

Public Property Get Height() As Long
   Height = m_bmp.bmHeight
End Property

Public Property Get WidthBytes() As Long
   WidthBytes = m_bmp.bmWidthBytes
End Property

Public Property Get Planes() As Integer
   Planes = m_bmp.bmPlanes
End Property

Public Property Get BitsPerPixel() As Integer
   BitsPerPixel = m_bmp.bmBitsPixel
End Property

Public Property Get Bits() As Long
   Bits = m_bmp.bmBits
End Property

' ****************************************************
' ****************************************************

