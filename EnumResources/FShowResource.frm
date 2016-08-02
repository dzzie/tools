VERSION 5.00
Begin VB.Form FShowResource 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "FShowResource.frx":0000
      Top             =   420
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "FShowResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©2000 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Const LOAD_LIBRARY_AS_DATAFILE = &H2

Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Any) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

Private Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private m_hIcon As Long

Private m_hBitmap As Long
Private m_DC As CPictureDC

Private m_hModule As Long
Private m_FreeLib As Boolean

Public Sub DisplayBitmap(ByVal FileName As String, ByVal ID As String)
   ' Try to load resource if we have module mapped.
   If LoadModule(FileName) Then
      If InStr(ID, "#") = 1 Then
         m_hBitmap = LoadBitmap(m_hModule, CLng(Mid$(ID, 2)))
      Else
         m_hBitmap = LoadBitmap(m_hModule, ID)
      End If
      
      If m_hBitmap Then
         Set m_DC = New CPictureDC
         m_DC.hBitmap = m_hBitmap
         Me.Caption = "Bitmap: " & ID
         Me.Show
      Else
         Debug.Print "LoadBitmap Error: " & Err.LastDllError
      End If
   End If

   ' Unload if we failed.
   If Me.Visible = False Then Unload Me
End Sub

Public Sub DisplayIcon(ByVal FileName As String, ByVal ID As String)
   ' Try to load resource if we have module mapped.
   If LoadModule(FileName) Then
      If InStr(ID, "#") = 1 Then
         m_hIcon = LoadIcon(m_hModule, CLng(Mid$(ID, 2)))
      Else
         m_hIcon = LoadIcon(m_hModule, ID)
      End If
      
      If m_hIcon Then
         Me.Caption = "Icon: " & ID
         Me.Show
      Else
         Debug.Print "LoadIcon error (" & Err.LastDllError;
         Debug.Print "): " & ApiErrorText(Err.LastDllError)
      End If
   End If

   ' Unload if we failed.
   If Me.Visible = False Then Unload Me
End Sub

Public Sub DisplayString(ByVal FileName As String, ByVal ID As String, Optional ByVal MaxLength As Long = 255)
   Dim nRet As Long
   Dim uID As Long
   Dim Buffer As String
   
   ' Allocate a good-sized buffer.
   Buffer = Space$(MaxLength)
   
   ' Try to load resource if we have module mapped.
   If LoadModule(FileName) Then
      If InStr(ID, "#") = 1 Then
         uID = CLng(Mid$(ID, 2))
      Else
         uID = CLng(ID)
      End If
      nRet = LoadString(m_hModule, uID, Buffer, Len(Buffer))
      
      If nRet Then
         Buffer = FixLFs(Left$(Buffer, nRet))
         With Text1
            .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            .Text = Buffer
            .Visible = True
         End With
         Me.Caption = "String: " & ID
         Me.Show
      Else
         ' Problem --> can't load string!
         Debug.Print "LoadString error (" & Err.LastDllError;
         Debug.Print "): " & ApiErrorText(Err.LastDllError)
      End If
   End If

   ' Unload if we failed.
   If Me.Visible = False Then Unload Me
End Sub

Private Function LoadModule(ByVal FileName As String) As Boolean
   Dim hModule As Long
   
   ' Clear flag to free module
   m_FreeLib = False
   
   ' Check first to see if the module is already
   ' mapped into this process.
   hModule = GetModuleHandle(FileName)
   If hModule = 0 Then
      hModule = LoadLibraryEx(FileName, 0&, LOAD_LIBRARY_AS_DATAFILE)
      If hModule = 0 Then
         ' Problem --> can't load module!
         Debug.Print "LoadLibraryEx error (" & Err.LastDllError;
         Debug.Print "): " & ApiErrorText(Err.LastDllError)
      Else
         m_FreeLib = True
      End If
   End If
   
   ' Cache module handle and return success
   m_hModule = hModule
   LoadModule = (hModule <> 0)
End Function

Private Sub Form_Load()
   ' Clean-up UI
   Set Me.Icon = Nothing
End Sub

Private Sub Form_Paint()
   Const SRCCOPY As Long = &HCC0020
   Const Offset As Long = 10


   ' Free library, if need be.
   If m_FreeLib Then
      Call FreeLibrary(m_hModule)
      m_FreeLib = False
      m_hModule = 0
   End If
   
   ' Display icon, if we have a handle.
   If m_hIcon Then
      Call DrawIcon(Me.hDC, Offset, Offset, m_hIcon)
   End If
   
   ' Display bitmap, if we have an object.
   If m_hBitmap Then
      Call BitBlt(Me.hDC, Offset, Offset, m_DC.Width, m_DC.Height, m_DC.hDC, 0, 0, SRCCOPY)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Clean up bitmap object
   If m_hBitmap Then
      Set m_DC = Nothing
      Call DeleteObject(m_hBitmap)
   End If
   
   ' Unload library, if we loaded it
   If m_FreeLib Then
      Call FreeLibrary(m_hModule)
   End If
End Sub

Private Function FixLFs(ByVal Buffer As String) As String
   Dim n As Long
   
   ' Convert all single LFs to CR/LF pairs
   n = InStr(Buffer, vbLf)
   Do While n
      If n > 1 Then
         If Asc(Mid$(Buffer, n - 1)) <> 13 Then
            If n < Len(Buffer) Then
               Buffer = Left$(Buffer, n - 1) & vbCrLf & Mid$(Buffer, n + 1)
            Else
               Buffer = Left$(Buffer, n - 1) & vbCrLf
            End If
         End If
      ElseIf n = 1 Then
         Buffer = vbCrLf & Mid$(Buffer, 2)
      End If
         
      n = InStr(n + 1, Buffer, vbLf)
   Loop
   FixLFs = Buffer
End Function
