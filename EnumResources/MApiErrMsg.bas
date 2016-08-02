Attribute VB_Name = "MApiErrMsg"
' *********************************************************************
'  Copyright ©1997-2000 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
   (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
   ByVal dwLanguageId As Long, lpBuffer As Any, ByVal nSize As Long, _
   Arguments As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" _
   (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200&
Private Const FORMAT_MESSAGE_FROM_HMODULE    As Long = &H800&
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Private Const LOAD_LIBRARY_AS_DATAFILE       As Long = 2&

' ---------------------------------------------
' Upper and lower bounds of network errors
' ---------------------------------------------
Private Const NERR_BASE                      As Long = 2100&
Private Const MAX_NERR                       As Long = NERR_BASE + 899&

' ---------------------------------------------
' Upper and lower bounds of Internet errors
' ---------------------------------------------
Private Const INTERNET_ERROR_BASE            As Long = 12000&
Private Const INTERNET_ERROR_LAST            As Long = INTERNET_ERROR_BASE + 171&

Public Function ApiErrorText(ByVal ErrNum As Long) As String
   Dim os As OSVERSIONINFO
   Dim Flags As Long
   Dim hModule As Long
   Dim msg As String
   Dim nRet As Long
   
   ' Load specific error message module, if available.
   Select Case ErrNum
      Case NERR_BASE To MAX_NERR
         ' This module is only available in NT.
         os.dwOSVersionInfoSize = Len(os)
         Call GetVersionEx(os)
         If os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            hModule = LoadLibraryEx("netmsg.dll", 0&, LOAD_LIBRARY_AS_DATAFILE)
         End If
      Case INTERNET_ERROR_BASE To INTERNET_ERROR_LAST
         ' This module is available in most Win9x/NT installs.
         hModule = LoadLibraryEx("wininet.dll", 0&, LOAD_LIBRARY_AS_DATAFILE)
   End Select
   
   ' Build flags for FormatMessage call.
   Flags = FORMAT_MESSAGE_FROM_SYSTEM Or _
           FORMAT_MESSAGE_IGNORE_INSERTS Or _
           FORMAT_MESSAGE_MAX_WIDTH_MASK
   If hModule Then
      Flags = Flags Or FORMAT_MESSAGE_FROM_HMODULE
   End If
   
   ' Prepare buffer, then retrieve error text.
   msg = Space(1024)
   nRet = FormatMessage(Flags, ByVal hModule, ErrNum, 0&, ByVal msg, Len(msg), ByVal 0&)
   If nRet Then
      ApiErrorText = Left(msg, nRet)
   Else
      ApiErrorText = "Error (" & Format(ErrNum) & ") not defined."
   End If
   
   ' Release library, if loaded.
   If hModule Then Call FreeLibrary(hModule)
End Function

