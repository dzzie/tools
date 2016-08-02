Attribute VB_Name = "MEnumResources"
' *********************************************************************
'  Copyright ©2000 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Private Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Any, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

' Predefined Resource Types
Public Enum ResTypes
   RT_CURSOR = 1&
   RT_BITMAP = 2&
   RT_ICON = 3&
   RT_MENU = 4&
   RT_DIALOG = 5&
   RT_STRING = 6&
   RT_FONTDIR = 7&
   RT_FONT = 8&
   RT_ACCELERATOR = 9&
   RT_RCDATA = 10&
   RT_MESSAGETABLE = 11&
   RT_GROUP_CURSOR = 12&
   RT_GROUP_ICON = 14&
   RT_VERSION = 16&
   RT_DLGINCLUDE = 17&
   RT_PLUGPLAY = 19&
   RT_VXD = 20&
   RT_ANICURSOR = 21&
   RT_ANIICON = 22&
   RT_HTML = 23&
End Enum

' LoadLibraryEx flags
Private Const DONT_RESOLVE_DLL_REFERENCES = &H1
Private Const LOAD_LIBRARY_AS_DATAFILE = &H2
Private Const LOAD_WITH_ALTERED_SEARCH_PATH = &H8

' Reference to callback interface
Private m_Callback As IEnumResources

' ***************************************************
'  Public enumeration entry points
' ***************************************************
Public Function EnumResources(Optional ByVal ModuleName As String = "") As Boolean
   Dim hModule As Long
   Dim FreeLib As Boolean
   Static Busy As Boolean
   
   ' This routine is *not* re-entrant!
   If Not Busy Then
      Busy = True
      
      ' Load library if needed. An empty string indicates
      ' that the current thread's module should be used.
      If Len(ModuleName) Then
         ' Check first to see if the module is already
         ' mapped into this process.
         hModule = GetModuleHandle(ModuleName)
         If hModule = 0 Then
            hModule = LoadLibraryEx(ModuleName, 0&, LOAD_LIBRARY_AS_DATAFILE)
            If hModule = 0 Then
               ' Problem --> can't load module!
               Debug.Print "LoadLibraryEx error (" & Err.LastDllError;
               Debug.Print "): " & ApiErrorText(Err.LastDllError)
               Busy = False
            Else
               ' Set a flag that reminds us to free this handle.
               FreeLib = True
            End If
         End If
      Else
         ' Enumerate currently running module (hModule=0).
      End If
      
      ' Only enumerate if no problems loading module.
      If Busy Then
         ' Start standard enumeration.
         Call EnumResourceTypes(hModule, AddressOf EnumResTypeProc, 0&)
         ' Close handle for any module we loaded.
         If FreeLib Then Call FreeLibrary(hModule)
         ' Clear re-entry flag, and return success.
         Busy = False
         EnumResources = True
      End If
   End If
End Function

Public Function EnumResourcesEx(Callback As IEnumResources, Optional ByVal ModuleName As String = "") As Boolean
   ' This routine is *not* re-entrant!
   If m_Callback Is Nothing Then
      Set m_Callback = Callback
         EnumResourcesEx = EnumResources(ModuleName)
      Set m_Callback = Nothing
   End If
End Function

' ***************************************************
'  Private enumeration callback routines
' ***************************************************
Private Function EnumResTypeProc(ByVal hModule As Long, ByVal lpszType As Long, ByVal lParam As Long) As Long
   'BOOL CALLBACK EnumResTypeProc(
   '    HANDLE hModule,  // resource-module handle
   '    LPTSTR lpszType, // pointer to resource type
   '    LONG lParam   // application-defined parameter
   '   );
   Dim ResType As String
   Dim Continue As Boolean
   
   ' Retrieve type or custom name of resource.
   Debug.Print ResTypeName(lpszType);
   ResType = DecodeResTypeName(lpszType)
   Debug.Print " (" & ResType & ")"
   
   ' Alert client. Continue enum by default.
   Continue = True
   If Not (m_Callback Is Nothing) Then
      m_Callback.EnumResourceSink hModule, "", ResType, Continue
   End If
   
   ' Enumerate resource names of this type.
   Call EnumResourceNames(hModule, lpszType, AddressOf EnumResNameProc, lParam)
   
   ' Continue enumeration...
   EnumResTypeProc = Continue
End Function
   
Private Function EnumResNameProc(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByVal lParam As Long) As Long
   'BOOL CALLBACK EnumResNameProc(
   '    HANDLE hModule,  // resource-module handle
   '    LPCTSTR lpszType,   // pointer to resource type
   '    LPTSTR lpszName, // pointer to resource name
   '    LONG lParam   // application-defined parameter
   '   );
   Dim ResName As String
   Dim ResType As String
   Dim Continue As Boolean
   Dim Buffer As String
   Dim nRet As Long
   Dim i As Long
   
   ' Retrieve resource ID.
   ResType = DecodeResTypeName(lpszType)
   ResName = DecodeResTypeName(lpszName)
   Debug.Print " --> "; ResName
   
   If lpszType = RT_STRING Then
      ' We have a block of 16 strings for this entry.
      ' Must determine which are valid.
      If Not HiWord(lpszName) Then
         ' Allocate a good-sized buffer.
         Buffer = Space$(255)
         For i = ((lpszName - 1) * 16) To (lpszName * 16)
            nRet = LoadString(hModule, i, Buffer, Len(Buffer))
            If nRet Then
               Debug.Print " ----> "; "#" & Format$(i) & "  (""" & Left$(Buffer, nRet) & """)"
               ' Alert client. Continue enum by default.
               Continue = True
               If Not (m_Callback Is Nothing) Then
                  m_Callback.EnumResourceSink hModule, "#" & Format$(i), ResType, Continue
               End If
            End If
         Next i
      End If
      
   Else
      ' Alert client. Continue enum by default.
      Continue = True
      If Not (m_Callback Is Nothing) Then
         m_Callback.EnumResourceSink hModule, ResName, ResType, Continue
      End If
   End If

   ' Continue enumeration...
   EnumResNameProc = Continue
End Function

' ***************************************************
'  Public utility methods
' ***************************************************
Public Function ResTypeName(ByVal ResType As ResTypes) As String
   Select Case ResType
      Case RT_ACCELERATOR
         ResTypeName = "Accelerator table"
      Case RT_ANICURSOR
         ResTypeName = "Animated cursor"
      Case RT_ANIICON
         ResTypeName = "Animated icon"
      Case RT_BITMAP
         ResTypeName = "Bitmap resource"
      Case RT_CURSOR
         ResTypeName = "Hardware-dependent cursor resource"
      Case RT_DIALOG
         ResTypeName = "Dialog box"
      Case RT_DLGINCLUDE
         ResTypeName = "Header file that contains menu and dialog box #define statements"
      Case RT_FONT
         ResTypeName = "Font resource"
      Case RT_FONTDIR
         ResTypeName = "Font directory resource"
      Case RT_GROUP_CURSOR
         ResTypeName = "Hardware-independent cursor resource"
      Case RT_GROUP_ICON
         ResTypeName = "Hardware-independent icon resource"
      Case RT_HTML
         ResTypeName = "HTML document"
      Case RT_ICON
         ResTypeName = "Hardware-dependent icon resource"
      Case RT_MENU
         ResTypeName = "Menu resource"
      Case RT_MESSAGETABLE
         ResTypeName = "Message-table entry"
      Case RT_PLUGPLAY
         ResTypeName = "Plug and play resource"
      Case RT_RCDATA
         ResTypeName = "Application-defined resource (raw data)"
      Case RT_STRING
         ResTypeName = "String-table entry"
      Case RT_VERSION
         ResTypeName = "Version resource"
      Case RT_VXD
         ResTypeName = "VXD"
      Case Else
         ResTypeName = "User-defined custom resource"
   End Select
End Function

' ***************************************************
'  Public utility properties
' ***************************************************
Public Property Get LoWord(LongIn As Long) As Integer
   Call CopyMemory(LoWord, LongIn, 2)
End Property

Public Property Let LoWord(LongIn As Long, ByVal NewWord As Integer)
   Call CopyMemory(LongIn, NewWord, 2)
End Property

Public Property Get HiWord(LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Property

Public Property Let HiWord(LongIn As Long, ByVal NewWord As Integer)
   Call CopyMemory(ByVal (VarPtr(LongIn) + 2), NewWord, 2)
End Property

' ***************************************************
'  Private utility methods
' ***************************************************
Private Function DecodeResTypeName(ByVal lpszValue As Long) As String
   If HiWord(lpszValue) Then
      ' Pointers will always be >64K
      DecodeResTypeName = PointerToStringA(lpszValue)
   Else
      ' Otherwise we have an ID.
      DecodeResTypeName = "#" & CStr(lpszValue)
   End If
End Function

Private Function PointerToDWord(ByVal lpDWord As Long) As Long
   Dim RetVal As Long
   Call CopyMemory(RetVal, ByVal lpDWord, 4)
   PointerToDWord = RetVal
End Function

Private Function PointerToStringW(ByVal lpString As Long) As String
   Dim sText As String
   Dim lLength As Long
   
   If lpString Then
      lLength = lstrlenW(lpString)
      If lLength Then
         sText = Space$(lLength)
         CopyMemory ByVal StrPtr(sText), ByVal lpString, lLength * 2
      End If
   End If
   PointerToStringW = sText
End Function

Private Function PointerToStringA(lpStringA As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringA, nLen
         PointerToStringA = StrConv(Buffer, vbUnicode)
      End If
   End If
End Function


