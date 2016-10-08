Attribute VB_Name = "DoctorMemory"
Option Explicit

'=======================================================================
'
'  (c) 2002  Jim White, t/as MathImagical Systems
'            Uki, NSW, Australia
'
'  DoctorMemory
'  ============
'
'    This module exports functions that provide platform-dependent
'    "foreign-process" memory allocation and R/W services via a
'    a common (platform-independent) interface.
'
'    These functions are used by all "Foreign Process Data Extraction" (FPDE)
'    services.
'
'  Export Table:
'
'=======================================================================
'    dmAllocateProcessMemory(nBytes, ProcessId)   => Allocates buffer
'                          (function - returns BufferAddress)
'
'    dmReleaseProcessMemory BufferAddress         => Releases buffer
'=======================================================================
'    dmReadProcessData memBufferAddress,  => Copy from memBuffer to user buffer
'                      vbBufferAddress,
'                      nBytes
'
'    dmWriteProcessData memBufferAddress, => Copy from user Buffer to memBuffer
'                      vbBufferAddress,
'                      nBytes
'=======================================================================
'
   Private PlatformKnown As Boolean  ' have we identified the platform?
   Private NTflag        As Boolean  ' if we have, are we NT or non-NT?
   
   Private fpHandle      As Long     ' the foreign-process instance handle. When we want
                                     ' memory on NT platforms, this is returned to us by
                                     ' OpenProcess, and we pass it in to VirtualAllocEx.
                                     ' We must preserve it, as we need it for read/write
                                     ' operations, and to release the memory when we've
                                     ' finished with it.
'
'================== Platform Identification
'
   Private Type OSVERSIONINFO
      dwOSVersionInfoSize As Long
      dwMajorVersion As Long
      dwMinorVersion As Long
      dwBuildNumber As Long
      dwPlatformId As Long
      szCSDVersion As String * 128
      End Type
   Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
'
'================== Win95/98   Process Memory functions
   Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
   Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
   Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
   Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'
'================== WinNT/2000 Process Memory functions
   Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
   Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
   Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
   Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
   Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'
'
'================== Common Platform services
'
   Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSource As Long, ByVal cBytes As Long)
   Private Declare Function lstrlenA Lib "kernel32" (ByVal lpsz As Long) As Long
   Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
   Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

' ----------
   Const PAGE_READWRITE = &H4
   Const MEM_RESERVE = &H2000&
   Const MEM_RELEASE = &H8000&
   Const MEM_COMMIT = &H1000&
   Const PROCESS_VM_OPERATION = &H8
   Const PROCESS_VM_READ = &H10
   Const PROCESS_VM_WRITE = &H20
   Const STANDARD_RIGHTS_REQUIRED = &HF0000
   Const SECTION_QUERY = &H1
   Const SECTION_MAP_WRITE = &H2
   Const SECTION_MAP_READ = &H4
   Const SECTION_MAP_EXECUTE = &H8
   Const SECTION_EXTEND_SIZE = &H10
   Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
   Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

Public Function dmMemAllocate(ByVal nBytes As Long, ByVal fpID As Long) As Long
   '
   ' Returns pointer to a share-able buffer (size nBytes) in target process
   ' fpID is the foreign process id - we only need it on NT platforms, actually
   '
   If WindowsNT Then
      dmMemAllocate = VirtualAllocNT(fpID, nBytes)
   Else
      dmMemAllocate = VirtualAlloc9X(nBytes)
      End If
   End Function

Public Sub dmMemRelease(mPointer As Long)
   If WindowsNT Then
      VirtualFreeNT mPointer
   Else
      VirtualFree9X mPointer
      End If
   mPointer = 0
   End Sub
   
Public Sub dmReadProcessData(ByVal pBuffer As Long, ByVal pData As Long, ByVal nBytes As Long)
   If WindowsNT Then
      ReadProcessMemory fpHandle, pBuffer, pData, nBytes, 0
   Else
      CopyMemory pData, pBuffer, nBytes
      End If
   End Sub

Public Sub dmWriteProcessData(ByVal pBuffer As Long, ByVal pData As Long, ByVal nBytes As Long)
   If WindowsNT Then
      WriteProcessMemory fpHandle, pBuffer, pData, nBytes, 0
   Else
      CopyMemory pBuffer, pData, nBytes
      End If
   End Sub
' =======================
' end of Public Functions
' =======================
Private Function WindowsNT() As Boolean
   If Not PlatformKnown Then
      Dim verinfo As OSVERSIONINFO
      verinfo.dwOSVersionInfoSize = Len(verinfo)
      If (GetVersionEx(verinfo)) = 0 Then Exit Function  ' in deep doo if this fails
      NTflag = (verinfo.dwPlatformId = 2)
      PlatformKnown = True
      End If
   WindowsNT = NTflag
   End Function

'============================================
'  The NT/2000 Allocate and Release functions
'============================================

Private Function VirtualAllocNT(ByVal fpID As Long, ByVal memSize As Long) As Long
   fpHandle = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, fpID)
   VirtualAllocNT = VirtualAllocEx(fpHandle, ByVal 0&, ByVal memSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
   End Function

Private Sub VirtualFreeNT(ByVal MemAddress As Long)
   Call VirtualFreeEx(fpHandle, ByVal MemAddress, 0&, MEM_RELEASE)
   CloseHandle fpHandle
   End Sub

'============================================
'  The 95/98 Allocate and Release functions
'============================================

Private Function VirtualAlloc9X(ByVal memSize As Long) As Long
   fpHandle = CreateFileMapping(&HFFFFFFFF, 0, PAGE_READWRITE, 0, memSize, vbNullString)
   VirtualAlloc9X = MapViewOfFile(fpHandle, FILE_MAP_ALL_ACCESS, 0, 0, 0)
   End Function

Private Sub VirtualFree9X(ByVal lpMem As Long)
   UnmapViewOfFile lpMem
   CloseHandle fpHandle
   End Sub

'============================================
' a few common-use functions
'============================================

Public Function dmWindowClass(ByVal hWindow As Long) As String
   Dim className As String, cLen As Long
   className = String(64, 0)
   cLen = GetClassName(hWindow, className, 63)
   If cLen > 0 Then className = Left(className, cLen)
   dmWindowClass = className
   End Function

Public Function dmGetStringA(ByVal lpszA As Long) As String
   ' if lpszA is a pointer to ANSI null-terminated string
   ' this will fetch it as a VB string (BSTR)
   Dim sBuf As String, sLen As Long
   sLen = lstrlenA(lpszA)        'get length of string (in chars)
   sBuf = String$(sLen + 2, 0)   'make a buffer to copy to
   CopyMemory StrPtr(sBuf), lpszA, sLen
   dmGetStringA = dmTrimSZ(StrConv(sBuf, vbUnicode))
   End Function

Public Function dmTrimSZ(sName As String) As String
   ' Keep left portion of string sName up to first 0, useful with Win API
   ' null-terminated strings whose length we might not know
   Dim X As Integer
   X = InStr(sName, Chr$(0))
   If X > 0 Then dmTrimSZ = Left$(sName, X - 1) Else dmTrimSZ = sName
   End Function


