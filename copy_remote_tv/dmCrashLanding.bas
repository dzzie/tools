Attribute VB_Name = "dmCrashLanding"
Option Explicit

   ' Dr Memory's CrashLanding - General Protection Fault recovery
   '
   '    by MathImagics, Uki, NSW Australia
   '    mathimagics@yahoo.co.uk
   '
   ' Calling "dmCrashMode" will allow a VB program to recover from run-time
   '    errors that would otherwise be Show-stoppers! Simply call it from the
   '    application's startup routine.
   '
   '        dmCrashMode 0   Terminates process, but nicely
   '  or    dmCrashMode 1   Raise a VB error (Err.Number = 5814)
   '
   '==============================================================
   ' Export Table
   '
   '    dmCrashMode N     ( N = 0 or 1)
   '
   '==============================================================
   
   Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal StdExceptionFilter As Long) As Long
   Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSource As Long, ByVal cBytes As Long)

   Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15

   Private Type EXCEPTION_RECORD
      ExceptionCode As Long
      ExceptionFlags As Long
      pExceptionRecord As Long  ' Pointer to an EXCEPTION_RECORD structure
      ExceptionAddress As Long
      NumberParameters As Long
      ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
      End Type

   Private Type EXCEPTION_POINTERS
      pExceptionRecord As Long
      ContextRecord As Long
      End Type

   Private Const EXCEPTION_EXECUTE_HANDLER = 1&

   Private dmCrashOption  As Integer

   Public Sub dmCrashMode(ByVal opt As Integer)
      dmCrashOption = opt
      Call SetUnhandledExceptionFilter(AddressOf dmCrash)
      End Sub

   Private Function dmCrash(ByRef CrashInfo As EXCEPTION_POINTERS) As Long
      Dim erec As EXCEPTION_RECORD, ecode&
      Dim emsg As String, edesc As String
      '
      '  FPDE applications need to ensure FP buffers are released
      '
      CopyMemory VarPtr(erec), ByVal CrashInfo.pExceptionRecord, Len(erec)
      ecode = erec.ExceptionCode
      Select Case ecode
         Case &HC0000005:  emsg = "General Protection Fault (Access Violation)"
         Case &HC00000FD:  emsg = "General Protection Fault (Stack Overflow"
         Case Else:        emsg = "General Protection Fault, Error code h" & Hex(ecode)
         End Select
      
      Select Case dmCrashOption
         Case 0
            MsgBox "FATAL error => " & emsg & Chr(10) & Chr(10) & "Program will terminate", vbCritical, "Dr Memory"
            dmCrash = EXCEPTION_EXECUTE_HANDLER
         Case Else
            Err.Number = vbObjectError + 5814
            Err.Raise 5814, "[Dr Memory]", emsg
         End Select
   End Function

