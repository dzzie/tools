Attribute VB_Name = "ShellWait"
Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&
Const INFINITE& = &HFFFF


Private Declare Function GetWindowsDirectory _
    Lib "kernel32" _
    Alias "GetWindowsDirectoryA" ( _
    ByVal lpBuffer As String, _
    ByVal nSize As Long _
    ) As Long


Private Declare Function OpenProcess _
    Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long _
    ) As Long


Private Declare Function WaitForSingleObject _
    Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long _
    ) As Long


Private Declare Function GetExitCodeProcess _
    Lib "kernel32" ( _
    ByVal hProcess As Long, _
    lpExitCode As Long _
    ) As Long


Private Declare Function CloseHandle _
    Lib "kernel32" ( _
    ByVal hObject As Long _
    ) As Long


Function ShellAndWait(cmdLine, wFocus As VbAppWinStyle) As Long
    On Error GoTo blah
    Dim idProg As Long, iExit As Long
    idProg = Shell(cmdLine, wFocus)
    iExit = fWait(idProg)
    ShellAndWait = iExit
    Exit Function
blah:
    MsgBox "ERROR in Shell&WAIT with commandline: " & cmdLine, vbCritical
    Err.Raise 21, "ShellAndWait", "Shell Error"
End Function


Private Function fWait(ByVal lProgID As Long) As Long
    Dim lExitCode As Long, hdlProg As Long
    ' Get program handle
    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, lProgID)
    ' Get current program exit code
    GetExitCodeProcess hdlProg, lExitCode

    Do While lExitCode = STILL_ACTIVE&
        DoEvents: DoEvents: DoEvents: DoEvents
        GetExitCodeProcess hdlProg, lExitCode
    Loop
   
   CloseHandle hdlProg
   fWait = lExitCode
End Function


Function fGetWinDir() As String
    ' Wrapper to return OS Path
    Dim lRet As Long, lSize As Long, sBuf As String * 512
    lSize = 512
    lRet = GetWindowsDirectory(sBuf, lSize)
    fGetWinDir = Left(sBuf, InStr(1, sBuf, Chr(0)) - 1)
End Function


