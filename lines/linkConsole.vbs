Option Explicit
'LinkConsole.vbs
'
'http://www.angelfire.com/mi4/bvo/vb/vbconio.htm
'
'This is a WSH script used to make it easier to edit
'a compiled VB6 EXE using LINK.EXE to create a console
'mode program.
'
'Drag the EXE's icon onto the icon for this file, or
'execute it from a command prompt as in:
'
'        LinkConsole.vbs 
'
'Be sure to set up strLINK to match your VB6 installation.

Dim strLINK, strEXE, WSHShell

strLINK = """C:\Program Files\Microsoft Visual Studio\VB98\LINK.EXE"""
strEXE = """" & WScript.Arguments(0) & """"
Set WSHShell = CreateObject("WScript.Shell")

WSHShell.Run strLINK & " /EDIT /SUBSYSTEM:CONSOLE " & strEXE

Set WSHShell = Nothing
WScript.Echo "Changed pe to console mode"