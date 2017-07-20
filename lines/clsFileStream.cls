VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum openType
    otRead = 0
    otWrite = 1
    otAppend = 2
    otBinary = 3
End Enum

Private isBinary As Boolean
Private fHandle As Long

Property Get h() As Long
    h = fHandle
End Property

Property Get eof() As Boolean
    eof = VBA.eof(fHandle)
End Property

Property Get lof() As Long
    lof = VBA.lof(fHandle)
End Property

Sub Close_()
    Close fHandle
    fHandle = 0
End Sub

Sub Open_(path, Optional t As openType = otRead)
    Dim f As Long

    f = FreeFile
    Select Case t
        Case otRead: Open path For Input As f
        Case otWrite: Open path For Output As f
        Case otAppend: Open path For Append As f
        Case otBinary: Open path For Binary As f
    End Select
    
    isBinary = (t = otBinary)
    fHandle = f
            
End Sub

Function ReadLine(Optional LineCount = 0) As String
    Dim tmp As String, i As Integer
    For i = 0 To LineCount
        Line Input #fHandle, tmp
        ReadLine = ReadLine & tmp & vbCrLf
    Next
End Function

Sub Write_(str)
    Print #fHandle, str;
End Sub

Sub WriteLine(str)
    Print #fHandle, str
End Sub

Sub WriteExpandedString(ByVal str)
    str = Replace(str, "\t", vbTab)
    str = Replace(str, "\n", vbCrLf)
    WriteLine str
End Sub

'Sub fprintf(formatStr, ParamArray vars())
'    Dim tmp As String, args() As Variant
'    args() = vars()
'    tmp = modGeneral.printf(formatStr, args)
'    WriteLine tmp
'End Sub

Sub WriteDivider(Optional divChar = "-", Optional num = 75)
    WriteLine String(num, divChar)
End Sub

Sub WriteBlankLine(Optional LineCount = 0)
    Dim i As Integer
    For i = 0 To LineCount
        Print #fHandle, ""
    Next
End Sub

Sub BinWriteDump(it, Optional offset = "")
    If Len(offset) = 0 Then
        Put fHandle, , it
    Else
        If offset = 0 Then offset = 1
        Put fHandle, offset, it
    End If
End Sub

Sub BinGetByteArray(buffer() As Byte, Optional offset = "")
    If Len(offset) = 0 Then
        Get fHandle, , buffer()
    Else
        If offset = 0 Then offset = 1
        Get fHandle, offset, buffer()
    End If
End Sub

Sub BinGetStrArray(buffer() As String, Optional offset = "")
    Dim b() As Byte, i As Long
    ReDim b(UBound(buffer))
    
    If Len(offset) = 0 Then
        Get fHandle, , b()
    Else
        If offset = 0 Then offset = 1
        Get fHandle, offset, b()
    End If
    
    For i = 0 To UBound(b)
        buffer(i) = Chr(b(i))
    Next
    
End Sub

Function BinGetChar() As Byte
    Get fHandle, , BinGetChar
End Function

Sub BinWriteNChar(b As Byte, Optional num As Long = 0)
    Dim i As Long
    For i = 0 To num
        Put fHandle, , b
    Next
End Sub

