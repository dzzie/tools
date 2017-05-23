Attribute VB_Name = "bin"
Public Sub binWrite(it, pth As String)
  Dim fin As String       'it= string array of characters
                          'does not matter what base array is
  f = FreeFile
  fin = Join(it, "") 'must be string to join
  Open pth For Binary As f
  Put f, , fin
  Close f
  
  'put filehandle, [insert at byte defult=1], character
  '(byte 1 = first byte of file)
End Sub

Public Function binAppend(it, pth As String) As Integer
  Dim fin As String  'it = array of characters (any base)
  
  If Dir(pth) = "" Then binAppend = 0: Exit Function
  f = FreeFile
  fin = Join(it, "")
  fl = FileLen(pth) + 1
  Open pth For Binary As f
  Put f, fl, fin
  Close f
  binAppend = 1
  
End Function

Public Function binRead(pth As String) 'returns 0 if !path
  On Error GoTo warn
  Dim binFile() As Byte  'f(x) returns string array of characters
  Dim strArray() As String       'ARRAY IS BASE 1
  If Dir(pth) = "" Then binRead = 0: Exit Function
  
  f = FreeFile
  fl = FileLen(pth) 'integer number of bytes in file
  ReDim binFile(1 To fl) 'set array size
  ReDim strArray(1 To fl)
  
  Open pth For Binary As f
  Get f, , binFile()
  Close f
  
  For i = 1 To fl
    strArray(i) = Chr(binFile(i))
  Next
   binRead = strArray 'return base 1 string array
Exit Function
warn: MsgBox "here is the error!" & Err.Description
End Function

Public Function binInsert(it, offset As Integer, pth As String)
  Dim fin As String      'it=string array
  
  If Dir(pth) = "" Then binInsert = 0: Exit Function
  f = FreeFile
  fin = Join(it, "")
  fl = offset
  Open pth For Binary As f
  Put f, fl, fin
  Close f
  binInsert = 1
End Function

Public Function binReadSegment(pth As String, offset As Long, length As Integer)
    Dim binFile() As Byte
    Dim strArray() As String
    f = FreeFile
    
    If length > FileLen(pth) Then length = FileLen(pth)
    If offset = 0 Then offset = 1
    
    ReDim binFile(1 To length) 'set array size
    ReDim strArray(1 To length)
    
    Open pth For Binary As f
    Get f, offset, binFile()
    Close f
    
    For i = 1 To length
      strArray(i) = Chr(binFile(i))
    Next
    binReadSegment = strArray 'return BASE 1 string array

End Function

Public Function HexReadSegment(pth As String, ByVal offset As Long, ByVal length As Long)
    'this assumes offset is 0 based not onebased! always adds 1 to it !
    Dim binFile() As Byte
    Dim strArray() As String
    f = FreeFile
    
    If length > FileLen(pth) Then length = FileLen(pth)
    offset = offset + 1 'one based array to load
    
    ReDim binFile(1 To length) 'set array size
    ReDim strArray(1 To length)
    
    Open pth For Binary As f
    Get f, offset, binFile()
    Close f
    
    For i = 1 To length
      strArray(i) = Hex(binFile(i))
      If Len(strArray(i)) = 1 Then strArray(i) = "0" & strArray(i)
    Next
    HexReadSegment = strArray 'return BASE 0 string array of bytes like 0A, 0D

End Function

Function HexDumpByteArray(ary() As Byte, offset, length) As String
    Dim strArray() As String, x As Variant
    'ary = base 1 byte array offset and length assume base 0 numbers!
    length = length + 1 ' editor display is base 0 need base1 for array
    If offset = UBound(ary) Then MsgBox "Ughh one page to far man"
    If offset + length > UBound(ary) Then length = UBound(ary) - offset
    
    ReDim strArray(1 To length)
    For i = (offset + 1) To (offset + length)
      x = x + 1
      strArray(x) = Hex(ary(i))
      If Len(strArray(x)) = 1 Then strArray(x) = "0" & strArray(x)
    Next
    HexDumpByteArray = HexDump(strArray, offset)
End Function

Public Function HexDump(ary, ByVal offset) As String
    Dim s() As String, chars As String, tmp As String
    
    If offset > 0 And offset Mod 16 <> 0 Then MsgBox "Hexdump isnt being used right! Offset not on boundry"

    'i am lazy and simplicity rules, make sure offset read
    'starts at standard mod 16 boundry or all offsets will
    'be wrong ! it is okay to read a length that ends off
    'boundry though..that was easy to fix...
    
    chars = "   "
    For i = 1 To UBound(ary)
        tmp = tmp & ary(i) & " "
        x = CInt("&h" & ary(i))
        chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            push s, h & "   " & tmp & chars
            offset = offset + 16:  tmp = Empty: chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        h = Hex(offset)
        While Len(h) < 6: h = "0" & h: Wend
        h = h & "   " & tmp
        While Len(h) <= 56: h = h & " ": Wend
        push s, h & chars
    End If
    
    HexDump = Join(s, vbCrLf)
End Function

Function ExtractHexFromDump(dump)
    On Error GoTo bail
    tmp = Split(dump, vbCrLf)
    For i = 0 To UBound(tmp)
        If tmp(i) <> Empty Then
            fs = InStr(tmp(i), "   ") + 3
            l = InStr(fs, tmp(i), "   ") - fs
            tmp(i) = Mid(tmp(i), fs, l)
        End If
    Next
    ExtractHexFromDump = Join(tmp, vbCrLf)
Exit Function
bail: ExtractHexFromDump = dump
End Function
