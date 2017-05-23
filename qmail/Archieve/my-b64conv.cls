VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "b64convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum ArgType
    isPath = 1
    isString = 0
End Enum
    
Private Type lib 'conversion library
  b64Chr(65) As String
  binHex(16) As String
  hexChr(16) As String
End Type
Dim lib As lib

Private Sub Class_Initialize()
  Call initAlpha
End Sub

Public Sub b64FileEncode(path As String)
On Error GoTo warn:
   If Dir(path) = "" Then MsgBox "file not found!": Exit Sub
   Dim fmax As Long: Dim offset As Long: Dim chunk As Integer
   Dim strdata() As String
   Call initAlpha
   fmax = FileLen(path)
   chunk = 72
   'chunk = 3024 'MUST be a multiple of 3 & 72! (is byte segments to process)
   offset = 1   'each 3024 segment takes about 1.5 secs to parse then doevents
   'this chunk size size means i am manipulating 25,000 characters in memory per cycle !!!
   
   If fmax < chunk Then
        strdata = b64EncodeEngine(binRead(path))
        Call binWrite(strdata, path & "_64enc")
   Else
        divs = Split(calcDivs(fmax, chunk), ",")
        'frmProgress.progbar.Max = divs(0): frmProgress.progbar.Value = 0
        'frmProgress.Show
        Call ketchup(1)
        For i = 1 To divs(0)
          strdata = b64EncodeEngine(binReadSegment(path, offset, chunk))
          If i = 1 Then Call binWrite(strdata, path & "_64enc") _
          Else: Call binAppend(strdata, path & "_64enc")
          offset = offset + chunk
          'frmProgress.progbar.Value = i
          Call ketchup(1)
        Next
        chunk = divs(1) 'last chunk size
        strdata = b64EncodeEngine(binReadSegment(path, offset, chunk))
        Call binAppend(strdata, path & "_64enc")
        'frmProgress.Hide
    End If
Exit Sub
warn: MsgBox "b64encode" & Err.Description
End Sub

Public Sub b64FileDecode(path As String, SaveAs As String)
On Error GoTo warn
'it is imperative that when you parse out the original mime section that
'you MUST REMOVE ALL CRLF AND OTHER WHITE SPACE before you try to decode
'strdata = b64DecodeEngine(binRead(path))
'Call binWrite(strdata, SaveAs)
   
   If Dir(path) = "" Then MsgBox "file not found!": Exit Sub
   Dim fmax As Long: Dim offset As Long: Dim chunk As Integer
   Dim strdata() As String
   Call initAlpha
   fmax = FileLen(path)
   chunk = 72
   'chunk = 3024 'MUST be a multiple of 3 & 72! 'byte segments to process
   offset = 1
   
   If fmax < chunk Then
        strdata = b64DecodeEngine(binRead(path))
        Call binWrite(strdata, SaveAs)
   Else
        divs = Split(calcDivs(fmax, chunk), ",")
        Call ketchup(1)
        For i = 1 To divs(0)
          strdata = b64DecodeEngine(binReadSegment(path, offset, chunk))
          If i = 1 Then Call binWrite(strdata, SaveAs) _
          Else: Call binAppend(strdata, SaveAs)
          offset = offset + chunk
          Call ketchup(1)
        Next
        chunk = divs(1)
        strdata = b64DecodeEngine(binReadSegment(path, offset, chunk))
        Call binAppend(strdata, SaveAs)
    End If
Exit Sub
warn: MsgBox "err in file decode..did you forget to remove all carriage returns andwhite space?" & vbCrLf & vbCrLf & Err.Description
End Sub

Public Function b64EncodeEngine(it)
   On Error GoTo warn
    Dim str As String  'it= BASE 1 string array of characters
    Dim s() As String  'returns BASE 1 string array of chars
    pad = 0            'how many times we had to pad val to encode (1-2)
    
    For i = 1 To UBound(it)  'ascii val-->hex val-->binary string
        it(i) = Hex2Bin(Hex(Asc(it(i))))
    Next
    
    str = Join(it, "")         'has to be div by 6 for now pad with 0's
    While Len(str) Mod 6 <> 0  'in final out put we must represent these as =
      str = str & "00"         'signs b64(64) which is 01000000 binary which cant
      pad = pad + 1            'be represented at this stage! so we must have counter :\
    Wend
    
    ReDim s(Len(str) Mod 6)    'fill s() with 6char div of str
    s = segment(str, 6)        'returns one based array!
                               
    For i = 1 To UBound(s)     'what letter corrosponds to it from
      s(i) = lib.b64Chr(Int("&H" & bin2Hex(s(i))))   'base64 alaphebet
    Next
        
    ReDim Preserve s(UBound(s) + pad)
    For i = 0 To pad - 1       'then remove them before processing?
       s(UBound(s) - i) = "="
    Next
    
    divs = Split(calcDivs(UBound(s), 72), ",") 'wrap characters at 72 chars
    For i = 1 To divs(0)                       'to conform to quoted
      s(i * 72) = s(i * 72) & vbCrLf           'printable standard (has to be < 76)
    Next                                       'but needs to have a multipul of 3 because of our chunked processing
    
    b64EncodeEngine = s 'BASE 1 string arrary of characters
Exit Function
warn: MsgBox "Err in B64EncodeEngine. This function accepts (and returns) only base 1 string arrays of individual characters" & vbCrLf & vbCrLf & Err.Description
End Function

Public Function b64DecodeEngine(it)
 On Error GoTo warn
  Dim s() As String        'it = BASE 1 string array of characters
  ReDim s(1 To UBound(it)) 'returns BASE 1 string array of characters
  Dim str As String
  pad = 0
  warning = " It appears this file was damaged in transport or not" & Chr(13) _
        & "or not correctly encoded, there is a good chance the" & Chr(13) _
        & "data is still viable."
  
  'should this be it(ubound(it)-i) ?
  For i = 0 To 1                        'only last two vals could be pads
    If it(UBound(it) - i) = "=" Then      'if it is a pad then we will have to
      pad = pad + 1                     'remove it from the array and then
    End If                              'remove 2*pad bits from the binary
  Next                                  'stream latter(added 2bits/pad before)
  
  For i = 1 To UBound(it) - pad         'get base64 Ascii(val) of each char
    s(i) = Hex2Bin(Hex(b64Asc(it(i))))  'convert it to hex --> binary
    s(i) = Right(s(i), 6)               'only want last 6 chars, since 64 is
  Next                                  'max dec. value possible, are only trimming 0's
  
  str = Join(s, "")
  If pad Then str = Mid(str, 1, (Len(str) - pad * 2)) 'each pad effectivly adds 2 bits to stream
  'If Len(str) Mod 8 <> 0 Then MsgBox warning         'now we get to remove them :)
  While Len(str) Mod 8 <> 0                           'rember was padded to encode properly
    str = str & "0"                                   'if it isnt encoded right we can usally salvage
  Wend                                                'and decodes with a max of 2 null chrs on end
  
  ReDim s(Len(str) Mod 8)    'clears contents redims to new size
  s = segment(str, 8)        'returns base 1 array of strings 8chr per
                             
  For i = 1 To UBound(s)
     s(i) = Chr(Int("&H" & bin2Hex(s(i))))
  Next
            
  b64DecodeEngine = s 'returns BASE 1 string array of characters
Exit Function
warn: MsgBox "Err in B64DecodeEngine. This function accepts(and returns) only base 1 string arrays of individual characters" & vbCrLf & vbCrLf & Err.Description
End Function

Private Function segment(str As String, div As Integer)
    Dim t() As String       'returns BASE 1 STRING ARRAY of str
    ReDim t(1)              'broken up w/ div characters per element
                            'make sure it is even divisible before!!
    
    For i = 1 To Len(str) Step div
        t(UBound(t)) = Mid(str, i, div)
        If i < Len(str) - div Then         'you dont know how much debugging it
           ReDim Preserve t(UBound(t) + 1)  'took to find why I was always getting
        End If                              '1 extra null byte! ubound(t) was returning
    Next                                    '1 to many elements! last one was null :0`~_

    segment = t
End Function

Private Function calcDivs(maxsz As Long, division As Integer) As String
        sz = maxsz  'using maxsz directly changed its val in calling fx!!
        tmp = 0     'returns (max_whole_divisions,remainder)
        While sz >= division
           sz = sz - division
           tmp = tmp + 1
        Wend
        calcDivs = tmp & "," & sz
End Function

Private Sub ketchup(revs As Integer)
    For i = 0 To revs * 2
      DoEvents
    Next
End Sub

'************   base conversion *****************
Private Function bin2Hex(it As String) As String
  Dim t() As String    'it=binary val as string
  Short = 8 - Len(it)  'returns 2 chr hex string
  ReDim t(3)           'because segment is base1
  
  If Short Then         'need 8 char string to test
     For j = 1 To Short 'pad front with nulls
         it = "0" & it  '(doesnt change value)
     Next
  End If
  
  t = segment(it, 4) 'segment returns base 1 array
  For i = 1 To 2
    For j = 0 To 15
       If t(i) = lib.binHex(j) Then t(i) = lib.hexChr(j): Exit For
    Next
  Next

  bin2Hex = Join(t, "")
End Function

Private Function Hex2Bin(it As String) As String
  Dim tmp As String  'it = 2 char hex string
  If Len(it) = 1 Then it = "0" & it 'need 01 not 1 for val=1
   For i = 1 To 2
      ch = Mid(it, i, 1)
      If IsNumeric(ch) Then
        tmp = tmp & lib.binHex(ch)
      Else
        tmp = tmp & lib.binHex((Asc(ch) - 65 + 10))
      End If      'chr A--> asc65 -->hex chr 10
    Next
  Hex2Bin = tmp
End Function

Private Function b64Asc(it) As Integer
   start = Asc(it)
   If start > 64 And start < 91 Then
      start = 0
   ElseIf start > 96 And start < 123 Then
      start = 26
   Else
      start = 52
   End If
   
   For i = start To 64
     If InStr(1, lib.b64Chr(i), it, vbBinaryCompare) > 0 Then
        b64Asc = i
        Exit For
     End If
   Next
End Function

'**************  end base conversion *************


'************** binary manipulations *************

Private Sub binWrite(it, pth As String)
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

Private Function binAppend(it, pth As String) As Integer
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

Private Function binRead(pth As String) 'returns 0 if !path
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

Private Function binInsert(it, offset As Integer, pth As String)
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

Private Function binReadSegment(pth As String, offset As Long, length As Integer)
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

Public Function prepMIME(pthOrStr As String, prepFile As ArgType)
        
    If prepFile = isPath Then
        strdata = Join(binRead(pthOrStr), "")
    Else
        strdata = pthOrStr
    End If
    
    strdata = Replace(strdata, vbCrLf, "")
    strdata = Replace(strdata, vbCr, "")
    strdata = Replace(strdata, vbLf, "")
    strdata = Replace(strdata, " ", "")
    
    If prepFile = isPath Then
      f = FreeFile
      Open pthOrStr For Input As f
      Print strdata
      Close f
    Else
      prepMIME = strdata
    End If
End Function

'************ end binary manipulations  **************


Private Sub initAlpha()
  With lib
    'b64Alaphabet array
    For i = 0 To 25 '1-25 --> A-Z
     .b64Chr(i) = Chr(65 + i)
    Next
    For i = 26 To 51 '26-51 --> a-z
     .b64Chr(i) = Chr(97 + (i - 26))
    Next
    For i = 0 To 9   '52-61 --> 0-9
     .b64Chr(52 + i) = i
    Next
    .b64Chr(62) = "+"
    .b64Chr(63) = "/"
    .b64Chr(64) = "=" 'since orig val mod 3 must =0 these are the pads

    'hex-->binary array
    .binHex(0) = "0000":   .binHex(8) = "1000"
    .binHex(1) = "0001":   .binHex(9) = "1001"
    .binHex(2) = "0010":   .binHex(10) = "1010"
    .binHex(3) = "0011":   .binHex(11) = "1011"
    .binHex(4) = "0100":   .binHex(12) = "1100":
    .binHex(5) = "0101":   .binHex(13) = "1101"
    .binHex(6) = "0110":   .binHex(14) = "1110":
    .binHex(7) = "0111":   .binHex(15) = "1111"
    
    'hex alaphebet by index array
    .hexChr(0) = "0": .hexChr(6) = "6":  .hexChr(11) = "B"
    .hexChr(1) = "1": .hexChr(7) = "7":  .hexChr(12) = "C"
    .hexChr(2) = "2": .hexChr(8) = "8":  .hexChr(13) = "D"
    .hexChr(3) = "3": .hexChr(9) = "9":  .hexChr(14) = "E"
    .hexChr(4) = "4": .hexChr(10) = "A": .hexChr(15) = "F"
    .hexChr(5) = "5"
  End With
End Sub

