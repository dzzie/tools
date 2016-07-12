Attribute VB_Name = "Module1"
'this function was made by Sebastian and is available on PSC
'Most gracious thanks for the posting...


Public Function Base64Decode(Basein As String) As String
On Error GoTo err
    Dim counter As Integer
    Dim Temp As String
    'For the dec. Tab
    Dim DecodeTable As Variant
    Dim Out(2) As Byte
    Dim inp(3) As Byte
    'DecodeTable holds the decode tab
    DecodeTable = Array("255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "62", "255", "255", "255", "63", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "255", "255", "255", "64", "255", "255", "255", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", _
    "18", "19", "20", "21", "22", "23", "24", "25", "255", "255", "255", "255", "255", "255", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255" _
    , "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255", "255")
    'Reads 4 Bytes in and decrypt them


    For counter = 1 To Len(Basein) Step 4
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '     !!!!!!!!!!!!!!!!!!!
        '!IF YOU WANT YOU CAN ADD AN ERRORCHECK:
        '     !
        '!If DecodeTable()=255 Then Error!!
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '     !!!!!!!!!!!!!!!!!!!
        '4 Bytes in -> 3 Bytes out
        inp(0) = DecodeTable(Asc(Mid$(Basein, counter, 1)))
        inp(1) = DecodeTable(Asc(Mid$(Basein, counter + 1, 1)))
        inp(2) = DecodeTable(Asc(Mid$(Basein, counter + 2, 1)))
        inp(3) = DecodeTable(Asc(Mid$(Basein, counter + 3, 1)))
        Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
        Out(1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
        Out(2) = ((inp(2) And &H3) * 64) Or inp(3)
        '* look for "=" symbols


        If inp(2) = 64 Then
            'If there are 2 characters left -> 1
            '     binary out
            Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
            Temp = Temp & Chr(Out(0) And &HFF)
        ElseIf inp(3) = 64 Then
            'If there are 3 characters left -> 2
            '     binaries out
            Out(0) = (inp(0) * 4) Or ((inp(1) \ 16) And &H3)
            Out(1) = ((inp(1) And &HF) * 16) Or ((inp(2) \ 4) And &HF)
            Temp = Temp & Chr(Out(0) And &HFF) & Chr(Out(1) And &HFF)
        Else 'Return three Bytes
            Temp = Temp & Chr(Out(0) And &HFF) & Chr(Out(1) And &HFF) & Chr(Out(2) And &HFF)
        End If
    Next
    Base64Decode = Temp
Exit Function
err:
  MsgBox "Error Decoding Authentication String : ("
  Base64Decode = "Error"
End Function

'Private Sub Decode_Click()
 '   Base64 needs x * 4 Bytes to work...
 '   If Base64 <> "" And (Len(Base64) Mod 4) = 0 Then
 '     Binary.Text = Base64Decode(Base64.Text)
 '  End If
'End Sub
