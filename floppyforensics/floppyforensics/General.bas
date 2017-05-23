Attribute VB_Name = "General"
Global Const FLOP_IMG_SIZE = 1474560

Private Type Track
    Data(0 To 18431) As Byte
End Type

Sub ImageFloppy(saveAs)
    If saveAs = Empty Then MsgBox "No path exiting": Exit Sub
    Dim b(0 To 79) As Track
    
    On Error GoTo oops
    
    If MsgBox("Make sure disk you want to image is in drive A:", vbOKCancel) = vbCancel Then Exit Sub
    
    f = FreeFile
    Open "\\.\A:" For Binary Access Read As f
        For i = 0 To 79
            Get f, , b(i)
        Next
    Close f

    Open saveAs For Binary Access Write As f
        For j = 0 To 79
            Put f, , b(j)
        Next
    Close f
   
    MsgBox "Copy Complete saved as " & saveAs

Exit Sub
oops: MsgBox Err.Description & vbCrLf & vbCrLf & "Debug output: Track=" & i, vbExclamation
      Close f
End Sub

Sub SaveImageToFloppy(path)
    If path = Empty Then MsgBox "No path exiting": Exit Sub
    Dim b(0 To 79) As Track
    
    On Error GoTo oops
    
    If MsgBox("Make sure disk you want to transfer image to is in drive A:", vbOKCancel) = vbCancel Then Exit Sub
    If MsgBox("Warning all data on floppy disk in drive A: will be lost!", vbOKCancel) = vbCancel Then Exit Sub
    
    f = FreeFile
    Open path For Binary Access Read As f
        For j = 0 To 79
            Get f, , b(j)
        Next
    Close f
    
    Open "\\.\A:" For Binary Access Write As f
        For i = 0 To 79
            Put f, , b(i)
        Next
    Close f
   
    MsgBox "Image successfully transfered to floppy!", vbInformation

Exit Sub
oops: MsgBox Err.Description & vbCrLf & vbCrLf & "Debug output: Track=" & i, vbExclamation
      Close f
End Sub
    
Function ReadTrack(trackx As Long, length As Integer) As String
    On Error GoTo shit
    If trackx > 79 Or trackx < 1 Or length < 1 Or length > 79 Then MsgBox "Ughh bad call": Exit Function
    
    Dim b() As Track
    ReDim b(length)
    
    If trackx > 1 Then trackx = trackx * 18431
    
    MsgBox "Insert Floppy to image", vbInformation
    
    f = FreeFile
    Open "\\.\A:" For Binary Access Read As f
        For i = 0 To length
            Get f, trackx, b(i)
        Next
    Close f

    Open App.path & "\flp.tmp" For Binary Access Write As f
        For j = 0 To length
            Put f, trackx, b(j)
        Next
    Close f
    
Exit Function
shit: MsgBox Err.Description, vbCritical
End Function

Sub WriteTrack(trackx As Long, length As Integer)
    On Error GoTo shit
    If trackx > 79 Or trackx < 1 Or length < 1 Or length > 79 Then MsgBox "Ughh bad call": Exit Sub
    
    Dim b() As Track
    ReDim b(length)
    
    If trackx > 1 Then trackx = trackx * 18431
    
    MsgBox "Insert Floppy to transfer tracks image onto", vbInformation
    
    f = FreeFile
    Open App.path & "\flp.tmp" For Binary Access Read As f
        For i = 0 To length
            Get f, trackx, b(i)
        Next
    Close f

    Open "\\.\A:" For Binary Access Write As f
        For j = 0 To length
            Put f, trackx, b(j)
        Next
    Close f
    
Exit Sub
shit: MsgBox Err.Description, vbCritical
End Sub


Function ParseFatEntries(t, Optional GenterateComments = True)
    Dim files(), tmp() As String
    On Error GoTo oops
    a = Replace(t, vbCrLf, " ")
    a = Replace(Trim(a), "  ", " ")
    tmp = Split(a, " ")
    
    If Slice(tmp, 0, 2) <> "F0 FF FF" Then
        MsgBox "Could not locate first 3 reserved bytes of FAT table.This can parse the FAT from any offset, just make sure you are on a 3 byte boundry from offset 200h. If what you are trying to parse is not the FAT table, these results will be nonsense.", vbInformation
    End If
    
    Dim b1 As String, b2 As String, b3 As String, list As String
    'heres the main logic of the code to unpack the entries
    For i = 0 To UBound(tmp) Step 3
        b1 = tmp(i)
        If i + 1 <= UBound(tmp) Then b2 = tmp(i + 1)
        If i + 2 <= UBound(tmp) Then b3 = tmp(i + 2)
        
        push files(), LoWord(b2) & b1
        push files(), b3 & HiWord(b2)
    Next
    
    If Not GenterateComments Then GoTo output
    'every thing below here to add comments to it
    Dim InFile As Boolean, newFile As Boolean, dirEntry As Long
    Dim report(), cChain As String, reportedEmpty As Boolean
    
    For i = 0 To UBound(files)
        If i < 2 Then
            push report(), files(i) & "  -  Reserved Entry"
        Else
           If files(i) = "FFF" Then
                If InFile Then
                    push report(), cChain & vbCrLf & String(75, "-")
                    push report(), "FFF - End Of MultiCluster File"
                Else
                    push report(), "FFF  -  New File: Single Cluster Directory Entry: " & dirEntry
                End If
                newFile = True: InFile = False
                dirEntry = dirEntry + 1
           ElseIf files(i) = "000" Then
                If Not reportedEmpty Then
                    dirEntry = dirEntry + 1
                    push report(), "000  -  Empty Area"
                    reportedEmpty = True
                End If
           ElseIf newFile Then
                cChain = files(i) & ","
                push report(), vbCrLf & "New File: MultiCluster; Directory Entry: " & dirEntry & ",  Cluster Chain Follows:" & vbCrLf & String(75, "-")
                newFile = False: InFile = True: reportedEmpty = False
           Else
                cChain = cChain & files(i) & ","
           End If
        End If
    Next
    
    ParseFatEntries = Join(report, vbCrLf)
    
Exit Function
output: ParseFatEntries = Join(files, vbCrLf)
Exit Function
oops: MsgBox Err.Description, vbCritical
End Function

Function LoWord(it)
    LoWord = Right(it, 1)
End Function

Function HiWord(it)
    HiWord = Left(it, 1)
End Function

Function parseDirectoryEntry(shortEntry As String, lfnEntry) As String 'lfnentry=array
    
    If Not AryIsEmpty(lfnEntry) Then LFN = parseLfnEntry(lfnEntry)
    
    Dim t()
    it = Split(Replace(shortEntry, "  ", " "), " ")
    If UBound(it) <> 31 Then Text1 = Join(it, " "): MsgBox UBound(it) '"ughh not 32 byte entry :(": Exit Function
    push t(), "Short File name: " & bytesToAscii(Slice(it, 0, 7)) & IIf(it(0) = "E5", " <-- DELETED", Empty)
    push t(), "Extension: " & bytesToAscii(Slice(it, 8, 10))
    push t(), "Attributes: " & it(11) & "h --> " & Hex2Bin(CStr(it(11)))
    push t(), "Reserved: " & Slice(it, 12, 21)
    push t(), "File Time: " & Slice(it, 22, 23)
    push t(), "File Date: " & Slice(it, 24, 25)
    push t(), "Start Cluster: " & byteswap(Slice(it, 26, 27))
    push t(), "File Length: " & GetFileSize(Slice(it, 28, 31))
    If LFN <> Empty Then push t(), "Long Filename Data: " & vbCrLf & LFN
    
    parseDirectoryEntry = vbCrLf & Join(t, vbCrLf) & vbCrLf
    
End Function

Function GetFileSize(it) As String
    '89 45 07 00 -> 00 07 45 89 -> Clng(from hex)
    t = Split(it, " ")
    For i = 3 To 0 Step -1
       ret = ret & t(i)
    Next
    bytes = CLng("&H" & ret)
    If bytes = 0 Then GetFileSize = 0 & " bytes": Exit Function
    clusters = bytes / 512
    remainder = bytes Mod 512
    tips = 512 - remainder
    If InStr(clusters, ".") > 0 Then clusters = Mid(clusters, 1, InStr(clusters, ".")) + 1
    GetFileSize = bytes & " bytes" & vbCrLf & _
                  "Clusters Used: " & clusters & vbCrLf & _
                  IIf(tips <> 512, "Cluster Tip Size: " & tips & " bytes", Empty)
End Function

Function parseLfnEntry(it)
     t = Join(it, " ")
     parseLfnEntry = HexDump(Split(t, " "), &H2600)
End Function
