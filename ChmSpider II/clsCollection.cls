VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'neither collections or property bags are built with what i would
'want in a collection object so here is my implementation of one
'
'Author:dzzie@yahoo.com
'Sight: http://geocities.com/dzzie

Option Explicit

'keys, values, misc data
Private k(), v(), m()
Private nextIndex As Long
Private myname As String
Private db As Boolean

Enum dumptype
    KEYS = 0
    MISC_ = 1
    VALS = 2
    ALL_ = 3
End Enum
    
Property Let Name(it As String)
 'for debug.print output
 myname = it
End Property

Property Let DebugOutput(t As Boolean)
    db = t
End Property


Property Get Length() As Long
    If aryIsEmpty(k) Then Length = -1 _
    Else Length = UBound(k)
End Property

Function GetNextKey()
    Dim i As Long
    
    If aryIsEmpty(k) Then Exit Function
    If nextIndex > UBound(k) Then nextIndex = 0
    For i = nextIndex To UBound(k)
        If k(i) <> "" Then
            GetNextKey = k(i)
            nextIndex = i + 1
            Exit Function
        End If
    Next
End Function

Function Add(key, value, Optional misc = "") As Boolean
    If keyExists(key) Then GoTo nope
    
    push k(), key
    push v(), value
    push m(), misc
    Add = True
        
    'Debug.Print myname & "  Added Key: " & key
    Exit Function
nope: Add = False
End Function

Function ListDump(which As dumptype) As String
    Select Case which
        Case KEYS: ListDump = Join(k, vbCrLf)
        Case VALS: ListDump = Join(v, vbCrLf)
        Case MISC_: ListDump = Join(m, vbCrLf)
        Case ALL_:
            Dim ret(), i As Long
            If aryIsEmpty(k) Then Exit Function
            For i = 0 To UBound(k)
                push ret(), k(i) & " , " & v(i) & " , " & m(i)
            Next
            ListDump = Join(ret, vbCrLf)
    End Select
End Function

Function GetVal(key)
    If keyExists(key) Then GetVal = v(keyIndex(key))
End Function

Function SetVal(key, value, Optional misc = "") As Boolean
    Dim i As Long
    If keyExists(key) Then
        i = keyIndex(key)
        v(i) = value
        If misc <> "" Then m(i) = misc
        SetVal = True
    End If
End Function

Function SetMiscData(key, misc) As Boolean
    Dim i As Long
    If keyExists(key) Then
        i = keyIndex(key)
        m(i) = misc
        SetMiscData = True
    End If
End Function

Function GetMiscData(key) As Variant
    Dim i As Long
    If keyExists(key) Then
        i = keyIndex(key)
        GetMiscData = m(i)
    End If
End Function

Sub Remove(key)
    Dim i As Long
    If keyExists(key) Then
        i = keyIndex(key)
        k(i) = ""
        v(i) = ""
        m(i) = ""
        'Debug.Print myname & "  Remed Key: " & key
    End If
End Sub

Sub RemoveEmptyElements()
    skinny k()
    skinny v()
    skinny m()
End Sub

Sub ResetCollection()
    Erase k()
    Erase v()
    Erase m()
End Sub

Function keyIndex(key) As Long
    Dim i As Long
    If aryIsEmpty(k) Then Exit Function
    keyIndex = -1
    For i = 0 To UBound(k)
        If k(i) = key Then
            keyIndex = i
            Exit For
        End If
    Next
    If db Then Debug.Print myname & " :KeyIndex : " & key & " : " & keyIndex
End Function

Function keyExists(key) As Boolean
    If aryIsEmpty(k) Then Exit Function
    If keyIndex(key) >= 0 Then keyExists = True
    If db Then Debug.Print myname & " :KeyExists : " & key & " : " & keyExists
End Function

Function keyIsUnique(key) As Boolean
     If aryIsEmpty(k) Then keyIsUnique = True: Exit Function
     If keyIndex(key) < 0 Then keyIsUnique = True
End Function

Function keyFromValue(value)
    Dim i As Long
    For i = 0 To UBound(v)
        If value = v(i) Then keyFromValue = k(i): Exit Function
    Next
End Function

Function keyFromMisc(misc)
    Dim i As Long
    For i = 0 To UBound(m)
        If misc = m(i) Then keyFromMisc = k(i): Exit Function
    Next
End Function

Function GenUniqueKey() As String
    Dim tmp As String
    tmp = RandomNum & "a"
    While Not keyIsUnique(tmp)
        tmp = RandomNum & "a"
    Wend
    GenUniqueKey = tmp
End Function


Private Function aryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim X As Long
    X = UBound(ary)
    aryIsEmpty = False
  Exit Function
oops: aryIsEmpty = True
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim X As Long
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub skinny(ary, Optional base As Long = 0) 'remove empty elements
    Dim ret() 'return adjustable base array
    Dim c As Long, i As Long
    c = base
    For i = base To UBound(ary)
        If ary(i) <> "" Then
            ReDim Preserve ret(c)
            ret(c) = ary(i)
            c = c + 1
        End If
    Next
    ary = ret() 'parent ary obj modified and passed back
End Sub

Function RandomNum()
    Dim tmp
    Randomize
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = Right(tmp, 8)
End Function

