VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "typedef Converter"
   ClientHeight    =   10605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "enum2Str"
      Height          =   375
      Left            =   10440
      TabIndex        =   13
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8925
      TabIndex        =   12
      Top             =   480
      Width           =   1185
   End
   Begin VB.CheckBox chkConDef 
      Caption         =   "Convert #define list to enum:"
      Height          =   285
      Left            =   6525
      TabIndex        =   11
      Top             =   435
      Width           =   2430
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy"
      Height          =   390
      Left            =   11640
      TabIndex        =   10
      Top             =   0
      Width           =   795
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Copy on convert"
      Height          =   195
      Left            =   9120
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Use pidl = Long replacements"
      Height          =   315
      Left            =   6525
      TabIndex        =   8
      Top             =   105
      Value           =   1  'Checked
      Width           =   2445
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Use oleexp UUID and PKEY replacements"
      Height          =   210
      Left            =   2850
      TabIndex        =   7
      Top             =   465
      Value           =   1  'Checked
      Width           =   3540
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Include comments"
      Height          =   240
      Left            =   2850
      TabIndex        =   6
      Top             =   120
      Value           =   1  'Checked
      Width           =   1620
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Remove tag from tagName"
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   510
      Value           =   1  'Checked
      Width           =   2880
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Private"
      Height          =   285
      Index           =   1
      Left            =   975
      TabIndex        =   4
      Top             =   90
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Public"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   135
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Default         =   -1  'True
      Height          =   360
      Left            =   10920
      TabIndex        =   2
      Top             =   480
      Width           =   1560
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   5640
      Width           =   12420
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   840
      Width           =   12540
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'typedef struct tagNMTVGETINFOTIPW
'{
'    NMHDR hdr;
'    LPWSTR pszText;
'    int cchTextMax;
'    HTREEITEM hItem;
'    LPARAM lParam;
'} NMTVGETINFOTIPW, *LPNMTVGETINFOTIPW;
Private Sub Command1_Click()
    Dim sTest As String
    Text1 = FixNewLineChars(Text1)
    sTest = Left(Text1.Text, InStr(Text1.Text, vbCrLf))
    If InStr(sTest, "#define") Then
        IsDefine
    ElseIf InStr(sTest, " struct ") Then
        IsStruct
    Else
        IsEnum
    End If
    On Error Resume Next
    If Check5.value = vbChecked Then
        Clipboard.Clear
        Clipboard.SetText Text2.Text
    End If
End Sub

Function FixNewLineChars(ByVal Text As String) As String

Dim LineFeed As String * 1
Dim CarrigeReturn As String * 1
Dim BeepChar As String * 1

LineFeed = Chr(10)
CarrigeReturn = Chr(13)
BeepChar = Chr(7)
Text = Replace(Text, vbCrLf, BeepChar)
Text = Replace(Text, LineFeed, BeepChar)
Text = Replace(Text, CarrigeReturn, BeepChar)
Text = Replace(Text, LineFeed & CarrigeReturn, BeepChar)

FixNewLineChars = Replace(Text, BeepChar, vbCrLf)
End Function

Private Sub IsDefine()
Dim sIn() As String
Dim sOut() As String
Dim i As Long, j As Long
Dim s1 As String, s2 As String, s3 As String
Dim sC As String
Text2.Text = ""
sIn = Split(Text1.Text, vbCrLf)
ReDim sOut(UBound(sIn))
For i = 0 To UBound(sIn)
    s1 = Trim$(sIn(i))
    s1 = Replace$(s1, vbTab, "")
    
    If Left$(s1, 2) = "//" Then 'whole line is comment
        If Check2.value = vbChecked Then
            sOut(i) = "' " & Mid$(s1, 3)
            GoTo nxt
        End If
    End If
    If InStr(s1, "//") Then
        sC = " '" & Mid$(s1, InStr(s1, "//") + 2) 'store comment before removing; add in later if wanted
        s1 = Left$(s1, InStr(s1, "//") - 1)
    End If
    For j = 20 To 2 Step -1
        s1 = Replace$(s1, Space$(j), " ")
    Next
    If Len(s1) > 2 Then
        s2 = Mid$(s1, InStr(s1, " ") + 1)
        s3 = Mid$(s2, InStr(s2, " ") + 1)
        s2 = Left$(s2, Len(s2) - Len(s3))
        
        
        If chkConDef.value = vbChecked Then
            sOut(i) = s2 & " = " & s3
        Else
            sOut(i) = "Public Const " & s2 & " = " & s3
        End If
        sOut(i) = Replace$(sOut(i), "0x", "&H")
        sOut(i) = Replace$(sOut(i), "|", " Or ")
        If Check2.value = vbChecked Then
            sOut(i) = sOut(i) & sC
        End If
        If chkConDef.value = vbChecked Then
            sOut(i) = vbTab & sOut(i)
        End If
    End If
nxt:
Next

If chkConDef.value = vbChecked Then
    Text2.Text = IIf(Option1(0).value = True, "Public", "Private") & " Enum " & Text3.Text & vbCrLf
End If
For i = 0 To UBound(sOut)
    If Len(sOut(i)) > 2 Then
        Text2.Text = Text2.Text & sOut(i) & vbCrLf
    End If
Next i
If chkConDef.value = vbChecked Then
    Text2.Text = Text2.Text & "End Enum"
End If
End Sub


Private Sub IsEnum()
    
    Dim sIn() As String
    Dim sOut() As String
    Dim i As Long
    Dim j As Long, k As Long
    Dim s1 As String, s2 As String, s3 As String
    Dim sC As String
    Dim sB As String, nB As Long
    
    Text2.Text = ""
    'If InStr(Text1.Text, ",") > 0 Then
    '    Text1.Text = Replace(Text1, ",", vbCrLf) 'multiple entries per line
    'End If
    
    sIn = Split(Text1.Text, vbCrLf)
    ReDim sOut(UBound(sIn))
    s1 = sIn(0)
    
    If InStr(s1, "//") Then
        sC = " '" & Mid$(s1, InStr(s1, "//") + 2)
        s1 = Left$(s1, InStr(s1, "//") - 1)
    End If
    
    s1 = Trim(s1)
    s1 = Replace(s1, " {", "")
    s1 = Right$(s1, Len(s1) - InStrRev(s1, " "))
    
    If Left$(s1, 1) = "_" Then s1 = Mid$(s1, 2) 'cant start with underscore
    If Left$(s1, 3) = "tag" Then s1 = Mid$(s1, 4)
    
    sOut(0) = IIf(Option1(0).value = True, "Public", "Private") & " Enum " & s1
    If Check2.value = vbChecked Then
        sOut(0) = sOut(0) & sC
    End If
    
    s2 = sIn(1)
    s2 = Trim(s2)
    s2 = Replace(s2, vbTab, "")
    
    If s2 = "{" Then
        j = 2
    Else
        j = 1
    End If
    
    For i = j To UBound(sOut) - 1
        sC = ""
        s1 = sIn(i)
        s1 = Trim$(s1)
        s1 = Replace(s1, vbTab, "")
        
        If Left$(s1, 2) = "//" Then
            'whole line is comment
            If Check2.value = vbChecked Then
                sOut(i) = vbTab & "' " & Mid$(s1, 3)
                GoTo nxt
            End If
        End If
        
        If InStr(s1, "//") Then
            sC = " '" & Mid$(s1, InStr(s1, "//") + 2) 'store comment before removing; add in later if wanted
            s1 = Left$(s1, InStr(s1, "//") - 1)
        End If
        
        If Len(s1) > 2 Then
            s1 = Replace(s1, " ", "")
            s1 = Replace(s1, ",", "")
            
            If InStr(s1, "=") Then
                'If Right$(s1, 2) = "=0" Then nB = 1
                'If Right$(s1, 2) = "=1" Then nB = 2 'number base for enum without =, if 0 is defined, next is 1 etc
                                                    'i haven't ever seen a sequential enum that didn't start with 0/1
                                                    'so didn't add the logic for an arbitrary number; change if needed
                s1 = Replace(s1, "=0x", " = &H")
                s1 = Replace$(s1, "|", " Or ")
            Else
                'we dont need to autonumber vb supports that internally....bugfix also no need for k
                's1 = s1 & " = " & (nB)   'if nB is defined, the first element has been defined
                '                         'and we should start numbering after that
                'nB = nB + 1
            End If
            
            s1 = Replace(s1, "*", "")
            sOut(i) = vbTab & s1
            
            If Check2.value = vbChecked Then
                sOut(i) = sOut(i) & sC
            End If
            
        End If
nxt:
    Next i
            
    Dim ret() As String
    Dim longest As Long
    Dim c As Long
    
    For i = 0 To UBound(sOut)
        If Len(sOut(i)) > 2 Then
            c = InStr(sOut(i), "'")
            If c > longest Then longest = c
        End If
    Next i
    
    For i = 0 To UBound(sOut)
        If Len(sOut(i)) > 2 Then
            c = InStr(sOut(i), "'")
            If c > 0 And c < longest Then
                sOut(i) = Mid(sOut(i), 1, c - 1) & Space(longest - c) & Mid(sOut(i), c)
            End If
            push ret, sOut(i)
        End If
    Next
    
    Text2.Text = Join(ret, vbCrLf) & vbCrLf & "End Enum"
        
End Sub



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub IsStruct()
    Dim sIn() As String
    Dim sOut() As String
    Dim i As Long
    Dim j As Long
    Dim s1 As String, s2 As String, s3 As String, sT As String
    Dim sC As String
    Dim note As String
    
    Text2.Text = ""
    sIn = Split(Text1.Text, vbCrLf)
    ReDim sOut(UBound(sIn))
    
    s1 = sIn(0)
    
    If InStr(s1, "//") Then
        sC = " '" & Mid$(s1, InStr(s1, "//") + 2)
        s1 = Left$(s1, InStr(s1, "//") - 1)
    End If
    
    s1 = Trim(s1)
    s1 = Replace(s1, " {", "")
    s1 = Right$(s1, Len(s1) - InStrRev(s1, " "))
    
    If Left$(s1, 1) = "_" Then s1 = Mid$(s1, 2) 'cant start with underscore
    If Left$(s1, 3) = "tag" Then s1 = Mid$(s1, 4)
    
    sOut(0) = IIf(Option1(0).value = True, "Public", "Private") & " Type " & s1
    
    If Check2.value = vbChecked Then
        sOut(0) = sOut(0) & sC
    End If
    
    s2 = sIn(1)
    s2 = Trim(s2)
    s2 = Replace(s2, vbTab, "")
    
    If s2 = "{" Then
        j = 2
    Else
        j = 1
    End If
    
    For i = j To UBound(sOut) - 1
        sC = ""
        s1 = sIn(i)
        sT = Trim$(sIn(i))
        
        If Left$(sT, 2) = "//" Then
            'whole line is comment
            If Check2.value = vbChecked Then
                sOut(i) = vbTab & "'" & Mid$(Trim$(s1), 3)
                GoTo nxt
            End If
        End If
        
        If InStr(s1, "//") Then
            sC = " '" & Mid$(s1, InStr(s1, "//") + 2)
            s1 = Left$(s1, InStr(s1, "//") - 1)
        End If
        
        s1 = Trim$(s1)
        If Len(s1) > 2 Then
        
            s2 = Right(s1, Len(s1) - InStrRev(s1, " "))
            s3 = Left$(s1, Len(s1) - Len(s2))
            s3 = Trim(s3)
            s3 = Replace(s3, vbTab, "")
            s2 = Left(s2, Len(s2) - 1) 'remove ;
            
            If Right$(s2, 1) = "]" Then 'we have an array
                s2 = DoArray(s2)
            End If
            's3 now contains type name, s2 contains var name
            s3 = s3 & " " 'add space to prevent name-within-name errors
                          'e.g. IMAGELISTDRAWPARAMS contains WPARAM, but isn't replaced, WPARAM is
                          'so it would come out as IMAGELISTDRALongS without this correction
            DoTypeReplace s3, note
            s3 = Trim$(s3)
            s3 = Replace(s3, "*", "")
            's2 = Replace(s2, "*", "") 'error you cant delete this..its different
            sOut(i) = vbTab & s2 & " As " & s3
            
            'If Check2.value = vbChecked Then
                If Len(sC) > 0 Or Len(note) > 0 Then
                    If Len(sC) = 0 Then sC = "'" & note Else sC = sC & note
                    sOut(i) = sOut(i) & sC
                End If
            'End If
            
        End If
nxt:
    Next i
    
    Dim ret() As String
    Dim longest As Long
    Dim c As Long
    
    For i = 0 To UBound(sOut)
        If Len(sOut(i)) > 2 Then
            c = InStr(sOut(i), "'")
            If c > longest Then longest = c
        End If
    Next i
    
    For i = 0 To UBound(sOut)
        If Len(sOut(i)) > 2 Then
            c = InStr(sOut(i), "'")
            If c > 0 And c < longest Then
                sOut(i) = Mid(sOut(i), 1, c - 1) & Space(longest - c) & Mid(sOut(i), c)
            End If
            push ret, sOut(i)
        End If
    Next
    
    Text2.Text = Join(ret, vbCrLf) & vbCrLf & "End Type"
    
End Sub

Private Function DoArray(ByVal sz As String) As String
Dim sLen As String
Dim nLen As Long
Dim sName As String

sName = Left$(sz, (InStr(sz, "[") - 1))
sLen = Replace(sz, sName, "")
sLen = Mid$(sLen, 2)
sLen = Left$(sLen, Len(sLen) - 1)
If IsAllNumbers(sLen) Then
    nLen = CLng(sLen) - 1
    DoArray = sName & "(0 To " & nLen & ")"
Else
    DoArray = sName & "(0 To (" & sLen & " - 1))"
End If
End Function
Private Sub DoTypeReplace(sz As String, Optional ByRef notes As String)
    'Debug.Print "DTR->" & sz
    
    notes = Empty
    
    'when adding your own types, it's important to consider downstream
    'replacements; e.g. make sure LPWORD is replaced before WORD
    sz = Replace(sz, "LPINT ", "Long")
    sz = Replace(sz, "LPCOLORREF ", "Long")
    sz = Replace(sz, "LPHTREEITEM ", "Long")
    sz = Replace(sz, "LPDWORD ", "Long")
    sz = Replace(sz, "LPWORD ", "Long")
    sz = Replace(sz, "LPLONG ", "Long")
    sz = Replace(sz, "LPBOOL ", "Long")
    sz = Replace(sz, "LPHANDLE ", "Long")
    sz = Replace(sz, "LPBYTE ", "Byte")
    
    sz = Replace(sz, "PBOOL ", "Long")
    sz = Replace(sz, "PBYTE ", "Byte")
    sz = Replace(sz, "PCHAR ", "Byte")
    sz = Replace(sz, "PDWORD32 ", "Long")
    sz = Replace(sz, "PDWORD64 ", "Currency")
    sz = Replace(sz, "PDWORD ", "Long")
    sz = Replace(sz, "PDWORDLONG ", "Long")
    sz = Replace(sz, "PDWORD_PTR ", "Long")
    sz = Replace(sz, "PFLOAT ", "Double")
    sz = Replace(sz, "PHANDLE ", "Long")
    sz = Replace(sz, "PHKEY ", "Long")
    sz = Replace(sz, "PINT_PTR ", "Long")
    sz = Replace(sz, "PINT32 ", "Long")
    sz = Replace(sz, "PINT64 ", "Currency")
    sz = Replace(sz, "PDWORD ", "Long")
    sz = Replace(sz, "PLCID ", "Long")
    sz = Replace(sz, "PLONGLONG ", "Long")
    sz = Replace(sz, "PLONG_PTR ", "Long")
    sz = Replace(sz, "PLONG32 ", "Long")
    sz = Replace(sz, "PLONG64 ", "Currency")
    sz = Replace(sz, "PLONG ", "Long")
    sz = Replace(sz, "POINTER_32 ", "Long")
    sz = Replace(sz, "POINTER_64 ", "Currency")
    sz = Replace(sz, "PSHORT ", "Integer")
    
    sz = Replace(sz, "PUINT64 ", "Currency")
    sz = Replace(sz, "PUINT32 ", "Long")
    sz = Replace(sz, "PUINT ", "Long")
    sz = Replace(sz, "PULONGLONG ", "Currency")
    sz = Replace(sz, "PULONG_PTR ", "Long")
    sz = Replace(sz, "PULONG32 ", "Long")
    sz = Replace(sz, "PULONG64 ", "Currency")
    sz = Replace(sz, "PULONG ", "Long")
    sz = Replace(sz, "PUSHORT ", "Integer")
    sz = Replace(sz, "PWORD ", "Long")
    
    sz = Replace(sz, "CHAR ", "Byte")
    sz = Replace(sz, "LPARAM ", "Long")
    sz = Replace(sz, "UINT_PTR ", "Long")
    sz = Replace(sz, "UINT64 ", "Currency")
    sz = Replace(sz, "UINT ", "Long")
    sz = Replace(sz, "INT_PTR ", "Long")
    sz = Replace(sz, "INT64 ", "Currency")
    sz = Replace(sz, "int ", "Long")
    sz = Replace(sz, "COLORREF ", "Long")
    sz = Replace(sz, "HTREEITEM ", "Long")
    sz = Replace(sz, "LPSTR ", "String")
    sz = Replace(sz, "BSTR ", "String")
    sz = Replace(sz, "LONG64 ", "Currency")
    sz = Replace(sz, "DWORDLONG ", "Currency")
    sz = Replace(sz, "DWORD_PTR ", "Long")
    sz = Replace(sz, "DWORD32 ", "Long")
    sz = Replace(sz, "DWORD64 ", "Currency")
    sz = Replace(sz, "DWORD ", "Long")
    sz = Replace(sz, "WORD ", "Long")
    sz = Replace(sz, "ULONG_PTR ", "Long")
    sz = Replace(sz, "LONG_PTR ", "Long")
    sz = Replace(sz, "LCID ", "Long")
    sz = Replace(sz, "HWND ", "Long")
    sz = Replace(sz, "HDC ", "Long")
    sz = Replace(sz, "HIMAGELIST ", "Long")
    sz = Replace(sz, "LPARAM ", "Long")
    sz = Replace(sz, "BOOL ", "Long")
    sz = Replace(sz, "ULONGLONG ", "Currency")
    sz = Replace(sz, "HBITMAP ", "Long")
    sz = Replace(sz, "HMENU ", "Long")
    sz = Replace(sz, "HKEY ", "Long")
    sz = Replace(sz, "HICON ", "Long")
    sz = Replace(sz, "HBRUSH ", "Long")
    sz = Replace(sz, "HCURSOR ", "Long")
    sz = Replace(sz, "HANDLE ", "Long")
    sz = Replace(sz, "HACCEL ", "Long")
    sz = Replace(sz, "HENMETAFILE ", "Long")
    sz = Replace(sz, "HMETAFILE ", "Long")
    sz = Replace(sz, "HMODULE ", "Long")
    sz = Replace(sz, "HPEN ", "Long")
    sz = Replace(sz, "HFONT ", "Long")
    sz = Replace(sz, "HINSTANCE ", "Long")
    sz = Replace(sz, "LPWSTR ", "Long")
    sz = Replace(sz, "LPCTSTR ", "Long")
    sz = Replace(sz, "LPCWSTR ", "Long")
    sz = Replace(sz, "WCHAR", "Integer")
    sz = Replace(sz, "ULONG ", "Long")
    sz = Replace(sz, "WPARAM  ", "Long")
    sz = Replace(sz, "SHORT", "Integer")
    sz = Replace(sz, "FLOAT ", "Double")
    sz = Replace(sz, "LPRECT ", "RECT")
    sz = Replace(sz, "HRESULT ", "Long")
    sz = Replace(sz, "LRESULT ", "Long")
    
    If InStr(1, sz, "unsigned", vbTextCompare) > 0 Then
        sz = Replace(sz, "unsigned", "")
        notes = " UNSIGNED"
    End If
    
    If InStr(1, sz, "uint", vbTextCompare) > 0 Then
        sz = Replace(sz, "uint", "int")
        notes = " UNSIGNED"
    End If
    
    sz = Replace(sz, "int64_t ", "Currency")
    sz = Replace(sz, "int32_t ", "Long")
    sz = Replace(sz, "int16_t ", "Short")
    sz = Replace(sz, "int8_t ", "Byte")


    
    If Check3.value = vbChecked Then
        sz = Replace(sz, "REFIID ", "UUID")
        sz = Replace(sz, "IID ", "UUID")
        sz = Replace(sz, "REFCLSID ", "UUID")
        sz = Replace(sz, "CLSID ", "UUID")
        sz = Replace(sz, "FOLDERTYPEID ", "UUID")
        sz = Replace(sz, "REFGUID ", "UUID")
        sz = Replace(sz, "GUID ", "UUID")
    End If
    
    If Check4.value = vbChecked Then
        sz = Replace(sz, "LPITEMIDLIST ", "Long")
        sz = Replace(sz, "LPCITEMIDLIST ", "Long")
        sz = Replace(sz, "PCIDLIST_ABSOLUTE ", "Long")
        sz = Replace(sz, "PCIDLIST_CHILD ", "Long")
        sz = Replace(sz, "PIDLIST_ABSOLUTE ", "Long")
        sz = Replace(sz, "PIDLIST_CHILD ", "Long")
        sz = Replace(sz, "PCUITEMID_ABSOLUTE ", "Long")
        sz = Replace(sz, "PCUITEMID_CHILD ", "Long")
        sz = Replace(sz, "PITEMID_CHILD ", "Long")
        sz = Replace(sz, "PUITEMID_CHILD ", "Long")
        sz = Replace(sz, "PCUIDLIST_RELATIVE ", "Long")
    End If
    
    
    

End Sub

Private Function IsAllNumbers(ByVal szIn As String) As Boolean
'Returns if a string is a >=0 integer
Dim i As Long, sC As String
Dim r As Boolean

For i = 1 To Len(szIn)
    sC = Mid$(szIn, i, 1)
    If (sC = "0") Or (sC = "1") Or (sC = "3") Or (sC = "4") Or (sC = "5") Or (sC = "6") Or (sC = "7") Or (sC = "8") Or (sC = "9") Or (sC = "2") Then
        r = True
    Else
        r = False
        Exit For
    End If
Next i

IsAllNumbers = r
End Function

Private Sub Command2_Click()
Clipboard.Clear
Clipboard.SetText Text2.Text, vbCFText
End Sub

Private Sub Command3_Click()
    
    'Public Enum x86_avx_bcast
    '    X86_AVX_BCAST_INVALID = 0 ' Uninitialized.
    '    X86_AVX_BCAST_2         ' AVX512 broadcast type {1to2}
    '    X86_AVX_BCAST_4         ' AVX512 broadcast type {1to4}
    '    X86_AVX_BCAST_8         ' AVX512 broadcast type {1to8}
    '    X86_AVX_BCAST_16        ' AVX512 broadcast type {1to16}
    'End Enum
    
    Dim ret() As String
    Dim tmp() As String
    Dim n, t, a, i
    
    tmp = Split(Text1, vbCrLf)
    
    Dim w() As String
    w = Split(tmp(0), " ")
    n = w(UBound(w))
    
    push ret, "function " & n & "2str(v as " & n & ") as string"
    push ret, "     dim r as string"
    
    For i = 1 To UBound(tmp)
        
        t = Trim(tmp(i))
        If VBA.Left(t, 1) = "'" Then t = Mid(t, 2)
        
        t = Trim(Replace(t, vbTab, Empty))
        If LCase(t) = "end enum" Then Exit For
        
        a = InStr(t, "=")
        If a > 0 Then t = Mid(t, 1, a - 1)
        a = InStr(t, "'")
        If a > 0 Then t = Mid(t, 1, a - 1)
        t = Trim(t)
        
        If Len(t) > 0 Then
            push ret, "    if v = " & t & " then r = """ & t & """"
        End If
    Next
    
    push ret, "    if len(r) = 0 then r = ""Unknown: "" & hex(v) "
    push ret, "    " & n & "2str = r"
    push ret, "end function"
    
    Text2 = Join(ret, vbCrLf)
        
End Sub

Private Sub Text1_DblClick()
    Text1 = Clipboard.GetText
    Command1_Click
End Sub
