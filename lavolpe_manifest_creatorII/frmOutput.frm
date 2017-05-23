VERSION 5.00
Begin VB.Form frmOutput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Output Options"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkOption 
      Caption         =   "Do Not Export Empty/Blank Attributes"
      Height          =   270
      Index           =   2
      Left            =   285
      TabIndex        =   7
      Top             =   945
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.ComboBox cboDest 
      Height          =   315
      ItemData        =   "frmOutput.frx":0000
      Left            =   525
      List            =   "frmOutput.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2325
      Width           =   2820
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue..."
      Height          =   420
      Left            =   3540
      TabIndex        =   3
      Top             =   2250
      Width           =   1395
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1620
      Width           =   4410
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Do Not Use Prefixed Name Spaces"
      Height          =   270
      Index           =   1
      Left            =   285
      TabIndex        =   1
      Top             =   570
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Indent Manifest -- Not Recommended for Res files"
      Height          =   270
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   180
      Value           =   1  'Checked
      Width           =   4485
   End
   Begin VB.Label Label1 
      Caption         =   "Language Identifier for Resource Files Only"
      Height          =   315
      Index           =   1
      Left            =   540
      TabIndex        =   6
      Top             =   1335
      Width           =   4155
   End
   Begin VB.Label Label1 
      Caption         =   "Destination"
      Height          =   315
      Index           =   0
      Left            =   525
      TabIndex        =   5
      Top             =   2070
      Width           =   1815
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private m_Manifest As cManifestEx

Friend Sub SetManifest(Manifest As cManifestEx)
    Set m_Manifest = Manifest
End Sub

Private Sub Command1_Click()

    Dim tDoc As DOMDocument60
    Dim cBrowser As UnicodeFileDialog
    
    If cboDest.ListIndex < 2 Then
        Set cBrowser = New UnicodeFileDialog
        With cBrowser
            .Flags = OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_OVERWRITEPROMPT
            .DialogTitle = "Select Destination"
            If cboDest.ListIndex = 0 Then
                .Filter = "Resource Files (*.res)|*.res|manifest|All Files|*.*"
                .DefaultExt = "res"
            Else
                .DefaultExt = "manifest"
                .Filter = "Manifest Files|*.manifest|All Files|*.*"
            End If
            .FilterIndex = 1
        End With
        If cBrowser.ShowSave(frmMain.hWnd) = False Then Exit Sub
        
        If DoesFileExists(cBrowser.FileName, True) = True Then
            ' verify res file; better safe than sorry
            If IsResourceFile(cBrowser.FileName) = True Then
                If cboDest.ListIndex = 1 Then
                    MsgBox "You chose to export mainfest to a manifest file but you selected a resource file. Aborting", vbInformation + vbOKOnly, "No Action Taken"
                    Exit Sub
                End If
            ElseIf cboDest.ListIndex = 0 Then
                MsgBox "You chose to export mainfest to a resource file but did not select one. Aborting", vbInformation + vbOKOnly, "No Action Taken"
                Exit Sub
            End If
        End If
    End If
        
    Set tDoc = m_Manifest.ExportXML(1&, (chkOption(0).Value = vbChecked), (chkOption(1).Value = vbChecked), (chkOption(2).Value = vbChecked))
    If Not tDoc Is Nothing Then
        If tDoc.parseError = 0 Then
            If cboDest.ListIndex = 2 Then
                Clipboard.Clear: Clipboard.SetText tDoc.xml
                MsgBox "Manifest placed on the clipboard", vbInformation + vbOKOnly, "Done"
            Else
                If cboDest.ListIndex = 1 Then
                    On Error Resume Next
                    tDoc.save cBrowser.FileName
                    If Err Then
                        MsgBox Err.Description, vbExclamation + vbOKOnly, "Error Writing File"
                    Else
                        MsgBox "Manifest written to file.", vbInformation + vbOKOnly, "Done"
                    End If
                    On Error GoTo 0
                Else
                    Dim bData() As Byte, sErr As String, lSize As Long
                    bData() = StrConv(tDoc.xml, vbFromUnicode)
                    lSize = UBound(bData) + 1&
                    If (lSize Mod 4) Then ReDim Preserve bData(0 To lSize + (lSize Mod 4) - 1)
                    sErr = InsertManifestToResource(cBrowser.FileName, bData(), cboLanguage.ItemData(cboLanguage.ListIndex), False)
                    If sErr = vbNullString Then
                        MsgBox "Manfiest written to resource file", vbInformation + vbOKOnly, "Done"
                    Else
                        MsgBox sErr, vbExclamation + vbOKOnly, "No Action Taken"
                    End If
                End If
            End If
        End If
    End If
    
ExitRoutine:
    Set tDoc = Nothing
    Unload Me
End Sub

Private Sub Form_Load()

    cboDest.ListIndex = 0
    Call pvSetLanguageID(m_Manifest.LCID)
    
End Sub


Private Sub pvFillLanguageIDs()
    ' http://msdn.microsoft.com/en-us/goglobal/bb964664.aspx
    With cboLanguage
        .AddItem "Afrikaans (South Africa)": .ItemData(.NewIndex) = 1078
        .AddItem "Albanian (Albania)": .ItemData(.NewIndex) = 1052
        .AddItem "Amharic (Ethiopia)": .ItemData(.NewIndex) = 1118
        .AddItem "Arabic (Saudi Arabia)": .ItemData(.NewIndex) = 1025
        .AddItem "Arabic (Algeria)": .ItemData(.NewIndex) = 5121
        .AddItem "Arabic (Bahrain)": .ItemData(.NewIndex) = 15361
        .AddItem "Arabic (Egypt)": .ItemData(.NewIndex) = 3073
        .AddItem "Arabic (Iraq)": .ItemData(.NewIndex) = 2049
        .AddItem "Arabic (Jordan)": .ItemData(.NewIndex) = 11265
        .AddItem "Arabic (Kuwait)": .ItemData(.NewIndex) = 13313
        .AddItem "Arabic (Lebanon)": .ItemData(.NewIndex) = 12289
        .AddItem "Arabic (Libya)": .ItemData(.NewIndex) = 4097
        .AddItem "Arabic (Morocco)": .ItemData(.NewIndex) = 6145
        .AddItem "Arabic (Oman)": .ItemData(.NewIndex) = 8193
        .AddItem "Arabic (Qatar)": .ItemData(.NewIndex) = 16385
        .AddItem "Arabic (Syria)": .ItemData(.NewIndex) = 10241
        .AddItem "Arabic (Tunisia)": .ItemData(.NewIndex) = 7169
        .AddItem "Arabic (U.A.E.)": .ItemData(.NewIndex) = 14337
        .AddItem "Arabic (Yemen)": .ItemData(.NewIndex) = 9217
        .AddItem "Armenian (Armenia)": .ItemData(.NewIndex) = 1067
        .AddItem "Assamese": .ItemData(.NewIndex) = 1101
        .AddItem "Azeri (Cyrillic)": .ItemData(.NewIndex) = 2092
        .AddItem "Azeri (Latin)": .ItemData(.NewIndex) = 1068
        .AddItem "Basque": .ItemData(.NewIndex) = 1069
        .AddItem "Belarusian": .ItemData(.NewIndex) = 1059
        .AddItem "Bengali (India)": .ItemData(.NewIndex) = 1093
        .AddItem "Bengali (Bangladesh)": .ItemData(.NewIndex) = 2117
        .AddItem "Bosnian (Bosnia/Herzegovina)": .ItemData(.NewIndex) = 5146
        .AddItem "Bulgarian": .ItemData(.NewIndex) = 1026
        .AddItem "Burmese": .ItemData(.NewIndex) = 1109
        .AddItem "Catalan": .ItemData(.NewIndex) = 1027
        .AddItem "Cherokee (United States)": .ItemData(.NewIndex) = 1116
        .AddItem "Chinese (People's Republic of China)": .ItemData(.NewIndex) = 2052
        .AddItem "Chinese (Singapore)": .ItemData(.NewIndex) = 4100
        .AddItem "Chinese (Taiwan)": .ItemData(.NewIndex) = 1028
        .AddItem "Chinese (Hong Kong SAR)": .ItemData(.NewIndex) = 3076
        .AddItem "Chinese (Macao SAR)": .ItemData(.NewIndex) = 5124
        .AddItem "Croatian": .ItemData(.NewIndex) = 1050
        .AddItem "Croatian (Bosnia/Herzegovina)": .ItemData(.NewIndex) = 4122
        .AddItem "Czech": .ItemData(.NewIndex) = 1029
        .AddItem "Danish": .ItemData(.NewIndex) = 1030
        .AddItem "Divehi": .ItemData(.NewIndex) = 1125
        .AddItem "Dutch (Netherlands)": .ItemData(.NewIndex) = 1043
        .AddItem "Dutch (Belgium)": .ItemData(.NewIndex) = 2067
        .AddItem "Edo": .ItemData(.NewIndex) = 1126
        .AddItem "English (United States)": .ItemData(.NewIndex) = 1033
        .AddItem "English (United Kingdom)": .ItemData(.NewIndex) = 2057
        .AddItem "English (Australia)": .ItemData(.NewIndex) = 3081
        .AddItem "English (Belize)": .ItemData(.NewIndex) = 10249
        .AddItem "English (Canada)": .ItemData(.NewIndex) = 4105
        .AddItem "English (Caribbean)": .ItemData(.NewIndex) = 9225
        .AddItem "English (Hong Kong SAR)": .ItemData(.NewIndex) = 15369
        .AddItem "English (India)": .ItemData(.NewIndex) = 16393
        .AddItem "English (Indonesia)": .ItemData(.NewIndex) = 14345
        .AddItem "English (Ireland)": .ItemData(.NewIndex) = 6153
        .AddItem "English (Jamaica)": .ItemData(.NewIndex) = 8201
        .AddItem "English (Malaysia)": .ItemData(.NewIndex) = 17417
        .AddItem "English (New Zealand)": .ItemData(.NewIndex) = 5129
        .AddItem "English (Philippines)": .ItemData(.NewIndex) = 13321
        .AddItem "English (Singapore)": .ItemData(.NewIndex) = 18441
        .AddItem "English (South Africa)": .ItemData(.NewIndex) = 7177
        .AddItem "English (Trinidad)": .ItemData(.NewIndex) = 11273
        .AddItem "English (Zimbabwe)": .ItemData(.NewIndex) = 12297
        .AddItem "Estonian": .ItemData(.NewIndex) = 1061
        .AddItem "Faroese": .ItemData(.NewIndex) = 1080
        .AddItem "Farsi": .ItemData(.NewIndex) = 1065
        .AddItem "Filipino": .ItemData(.NewIndex) = 1124
        .AddItem "Finnish": .ItemData(.NewIndex) = 1035
        .AddItem "French (France)": .ItemData(.NewIndex) = 1036
        .AddItem "French (Belgium)": .ItemData(.NewIndex) = 2060
        .AddItem "French (Cameroon)": .ItemData(.NewIndex) = 11276
        .AddItem "French (Canada)": .ItemData(.NewIndex) = 3084
        .AddItem "French (Democratic Rep. of Congo)": .ItemData(.NewIndex) = 9228
        .AddItem "French (Cote d'Ivoire)": .ItemData(.NewIndex) = 12300
        .AddItem "French (Haiti)": .ItemData(.NewIndex) = 15372
        .AddItem "French (Luxembourg)": .ItemData(.NewIndex) = 5132
        .AddItem "French (Mali)": .ItemData(.NewIndex) = 13324
        .AddItem "French (Monaco)": .ItemData(.NewIndex) = 6156
        .AddItem "French (Morocco)": .ItemData(.NewIndex) = 14348
        .AddItem "French (North Africa)": .ItemData(.NewIndex) = 58380
        .AddItem "French (Reunion)": .ItemData(.NewIndex) = 8204
        .AddItem "French (Senegal)": .ItemData(.NewIndex) = 10252
        .AddItem "French (Switzerland)": .ItemData(.NewIndex) = 4108
        .AddItem "French (West Indies)": .ItemData(.NewIndex) = 7180
        .AddItem "Frisian (Netherlands)": .ItemData(.NewIndex) = 1122
        .AddItem "Fulfulde (Nigeria)": .ItemData(.NewIndex) = 1127
        .AddItem "FYRO Macedonian": .ItemData(.NewIndex) = 1071
        .AddItem "Gaelic (Ireland)": .ItemData(.NewIndex) = 2108
        .AddItem "Gaelic (Scotland)": .ItemData(.NewIndex) = 1084
        .AddItem "Galician": .ItemData(.NewIndex) = 1110
        .AddItem "Georgian": .ItemData(.NewIndex) = 1079
        .AddItem "German (Germany)": .ItemData(.NewIndex) = 1031
        .AddItem "German (Austria)": .ItemData(.NewIndex) = 3079
        .AddItem "German (Liechtenstein)": .ItemData(.NewIndex) = 5127
        .AddItem "German (Luxembourg)": .ItemData(.NewIndex) = 4103
        .AddItem "German (Switzerland)": .ItemData(.NewIndex) = 2055
        .AddItem "Greek": .ItemData(.NewIndex) = 1032
        .AddItem "Guarani (Paraguay)": .ItemData(.NewIndex) = 1140
        .AddItem "Gujarati": .ItemData(.NewIndex) = 1095
        .AddItem "Hausa (Nigeria)": .ItemData(.NewIndex) = 1128
        .AddItem "Hawaiian (United States)": .ItemData(.NewIndex) = 1141
        .AddItem "Hebrew": .ItemData(.NewIndex) = 1037
        .AddItem "HID (Human Interface Device)": .ItemData(.NewIndex) = 1279
        .AddItem "Hindi": .ItemData(.NewIndex) = 1081
        .AddItem "Hungarian": .ItemData(.NewIndex) = 1038
        .AddItem "Ibibio (Nigeria)": .ItemData(.NewIndex) = 1129
        .AddItem "Icelandic": .ItemData(.NewIndex) = 1039
        .AddItem "Igbo (Nigeria)": .ItemData(.NewIndex) = 1136
        .AddItem "Indonesian": .ItemData(.NewIndex) = 1057
        .AddItem "Inuktitut": .ItemData(.NewIndex) = 1117
        .AddItem "Italian (Italy)": .ItemData(.NewIndex) = 1040
        .AddItem "Italian (Switzerland)": .ItemData(.NewIndex) = 2064
        .AddItem "Japanese": .ItemData(.NewIndex) = 1041
        .AddItem "Kannada": .ItemData(.NewIndex) = 1099
        .AddItem "Kanuri (Nigeria)": .ItemData(.NewIndex) = 1137
        .AddItem "Kashmiri": .ItemData(.NewIndex) = 2144
        .AddItem "Kashmiri (Arabic)": .ItemData(.NewIndex) = 1120
        .AddItem "Kazakh": .ItemData(.NewIndex) = 1087
        .AddItem "Khmer": .ItemData(.NewIndex) = 1107
        .AddItem "Konkani": .ItemData(.NewIndex) = 1111
        .AddItem "Korean": .ItemData(.NewIndex) = 1042
        .AddItem "Kyrgyz (Cyrillic)": .ItemData(.NewIndex) = 1088
        .AddItem "Lao": .ItemData(.NewIndex) = 1108
        .AddItem "Latin": .ItemData(.NewIndex) = 1142
        .AddItem "Latvian": .ItemData(.NewIndex) = 1062
        .AddItem "Lithuanian": .ItemData(.NewIndex) = 1063
        .AddItem "Malay (Malaysia)": .ItemData(.NewIndex) = 1086
        .AddItem "Malay (Brunei Darussalam)": .ItemData(.NewIndex) = 2110
        .AddItem "Malayalam": .ItemData(.NewIndex) = 1100
        .AddItem "Maltese": .ItemData(.NewIndex) = 1082
        .AddItem "Manipuri": .ItemData(.NewIndex) = 1112
        .AddItem "Maori (New Zealand)": .ItemData(.NewIndex) = 1153
        .AddItem "Marathi": .ItemData(.NewIndex) = 1102
        .AddItem "Mongolian (Cyrillic)": .ItemData(.NewIndex) = 1104
        .AddItem "Mongolian (Mongolian)": .ItemData(.NewIndex) = 2128
        .AddItem "Nepali": .ItemData(.NewIndex) = 1121
        .AddItem "Nepali (India)": .ItemData(.NewIndex) = 2145
        .AddItem "Norwegian (Bokmål)": .ItemData(.NewIndex) = 1044
        .AddItem "Norwegian (Nynorsk)": .ItemData(.NewIndex) = 2068
        .AddItem "Oriya": .ItemData(.NewIndex) = 1096
        .AddItem "Oromo": .ItemData(.NewIndex) = 1138
        .AddItem "Papiamentu": .ItemData(.NewIndex) = 1145
        .AddItem "Pashto": .ItemData(.NewIndex) = 1123
        .AddItem "Polish": .ItemData(.NewIndex) = 1045
        .AddItem "Portuguese (Brazil)": .ItemData(.NewIndex) = 1046
        .AddItem "Portuguese (Portugal)": .ItemData(.NewIndex) = 2070
        .AddItem "Punjabi": .ItemData(.NewIndex) = 1094
        .AddItem "Punjabi (Pakistan)": .ItemData(.NewIndex) = 2118
        .AddItem "Quecha (Bolivia)": .ItemData(.NewIndex) = 1131
        .AddItem "Quecha (Ecuador)": .ItemData(.NewIndex) = 2155
        .AddItem "Quecha (Peru)": .ItemData(.NewIndex) = 3179
        .AddItem "Rhaeto-Romanic": .ItemData(.NewIndex) = 1047
        .AddItem "Romanian": .ItemData(.NewIndex) = 1048
        .AddItem "Romanian (Moldava)": .ItemData(.NewIndex) = 2072
        .AddItem "Russian": .ItemData(.NewIndex) = 1049
        .AddItem "Russian (Moldava)": .ItemData(.NewIndex) = 2073
        .AddItem "Sami (Lappish)": .ItemData(.NewIndex) = 1083
        .AddItem "Sanskrit": .ItemData(.NewIndex) = 1103
        .AddItem "Sepedi": .ItemData(.NewIndex) = 1132
        .AddItem "Serbian (Cyrillic)": .ItemData(.NewIndex) = 3098
        .AddItem "Serbian (Latin)": .ItemData(.NewIndex) = 2074
        .AddItem "Sindhi (India)": .ItemData(.NewIndex) = 1113
        .AddItem "Sindhi (Pakistan)": .ItemData(.NewIndex) = 2137
        .AddItem "Sinhalese (Sri Lanka)": .ItemData(.NewIndex) = 1115
        .AddItem "Slovak": .ItemData(.NewIndex) = 1051
        .AddItem "Slovenian": .ItemData(.NewIndex) = 1060
        .AddItem "Somali": .ItemData(.NewIndex) = 1143
        .AddItem "Sorbian": .ItemData(.NewIndex) = 1070
        .AddItem "Spanish (Spain (Modern Sort))": .ItemData(.NewIndex) = 3082
        .AddItem "Spanish (Spain (Traditional Sort))": .ItemData(.NewIndex) = 1034
        .AddItem "Spanish (Argentina)": .ItemData(.NewIndex) = 11274
        .AddItem "Spanish (Bolivia)": .ItemData(.NewIndex) = 16394
        .AddItem "Spanish (Chile)": .ItemData(.NewIndex) = 13322
        .AddItem "Spanish (Colombia)": .ItemData(.NewIndex) = 9226
        .AddItem "Spanish (Costa Rica)": .ItemData(.NewIndex) = 5130
        .AddItem "Spanish (Dominican Republic)": .ItemData(.NewIndex) = 7178
        .AddItem "Spanish (Ecuador)": .ItemData(.NewIndex) = 12298
        .AddItem "Spanish (El Salvador)": .ItemData(.NewIndex) = 17418
        .AddItem "Spanish (Guatemala)": .ItemData(.NewIndex) = 4106
        .AddItem "Spanish (Honduras)": .ItemData(.NewIndex) = 18442
        .AddItem "Spanish (Latin America)": .ItemData(.NewIndex) = 58378
        .AddItem "Spanish (Mexico)": .ItemData(.NewIndex) = 2058
        .AddItem "Spanish (Nicaragua)": .ItemData(.NewIndex) = 19466
        .AddItem "Spanish (Panama)": .ItemData(.NewIndex) = 6154
        .AddItem "Spanish (Paraguay)": .ItemData(.NewIndex) = 15370
        .AddItem "Spanish (Peru)": .ItemData(.NewIndex) = 10250
        .AddItem "Spanish (Puerto Rico)": .ItemData(.NewIndex) = 20490
        .AddItem "Spanish (United States)": .ItemData(.NewIndex) = 21514
        .AddItem "Spanish (Uruguay)": .ItemData(.NewIndex) = 14346
        .AddItem "Spanish (Venezuela)": .ItemData(.NewIndex) = 8202
        .AddItem "Sutu": .ItemData(.NewIndex) = 1072
        .AddItem "Swahili": .ItemData(.NewIndex) = 1089
        .AddItem "Swedish": .ItemData(.NewIndex) = 1053
        .AddItem "Swedish (Finland)": .ItemData(.NewIndex) = 2077
        .AddItem "Syriac": .ItemData(.NewIndex) = 1114
        .AddItem "Tajik": .ItemData(.NewIndex) = 1064
        .AddItem "Tamazight (Arabic)": .ItemData(.NewIndex) = 1119
        .AddItem "Tamazight (Latin)": .ItemData(.NewIndex) = 2143
        .AddItem "Tamil": .ItemData(.NewIndex) = 1097
        .AddItem "Tatar": .ItemData(.NewIndex) = 1092
        .AddItem "Telugu": .ItemData(.NewIndex) = 1098
        .AddItem "Thai": .ItemData(.NewIndex) = 1054
        .AddItem "Tibetan (Bhutan)": .ItemData(.NewIndex) = 2129
        .AddItem "Tibetan (People's Republic of China)": .ItemData(.NewIndex) = 1105
        .AddItem "Tigrigna (Eritrea)": .ItemData(.NewIndex) = 2163
        .AddItem "Tigrigna (Ethiopia)": .ItemData(.NewIndex) = 1139
        .AddItem "Tsonga": .ItemData(.NewIndex) = 1073
        .AddItem "Tswana": .ItemData(.NewIndex) = 1074
        .AddItem "Turkish": .ItemData(.NewIndex) = 1055
        .AddItem "Turkmen": .ItemData(.NewIndex) = 1090
        .AddItem "Uighur (China)": .ItemData(.NewIndex) = 1152
        .AddItem "Ukrainian": .ItemData(.NewIndex) = 1058
        .AddItem "Urdu": .ItemData(.NewIndex) = 1056
        .AddItem "Urdu (India)": .ItemData(.NewIndex) = 2080
        .AddItem "Uzbek (Cyrillic)": .ItemData(.NewIndex) = 2115
        .AddItem "Uzbek (Latin)": .ItemData(.NewIndex) = 1091
        .AddItem "Venda": .ItemData(.NewIndex) = 1075
        .AddItem "Vietnamese": .ItemData(.NewIndex) = 1066
        .AddItem "Welsh": .ItemData(.NewIndex) = 1106
        .AddItem "Xhosa": .ItemData(.NewIndex) = 1076
        .AddItem "Yi": .ItemData(.NewIndex) = 1144
        .AddItem "Yiddish": .ItemData(.NewIndex) = 1085
        .AddItem "Yoruba": .ItemData(.NewIndex) = 1130
        .AddItem "Zulu": .ItemData(.NewIndex) = 1077
    End With
End Sub

Private Sub pvSetLanguageID(LCID As Long)
    
    ' routine retrieves user's language ID
    
    Dim Buffer As String, x As Long
    Const LOCALE_USER_DEFAULT = &H400
    Const LOCALE_ILANGUAGE = &H1
    
    Call pvFillLanguageIDs
    If LCID = -1& Or LCID = 0& Then
        Buffer = String$(256, 0)
        LCID = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ILANGUAGE, Buffer, Len(Buffer))
        If LCID > 0& Then LCID = Val("&H" & (Left$(Buffer, LCID - 1)))
        If LCID = 0& Then LCID = 1033 ' default to English
    End If
    For x = 0 To cboLanguage.ListCount - 1
        If cboLanguage.ItemData(x) = LCID Then Exit For
    Next
    If x = cboLanguage.ListCount Then
        cboLanguage.AddItem "Unknown (" & CStr(LCID) & ")"
        x = cboLanguage.NewIndex
        cboLanguage.ItemData(x) = LCID
    End If
    cboLanguage.ListIndex = x

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_Manifest = Nothing
End Sub
