VERSION 5.00
Begin VB.Form frmMultiAttrValue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi-Value Element/Attribute"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4500
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
   ScaleHeight     =   3705
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNew 
      Caption         =   "Custom List &Value"
      Height          =   450
      Left            =   135
      TabIndex        =   3
      Top             =   3225
      Width           =   1950
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Cancel"
      Height          =   450
      Index           =   1
      Left            =   2070
      TabIndex        =   2
      Top             =   3225
      Width           =   1125
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Ok"
      Height          =   450
      Index           =   0
      Left            =   3180
      TabIndex        =   1
      Top             =   3225
      Width           =   1125
   End
   Begin VB.ListBox lstValues 
      Height          =   2595
      Left            =   150
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   615
      Width           =   4155
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Label1"
      Height          =   480
      Left            =   210
      TabIndex        =   4
      Top             =   90
      Width           =   4065
   End
End
Attribute VB_Name = "frmMultiAttrValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' comma-delimited list selection options

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const LB_FINDSTRING As Long = &H18F

Private mElement As cManifestEntryEx
Private mAttrIindex As Variant
Private mResult As String

Friend Property Get Value() As String
    Value = mResult
End Property

Friend Sub SetValueList(Value As cManifestEntryEx, AttrIndex As Variant)
    Set mElement = Value
    mAttrIindex = AttrIndex
    
    Dim n As Long, lIndex As String, v() As String
    Dim colValues As Collection, sValue As String
    
    ' setup the listbox values
    lstValues.Clear
    Set colValues = New Collection
    For n = 0& To mElement.GetValueListCount(mAttrIindex) - 1&
        lstValues.AddItem mElement.GetValueListItem(n, mAttrIindex, False)
        ' cross-reference from display value to actual value
        colValues.Add n, mElement.GetValueListItem(n, mAttrIindex)
    Next
    
    ' set checkbox for currently selected item(s)
    v() = Split(mElement.GetValue(mAttrIindex), ",")
    On Error Resume Next
    For n = 0& To UBound(v)
        sValue = Trim$(v(n))
        lIndex = colValues.Item(sValue)
        If Err Then     ' custom value not in our list, add it
            Err.Clear
            lIndex = lstValues.ListCount
            lstValues.AddItem sValue
            lstValues.ItemData(lIndex) = 1      ' flag indicating custom value
        End If
        lstValues.Selected(lIndex) = True
    Next
    On Error GoTo 0
    Set colValues = Nothing
    lblPrompt.AutoSize = True
    
End Sub

Private Sub cmdGo_Click(Index As Integer)
    If Index = 1 Then       ' cancel
        mResult = vbNullString
    Else                    ' ok
        mResult = ""        ' default value
        If lstValues.SelCount > 0 Then
            Dim n As Long
            ' build comma-delimited list
            For n = 0& To lstValues.ListCount - 1&
                If lstValues.Selected(n) = True Then
                    If lstValues.ItemData(n) = 1 Then
                        mResult = mResult & "," & lstValues.List(n)
                    Else
                        mResult = mResult & "," & mElement.GetValueListItem(n, mAttrIindex, True)
                    End If
                End If
            Next
            mResult = Mid$(mResult, 2)
        End If
    End If
    Me.Hide
End Sub

Private Sub cmdNew_Click()
    Dim sValue As String, lIndex As Long
    
    ' add new custom value to the listbox
    sValue = InputBox("Enter new manifest-valid list value below. Case-sensitivity should be enforced", "New List Value")
    If StrPtr(sValue) Then
        sValue = Trim$(sValue)
        If Not sValue = vbNullString Then
            lIndex = SendMessage(lstValues.hWnd, LB_FINDSTRING, -1&, sValue)
            If lIndex = -1 Then
                ' prevent adding actual value that is already referenced by a user-friendly display value
                For lIndex = mElement.GetValueListCount(mAttrIindex) - 1& To 0& Step -1&
                    If StrComp(sValue, mElement.GetValueListItem(lIndex, mAttrIindex), vbTextCompare) = 0 Then
                        lstValues.Selected(lIndex) = True
                        MsgBox "That item is already in the list and is now selected", vbInformation + vbOKOnly, "No Action Taken"
                        Exit For
                    End If
                Next
                If lIndex = -1& Then        ' add new custom value
                    lIndex = lstValues.ListCount
                    lstValues.AddItem sValue
                    lstValues.ItemData(lIndex) = 1
                    lstValues.Selected(lIndex) = True
                End If
            ElseIf lstValues.ItemData(lIndex) = 1 Then
                ' assume changing case-sensitivity
                lstValues.List(lIndex) = sValue
                lstValues.Selected(lIndex) = True
            Else
                MsgBox "That item is already in the list", vbInformation + vbOKOnly, "No Action Taken"
            End If
        End If
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = vbFormCode Then
        Cancel = True: Me.Hide
        mResult = vbNullString
    Else
        Set mElement = Nothing
    End If
End Sub
