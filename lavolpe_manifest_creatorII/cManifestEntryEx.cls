VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cManifestEntryEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ///// PURPOSE \\\\
' the class is simpy used to attach properties and methods to an xml element.
' the class is customized for this project only; it is not plug & play for other projects
' this class should not be created outside the cManifestEx class
' the class if referenced from the cManifestEx class' Item() property

Public Enum AttrPropEnum
    attr_Required = 1       ' attribute is required; cannot be blank (else optional)
    attr_Fixed = 2          ' attribute is set, cannot be changed
    attr_ManualEntry = 4    ' attribute value can be manually entered
    attr_HasValueList = 8   ' suggested values are provided. Manual_Entry handled by calling routine
    attr_CanBeBlank = 16    ' if attr_HasValueList is included, among the values, calling routine allows blank entry
End Enum
Public Enum ElementPropEnum
    elem_Required = 1       ' element is required, must be exported
    elem_Fixed = 2          ' element value is fixed; cannot be changed
    elem_ManualEntry = 4    ' element value can be manually entered (element text value)
    elem_HasValueList = 8   ' suggested values are provided. Manual_Entry handled by calling routine
    elem_CanBeBlank = 16    ' if elem_HasValueList is included, among the values, calling routine allows blank entry
    elem_HasTextElement = 32 ' suggested not to allow user to edit attributes
End Enum

Private Type ValueListStruct
    Values() As String                  ' pairs: actual value then display-friendly value
    Count As Long                       ' size of Values() array: max is 32767 pairs
    Flags As Long                       ' coder-provided values
End Type

'/// This value will force the newly loaded manifest to case-sensitive
'    values for known element/attribute values, i.e., True,TRUE becomes true
' If this is not desired, set the constant to False
' The enforcement only occurs in AddAttributeValueList & AddElementValueList, not in SetValue
Private Const INIT_CASESENSITIVE_ENFORCED As Boolean = True

Private m_Active As Long                ' 1 = not an original element in the source manifest
Private m_Element As IXMLDOMElement
Private m_Options As ValueListStruct    ' .Count: 0x7FFF is size of Values(), 0x8000 = m_Attrs() sized t/f
Private m_Attrs() As ValueListStruct    ' one entry for each attribute
Private m_Parent As String

Public Property Get ManifestElement() As IXMLDOMElement
    Set ManifestElement = m_Element
End Property

Public Sub Attach(Element As IXMLDOMElement, isActive As Long)

    ' Attach this class to a specific element, must be called before any other class methods

    If Element Is Nothing Then
        Err.Raise 380, , "NULL Element Not Valid"
    ElseIf Element.Attributes.length > &H7FFE& Then
        Err.Raise 9, , "Element has too many Attributes"
    Else
        
        Set m_Element = Nothing
        Erase m_Options.Values()
        Erase m_Attrs()
        m_Options.Flags = 0&
        m_Options.Count = 0&
        
        Set m_Element = Element
        m_Active = isActive
    End If
End Sub

Public Function AddElementValueList(ElemProps As ElementPropEnum, ParamArray Values()) As Boolean
    
    ' ElemProps: coder-defined values
    ' Values() ParamArray: pairs of entries, max of 32767 pairs
    '   1st entry is manifest-valid element value, i.e., asInvoker
    '   2nd entry is user-friendly element value, i.e., Standard User Rights
    '   if 2nd entry is null, then 1st entry is used when querying for display value
    
    If m_Element Is Nothing Then Exit Function
    Dim p As Long, sValue As String
    
    With m_Options
        p = UBound(Values)
        If p < 0& Then
            .Count = .Count And &H8000&
            Erase .Values()
        Else
            If (p And 1&) = 0 Then p = p + 1&   ' ensure even number of values
            If p > &H7FFE& Then Exit Function   ' max pairs is 32767
            ReDim .Values(0& To p)
            For p = 0& To UBound(Values)
                .Values(p) = Values(p)
            Next
            .Count = (p \ 2&) Or (.Count And &H8000&)
        
            If INIT_CASESENSITIVE_ENFORCED = True Then   ' change value case as needed
                sValue = m_Element.Text
                If Not sValue = vbNullString Then
                    For p = 0& To p \ 2& - 1& Step 2&
                        If StrComp(.Values(p), sValue, vbTextCompare) = 0 Then
                            If Not .Values(p) = sValue Then m_Element.Text = .Values(p)
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        .Flags = ElemProps
    End With
    AddElementValueList = True

End Function

Public Function AddAttributeValueList(Index As Variant, AttrProps As AttrPropEnum, ParamArray Values()) As Boolean
    
    ' Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
    ' DisplayName: user-friendly name, if null then actual element name is used
    ' AttrProps: coder-defined values
    ' Values() ParamArray: pairs of entries, max of 32767 pairs
    '   1st entry is manifest-valid element value, i.e., asInvoker
    '   2nd entry is user-friendly element value, i.e., Standard User Rights
    '   if 2nd entry is null, then 1st entry is used when querying for display value

    If m_Element Is Nothing Then Exit Function
    Dim p As Long, c As Long, sValue As String
    
    ' locate the correct attribute.
    c = pvIndexFromKey(Index)
    
    ' resize the attribute array as needed
    If (m_Options.Count And &H8000&) = 0& Then
        ReDim m_Attrs(0 To m_Element.Attributes.length - 1&)
        For p = 0& To UBound(m_Attrs): m_Attrs(p).Flags = attr_ManualEntry Or attr_CanBeBlank: Next
        m_Options.Count = m_Options.Count Or &H8000&
    ElseIf c > UBound(m_Attrs) Then ' this should only be triggered if an attribute is added
        p = UBound(m_Attrs) + 1&    ' to the element after we've imported the manifest
        ReDim Preserve m_Attrs(0 To c)
        For p = p To c - 1&: m_Attrs(p).Flags = attr_ManualEntry Or attr_CanBeBlank: Next
    End If
    
    With m_Attrs(c)
        p = UBound(Values)
        If p < 0& Then
            .Count = 0&
            Erase .Values()
        Else
            If (p And 1&) = 0 Then p = p + 1&   ' ensure even number of values
            If p > &H7FFE& Then Exit Function   ' max pairs is 32767
            ReDim .Values(0& To p)
            For p = 0& To UBound(Values)
                .Values(p) = Values(p)
            Next
            .Count = p \ 2&
            
            If INIT_CASESENSITIVE_ENFORCED = True Then   ' change value case as needed
                sValue = m_Element.Attributes(c).nodeValue
                If Not sValue = vbNullString Then
                    For p = 0& To .Count - 1& Step 2&
                        If StrComp(.Values(p), sValue, vbTextCompare) = 0 Then
                            If Not .Values(p) = sValue Then m_Element.Attributes(c).nodeValue = .Values(p)
                            Exit For
                        End If
                    Next
                End If
            End If
            
        End If
        .Flags = AttrProps
    End With

End Function

Public Function GetName(Optional AttributeIndex As Variant, _
                        Optional IncludePrefix As Boolean = False) As String

    ' when calling this for the element, do not pass the AttributeIndex or pass it as Empty
    ' otherwise, Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
    ' if IncludePrefix is true, the namespace prefix is included in the return, i.e., asmv2:trustInfo
    
    If m_Element Is Nothing Then Exit Function
    Dim sValue As String, p As Long
    
    If IsMissing(AttributeIndex) Or IsEmpty(AttributeIndex) Then
        If IncludePrefix Then sValue = m_Element.nodeName Else sValue = m_Element.baseName
    Else
        p = pvIndexFromKey(AttributeIndex)
        If IncludePrefix Then
            sValue = m_Element.Attributes(p).nodeName
        Else
            sValue = m_Element.Attributes(p).baseName
        End If
    End If
    GetName = sValue

End Function

Public Function GetValue(Optional AttributeIndex As Variant, Optional ActualValue As Boolean = True) As String

    ' when calling this for the element, do not pass the AttributeIndex or pass it as Empty
    ' otherwise, Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
    ' if ActualValue is false, then the display-friendly value is returned if any
    
    If m_Element Is Nothing Then Exit Function
    Dim sValue As String, p As Long
    
    If IsMissing(AttributeIndex) Or IsEmpty(AttributeIndex) Then
        sValue = m_Element.Text
        If ActualValue = False And (m_Options.Count And &H7FFF&) > 0& Then
            For p = 0& To (m_Options.Count And &H7FFF&) - 1&
                If m_Options.Values(p * 2&) = sValue Then
                    If m_Options.Values(p * 2& + 1&) <> vbNullString Then sValue = m_Options.Values(p * 2& + 1&)
                    Exit For
                End If
            Next
        End If
    Else
        p = pvIndexFromKey(AttributeIndex)
        sValue = m_Element.Attributes(p).nodeValue
        If ActualValue = False Then
            If (m_Options.Count And &H8000&) > 0& Then
                If p <= UBound(m_Attrs) Then
                    With m_Attrs(p)
                        ' find the display value from the actual value
                        For p = 0& To .Count - 1&
                            If .Values(p * 2&) = sValue Then
                                If .Values(p * 2& + 1&) <> vbNullString Then sValue = .Values(p * 2& + 1&)
                                Exit For
                            End If
                        Next
                    End With
                End If
            End If
        End If
    End If
    GetValue = sValue

End Function

Public Function SetValue(Value As String, Optional AttributeIndex As Variant) As Boolean
    ' when calling this for the element, do not pass the AttributeIndex or pass it as Empty
    ' otherwise, Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
    
    If m_Element Is Nothing Then Exit Function
    If IsMissing(AttributeIndex) Or IsEmpty(AttributeIndex) Then
        m_Element.Text = Value
        SetValue = True
    Else
        Dim p As Long
        ' Error will be raised if the new value is not xml-valid
        p = pvIndexFromKey(AttributeIndex)
        m_Element.Attributes(p).nodeValue = Value
        SetValue = True
    End If
End Function

Public Function GetValueListCount(Optional AttributeIndex As Variant)

    ' when calling this for the element, do not pass the AttributeIndex or pass it as Empty
    ' otherwise, Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
    
    If m_Element Is Nothing Then Exit Function
    If IsMissing(AttributeIndex) Or IsEmpty(AttributeIndex) Then
        GetValueListCount = m_Options.Count And &H7FFF&
    Else
        Dim p As Long
        p = pvIndexFromKey(AttributeIndex)
        If (m_Options.Count And &H8000&) Then
            If p <= UBound(m_Attrs) Then GetValueListCount = m_Attrs(p).Count
        End If
    End If
    
End Function

Public Function GetValueListItem(Index As Long, Optional AttributeIndex As Variant, _
                                Optional ActualValue As Boolean = True) As String
                            
    ' when calling this for the element, do not pass the AttributeIndex or pass it as Empty
    ' otherwise, Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
                            
    If m_Element Is Nothing Then Exit Function
    Dim sValue As String
    If IsMissing(AttributeIndex) Or IsEmpty(AttributeIndex) Then
        If Index < (m_Options.Count And &H7FFF&) Then
            If ActualValue = False Then sValue = m_Options.Values((Index * 2&) + 1&)
            If sValue = vbNullString Then sValue = m_Options.Values(Index * 2&)
        End If
    Else
        Dim p As Long
        p = pvIndexFromKey(AttributeIndex)
        If (m_Options.Count And &H8000&) > 0 Then
            If p <= UBound(m_Attrs) Then
                If Index < m_Attrs(p).Count Then
                    If ActualValue = False Then sValue = m_Attrs(p).Values((Index * 2&) + 1&)
                    If sValue = vbNullString Then sValue = m_Attrs(p).Values((Index * 2&))
                End If
            End If
        End If
    End If
    GetValueListItem = sValue

End Function

Public Property Get ParentKey() As String
    ParentKey = m_Parent
End Property
Public Property Let ParentKey(Value As String)
    m_Parent = Value
End Property

Public Property Get isActive() As Long
    isActive = m_Active
End Property
Public Property Let isActive(Value As Long)
    m_Active = Value
End Property

Public Property Get NumberAttributes() As Long
    If m_Element Is Nothing Then Exit Property
    NumberAttributes = m_Element.Attributes.length
End Property

Public Property Get isElementTopLevel() As Boolean
    ' returns true if element is direct child of the xml root node
    If m_Element Is Nothing Then Exit Property
    isElementTopLevel = (m_Element.parentNode Is m_Element.ownerDocument.documentElement)
End Property

Public Function isElementDescendantOf(Ancestor As String, Optional IncludePrefix As Boolean = True) As Boolean
    ' returns whether the element is a descendant of the passed Ancestor; case-sensitive
    '   if IncludePrefix is true, then the Ancestor parameter must include a prefix, unless null namespace
    ' If Ancestor is passed as null string, the result is same as calling: isElementTopLevel
    If Not m_Element Is Nothing Then
        If Ancestor = vbNullString Then
            isElementDescendantOf = Me.isElementTopLevel
        ElseIf IncludePrefix = True Then
            isElementDescendantOf = Not (m_Element.selectSingleNode("ancestor::*[name()='" & Ancestor & "']") Is Nothing)
        Else
            isElementDescendantOf = Not (m_Element.selectSingleNode("ancestor::*" & "[local-name()='" & Ancestor & "']") Is Nothing)
        End If
    End If

End Function

Public Property Get NameSpace(Optional AttributeIndex As Variant) As String
    ' when calling this for the element, do not pass the AttributeIndex or pass it as -1&
    ' otherwise, Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
    '   Note: for attributes, this should return a null string because of the above statement
    NameSpace = pvGetNSinfo(AttributeIndex, False)
End Property
Public Property Get NameSpacePrefix(Optional AttributeIndex As Variant) As String
    ' when calling this for the element, do not pass the AttributeIndex or pass it as -1&
    ' otherwise, Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
    '   Note: for attributes, this should return a null string because of the above statement
    NameSpacePrefix = pvGetNSinfo(AttributeIndex, True)
End Property

Public Property Let Flags(Optional AttributeIndex As Variant, Value As Long)
    ' when calling this for the element, do not pass the AttributeIndex or pass it as Empty
    ' otherwise, Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
    If m_Element Is Nothing Then Exit Property
    If IsMissing(AttributeIndex) Or IsEmpty(AttributeIndex) Then
        m_Options.Flags = Value
    Else
        Dim p As Long
        p = pvIndexFromKey(AttributeIndex)
        If (m_Options.Count And &H8000&) = 0& Then
            Me.AddAttributeValueList AttributeIndex, Value
        ElseIf p > UBound(m_Attrs) Then
            Me.AddAttributeValueList AttributeIndex, Value
        Else
            m_Attrs(p).Flags = Value
        End If
    End If
End Property
Public Property Get Flags(Optional AttributeIndex As Variant) As Long
    ' when calling this for the element, do not pass the AttributeIndex or pass it as Empty
    ' otherwise, Index is either a 0-bound index of the attribute or its case-sensitive name, i.e., processorArchitecture
    '   note: if applicable and Index is string, it must include the namespace prefix, i.e., prefix:attrName
    '   Attributes with defined Namespaces do not exist in manifests at this time
    If m_Element Is Nothing Then Exit Property
    If IsMissing(AttributeIndex) Or IsEmpty(AttributeIndex) Then
        Flags = m_Options.Flags
    Else
        Dim p As Long
        p = pvIndexFromKey(AttributeIndex)
        If (m_Options.Count And &H8000&) = 0& Then
            Flags = elem_CanBeBlank Or elem_ManualEntry
        ElseIf p > UBound(m_Attrs) Then
            Flags = elem_CanBeBlank Or elem_ManualEntry
        Else
            Flags = m_Attrs(p).Flags
        End If
    End If
End Property

Public Function GetAttributeIndex(Name As String) As Long
    GetAttributeIndex = pvIndexFromKey(Name)
End Function

Public Property Get Key() As String
    Key = "c" & ObjPtr(Me)
End Property

Private Function pvGetNSinfo(AttributeIndex As Variant, AsPrefix As Boolean) As String
    
    If m_Element Is Nothing Then Exit Function
    Dim p As Long, sValue As String
    
    If IsMissing(AttributeIndex) Or IsEmpty(AttributeIndex) Then
        If AsPrefix = True Then sValue = m_Element.prefix Else sValue = m_Element.namespaceURI
    Else
        p = pvIndexFromKey(AttributeIndex)
        If AsPrefix Then
            sValue = m_Element.Attributes(p).prefix
        Else
            sValue = m_Element.Attributes(p).namespaceURI
        End If
    End If
    pvGetNSinfo = sValue

End Function

Private Function pvIndexFromKey(Key As Variant) As Long

    Dim c As Long
    On Error Resume Next
    If VarType(Key) = vbString Then
        For c = m_Element.Attributes.length - 1& To 0& Step -1&
            If m_Element.Attributes(c).nodeName = Key Then Exit For
        Next
    Else
        c = CLng(Key)
        If Err Then
            c = -1&
        ElseIf c < 0& Or c >= m_Element.Attributes.length Then
            c = -1&
        End If
    End If
    On Error GoTo 0
    pvIndexFromKey = c
    If c = -1& Then Err.Raise 380, , "Invalid Key or Index"

End Function

Private Sub Class_Terminate()
    Set m_Element = Nothing
    Erase m_Options.Values
    Erase m_Attrs()
End Sub
