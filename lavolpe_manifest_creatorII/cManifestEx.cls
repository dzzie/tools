VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cManifestEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' PURPOSE: Load, Modify & Create Application Manifests. Moderate support exists for Assembly Manifests
' Notes: The manifest, internally, includes qualified names for all elements
'   The manifest can be written without qualified names if desired
'   The root element will always contain all declared namespaces unless manifest exported without Qualified Names
' Attributes with declared namespaces are handled, but do not exist in manifests at this time

' Note to self for future updates
'   Update the pvCreateLUT routine for any new namespaces added to manifest schema
'   Update following routines if adding new elements/attributes to the base manifest, as needed
'       pvAppendElements, pvMergeElement
'       pvValidateAttributes, pvValidateElement
'       Replace the base manifest; see notes in pvCreateDefaultManifest
' Tip: Check back with this URL every now & again to look for additions/changes
' New software versions and/or revisions may trigger changes to the Manifest schema
'   https://msdn.microsoft.com/en-us/library/windows/desktop/aa374191(v=vs.85).aspx
'   For example, within a week of rewriting this project, msdn added two new items to that page:
'       1) 2017/WindowsSettings: gdiScaling
'       2) an extra attribute value for dpiAwareness: permonitorv2

Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByVal paTypes As Long, ByVal paValues As Long, ByRef retVAR As Variant) As Long

Public Event Begin(lParam As Long)
    ' sent to caller to initialize anything
Public Event ElementAdded(Element As cManifestEntryEx, lParam As Long)
    ' sent for each element that is deemed editable
Public Event Finish(lParam As Long)
    ' sent to caller to allow any finalizing/clean-up

Private m_XML As DOMDocument60
Private m_Element As cManifestEntryEx
Private m_ElementProps As Collection
Private m_nsLUT As Collection
Private m_LCID As Long

Public Property Get xml() As DOMDocument60
    Set xml = m_XML
End Property

Public Property Get LCID() As Long
    LCID = m_LCID
End Property

Public Function CreateManifest(Source As Variant, Optional UserParam As Long, Optional Appended As Boolean = False) As Boolean
    ' reads manifest xml from file or VB resource file
    ' valid Source parameters
    ' String FileName
    ' Nothing for base/default manifest
    ' Clipbaord to paste manifest from clipboard
    
    Dim bOk As Boolean, bImport As Boolean, n As Long
    
    bImport = True
    If IsObject(Source) Then
        If Source Is Nothing Then
            bOk = pvCreateDefaultManifest()
            bImport = False
        Else
            bOk = pvImportManifest(Source, Appended)
        End If
        If bOk Then m_LCID = -1&
        
    ElseIf VarType(Source) = vbString Then
        Dim b() As Byte, lID As Long
        If IsResourceFile(CStr(Source)) = True Then
            b() = StrConv(vbNullString, vbUnicode)
            Call InsertManifestToResource(CStr(Source), b(), lID, True)
            If UBound(b) > 31& Then
                bOk = pvImportManifest(b(), False)
                If bOk Then m_LCID = lID
            Else
                MsgBox "No manifest could be extracted from the resource file", vbInformation + vbOKOnly, "No Action Taken"
            End If
            Erase b()
        Else
            bOk = pvImportManifest(CStr(Source), Appended)
            If bOk Then m_LCID = -1&
        End If
    End If
    If bOk Then
        RaiseEvent Begin(UserParam)
        For n = 1& To m_ElementProps.Count
            Set m_Element = m_ElementProps.Item(n)
            pvValidateElement bImport
            RaiseEvent ElementAdded(m_Element, UserParam)
        Next
        RaiseEvent Finish(UserParam)
        CreateManifest = True
    End If
    Set m_Element = Nothing
    
End Function

Public Function ExportXML(isActiveState As Long, Indentation As Boolean, _
                            NoPrefixes As Boolean, NoEmptyAttrs As Boolean, _
                            Optional IncludeDeclaration As Boolean = True) As DOMDocument60
                            
    ' export the manifest on demand
    ' Parameters:
    '   isActiveState - a bit-wise value contained in the cManifestEntryEx class .isActive property
    '       that value determines whether the element is exported or not
    '   Indentation - true or false (false should be used when targeting res files)
    '   NoPrefixes - true to prevent usage of Qualified Names, i.e., <asmv1:assembly/>
    '   NoEmptyAttrs - true to strip any attributes that are empty/blank
    '   IncludeDeclaration - false results in declarations not included in export
    
    Dim domExport As DOMDocument60, xslt As DOMDocument60
    Dim xmlElement As IXMLDOMElement
    Dim sPrefix As String, sNS As String, n As Long, a As Long
    Dim b() As Byte, xmlAttrs As IXMLDOMSelection
    
    ' flag each main element that will be written to the manifest
    For n = 1& To m_ElementProps.Count
        Set m_Element = m_ElementProps.Item(n)
        If (m_Element.isActive And isActiveState) Then
            Set xmlElement = m_Element.ManifestElement
            If NoEmptyAttrs Then
                For a = 0& To xmlElement.Attributes.length - 1&
                    If (m_Element.Flags(a) And attr_Required) Then
                        If xmlElement.Attributes(a).nodeValue = vbNullString Then
                            ' ensure required attributes are exported, even if empty/blank
                            xmlElement.Attributes(a).nodeValue = "dummy" & ObjPtr(Me)
                        End If
                    End If
                Next
            End If
            xmlElement.setAttribute "mcIIid", "ok"  ' flag the element
        End If
    Next
    Set m_Element = Nothing
    
    ' load a XSLT tranformation to extract the flagged elements and their ancestors
    ' fyi: the xslt can be extracted manually from the res file by opening that file in NotePad
    Set xslt = New DOMDocument60
    b() = LoadResData("XSLT-SUBLST", "CUSTOM")
    xslt.loadXML StrConv(b(), vbUnicode)
    Erase b()
    ' set the XLST parameter to match the flag we set above
    Set xmlElement = xslt.documentElement.selectSingleNode("*[local-name(.)='param']")
        xmlElement.setAttribute "select", "//*[@mcIIid]"
    Set xmlElement = Nothing
    
    ' create the extracted xml. note: encoding will be UTF-16 & we'll clean that up in pvPrettyPrintXML
    Set domExport = New DOMDocument60
    m_XML.transformNodeToObject xslt, domExport
    Set xslt = Nothing
    
    ' remove the temp attrs set and any blank attrs
    If NoEmptyAttrs Then
        Set xmlAttrs = domExport.documentElement.selectNodes("//attribute::*[(.='')]")
        If xmlAttrs.length Then xmlAttrs.removeAll
        Set xmlAttrs = domExport.documentElement.selectNodes("//attribute::*[.='dummy" & ObjPtr(Me) & "']")
        For n = xmlAttrs.length - 1& To 0& Step -1&
            xmlAttrs.Item(n).nodeValue = vbNullString
        Next
    End If
    Set xmlAttrs = domExport.documentElement.selectNodes("//attribute::*[name(.)='mcIIid']")
    If xmlAttrs.length Then xmlAttrs.removeAll
    ' finally, remove the flag from the base manifest
    Set xmlAttrs = m_XML.documentElement.selectNodes("//attribute::*[name(.)='mcIIid']")
    If xmlAttrs.length Then xmlAttrs.removeAll
    Set xmlAttrs = Nothing
    
    ' clear unused namespaces
    With domExport
        For n = .namespaces.length - 1& To 0& Step -1&
            sNS = .namespaces(n)
            sPrefix = Me.GetNameSpacePrefix(.namespaces(n))
            If Right$(sPrefix, 1) = ":" Then
                If .documentElement.selectSingleNode("//*[namespace-uri()='" & sNS & "']") Is Nothing Then
                    .documentElement.removeAttribute "xmlns:" & Left$(sPrefix, Len(sPrefix) - 1&)
                End If
            End If
        Next
    End With
    
    ' convert to prefix-less format if requested. note: encoding will be UTF-16 & we'll clean that up in pvPrettyPrintXML
    If NoPrefixes Then Call pvRemovePrefixes(domExport)
    
    ' finalize the export
    Call pvPrettyPrintXML(domExport, Indentation, IncludeDeclaration)
    
    Set ExportXML = domExport

End Function

Public Property Get Item(Key As Variant) As cManifestEntryEx
    ' allow error to be raised
    Set Item = m_ElementProps.Item(Key) ' key is a property of the cManifestEntryEx class or ordinal 1-based index
End Property

Public Property Get ItemCount() As Long
    If Not m_ElementProps Is Nothing Then ItemCount = m_ElementProps.Count
End Property

Public Function AddItem(ByVal Element As IXMLDOMElement, Optional UserParam As Long) As String

    ' return value is the Key of the added element, if successful

    ' We will be calling the pvImportManifest method which expects a complete XML,
    '   including a root node. We will create a dummy root node for the passed element.
    ' To prevent potential of changing passed element's namespace, use its namespace.

    Dim tElement As IXMLDOMElement, n As Long
    Set tElement = Element.selectSingleNode("ancestor::*")  ' get passed node's top ancestor (if any)
    If tElement Is Nothing Then Set tElement = Element      ' if it is top already, use it instead
    ' create the dummy root and append passed node as its child
    Set tElement = m_XML.createNode(NODE_ELEMENT, "assembly", Element.namespaceURI).appendChild(tElement)
    
    n = m_ElementProps.Count
    If pvImportManifest(tElement.parentNode, False) = False Then Exit Function
    pvValidateElement True
    RaiseEvent Begin(UserParam)
    For n = n + 1& To m_ElementProps.Count
        Set m_Element = m_ElementProps.Item(n)
        RaiseEvent ElementAdded(m_Element, UserParam)
    Next
    RaiseEvent Finish(UserParam)
    AddItem = m_Element.Key
    Set m_Element = Nothing
    
End Function

Public Function AddSubItem(ByVal Element As IXMLDOMElement, Optional parentKey As String, Optional UserParam As Long) As String

    ' add child element to existing manifest. Child element can be displayed/edited
    ' parentKey must be provided if the new element's parent is displayed on the treeview
    
    Dim n As Long
    
    n = m_ElementProps.Count
    Set m_Element = New cManifestEntryEx
    m_Element.Attach Element, 1&
    m_Element.parentKey = parentKey
    pvValidateElement True
    m_ElementProps.Add m_Element, m_Element.Key
    RaiseEvent Begin(UserParam)
    For n = n + 1& To m_ElementProps.Count
        Set m_Element = m_ElementProps.Item(n)
        RaiseEvent ElementAdded(m_Element, UserParam)
    Next
    RaiseEvent Finish(UserParam)
    AddSubItem = m_Element.Key
    Set m_Element = Nothing

End Function

Public Sub RemoveItem(Key As Variant)
    
    ' allow error to be raised; key is a property of the cManifestEntryEx class or ordinal 1-based index
    Dim cNode As IXMLDOMElement
    Set cNode = m_ElementProps.Item(Key).ManifestElement
    cNode.parentNode.removeChild cNode
    Set cNode = Nothing
    m_ElementProps.Remove Key
    
End Sub

Private Function pvCreateDefaultManifest() As Boolean

    ' the base/default manifest is stored in the resource file: CUSTOM, BASE
    ' if for any reason the base needs to be updated, it must be rewritten to file
    '   and re-inserted into the resource file. Reasons for updating could be:
    '   - a namespace is deemed incorrect
    '   - a new manifest item is to be added
    '   - other corrections/additions that cannot be resolved via pvValidateAttributes or pvValidateElement
    ' note: no need to change the base manifest in these scenarios:
    '   - new attribute: ensure attribute is added via: pvValidateAttributes
    
    ' To create the base, follow these steps
    '   1. Run the project and create a new default manifest
    '   2. Select every one of the manifest elements available
    '   3. Use the menu to write the XML to a new file & choose the file name.
    '       - INCLUDE empty attributes
    '       - This is what will be placed in res file
    '   4. Open the file in a text editor and make whatever changes are needed
    '       - add new namespaces to the root element's attribute list as needed
    '       - ensure all known optional attributes are included with elements, if applicable
    '   5. Close this project and open the res file in the resource editor
    '   6. Remove the existing BASE entry, and insert the new file
    '   7. Change the resource ID from 101 (or whatever it was given) to BASE
    '       - double click the resource to rename it
    '   8. Make any code changes in this project if needed
    '       see top of class for common routines to modify. see top of frmMain for changes there
    '   9. Run project and verify your changes

    Dim xmlElem As IXMLDOMElement, b() As Byte
    Dim ns As String, n As Long
    
    pvCreateLUT ' initialize our "namespace manager"
    Set m_ElementProps = New Collection
    
    ' load the base manifest. This base includes one-each of every known possible manifest entry
    ' fyi: the base can be extracted manually from the res file by opening that file in NotePad
    b() = LoadResData("BASE", "CUSTOM")
    Set m_XML = New DOMDocument60
    m_XML.loadXML StrConv(b(), vbUnicode)
    Erase b()
    
    ' create collection of elements that could be modified (about 2 dozen)
    With m_XML.documentElement.selectNodes("//*[not(*)]")
        For n = 0& To .length - 1&
            Set m_Element = New cManifestEntryEx
            m_Element.Attach .Item(n), 0&
            m_ElementProps.Add m_Element, m_Element.Key
        Next
    End With
    
    ' build SelectionNamespace property
    '   Every element in the base manifest has qualified names
    '   Any namespace referenced in a qualified name, exists in the root elements attribute list
    With m_XML.documentElement.Attributes
        For n = 0& To .length - 1&
            If .Item(n).prefix = "xmlns" Then
                ns = ns & " xmlns:" & .Item(n).baseName & "='" & .Item(n).nodeValue & "'"
            End If
        Next
    End With
    m_XML.setProperty "SelectionNamespaces", ns: ns = vbNullString
    
    pvCreateDefaultManifest = True
    
ErrHandler:
    If Err Then
        Err.Clear
    Else
        pvCreateDefaultManifest = True
    End If
End Function

Public Function GetNameSpace(nsPrefix As String) As String

    ' returns Namespace referenced by passed nsPrefix

    Dim ns As String, bAppend As Boolean
    On Error Resume Next
    If Trim$(nsPrefix) = vbNullString Then
        nsPrefix = vbNullString
    Else
        ns = m_nsLUT.Item(nsPrefix)
        If Err Then     ' never called based on an prefix from an external source; only from internal manifest
            Stop        ' should not get here, code calling this must ensure NS exists beforehand
        Else
            nsPrefix = m_nsLUT.Item(ns)
            bAppend = IsNull(m_XML.documentElement.getAttribute("xmlns:" & nsPrefix))
        End If
    End If
    On Error GoTo 0
    If bAppend Then
        m_XML.documentElement.setAttribute "xmlns:" & nsPrefix, ns
        m_XML.setProperty "SelectionNamespaces", _
            m_XML.getProperty("SelectionNamespaces") & " xmlns:" & nsPrefix & "='" & ns & "'"
    End If
    GetNameSpace = ns

End Function

Public Function GetNameSpacePrefix(NameSpace As String, Optional PreferredPrefix As String) As String
    
    ' returns the case-sensitive namespace and prefix
    ' optionally appends new namespaces to our collection, using PreferredPrefix or generated prefix
    
    Dim sPrefix As String, lCount As Long, bAppend As Boolean
    On Error Resume Next
    If Trim$(NameSpace) = vbNullString Then
        NameSpace = sPrefix
        Exit Function               ' don't allow the colon suffix
    Else
        sPrefix = m_nsLUT.Item(NameSpace)
        If Err Then                 ' add it to our collection
            Err.Clear
            If PreferredPrefix = vbNullString Then sPrefix = "nsGen" Else sPrefix = PreferredPrefix
            Do
                sPrefix = m_nsLUT.Item(sPrefix)
                If Err Then
                    Err.Clear: Exit Do
                End If
                lCount = lCount + 1
                If PreferredPrefix = vbNullString Then sPrefix = "nsGen" & lCount Else sPrefix = PreferredPrefix & lCount
            Loop
            On Error GoTo 0
            m_nsLUT.Add sPrefix, NameSpace
            m_nsLUT.Add NameSpace, sPrefix
            bAppend = True
        Else
            On Error GoTo 0
            bAppend = IsNull(m_XML.documentElement.getAttribute("xmlns:" & sPrefix))
        End If
    End If
    If bAppend Then
        m_XML.documentElement.setAttribute "xmlns:" & sPrefix, NameSpace
        m_XML.setProperty "SelectionNamespaces", _
            m_XML.getProperty("SelectionNamespaces") & " xmlns:" & sPrefix & "='" & NameSpace & "'"
    End If
    GetNameSpacePrefix = sPrefix & ":"
    
End Function

Private Sub pvCreateLUT()

    ' mini-namespace manager. Reverse look-up style:
    '   namespace returns prefix & prefix returns namespace
    ' collection appended, as needed, via GetNameSpacePrefix method
    
    ' Goals
    '   1. ensure no prefix (case-insensitive) is used for more than 1 namespace
    '   2. ensure any namespace is not duplicated due to only upper/lower case differences
    '   3. prefer namespace case-sensitivity from list below; assumed known-good

    Set m_nsLUT = New Collection
    m_nsLUT.Add "urn:schemas-microsoft-com:asm.v1", "asmv1"
        m_nsLUT.Add "asmv1", m_nsLUT.Item(m_nsLUT.Count)
    m_nsLUT.Add "urn:schemas-microsoft-com:asm.v2", "asmv2"
        m_nsLUT.Add "asmv2", m_nsLUT.Item(m_nsLUT.Count)
    m_nsLUT.Add "urn:schemas-microsoft-com:asm.v3", "asmv3"
        m_nsLUT.Add "asmv3", m_nsLUT.Item(m_nsLUT.Count)
    m_nsLUT.Add "urn:schemas-microsoft-com:compatibility.v1", "comp"
        m_nsLUT.Add "comp", m_nsLUT.Item(m_nsLUT.Count)
    m_nsLUT.Add "http://schemas.microsoft.com/SMI/2005/WindowsSettings", "ws2005"
        m_nsLUT.Add "ws2005", m_nsLUT.Item(m_nsLUT.Count)
    m_nsLUT.Add "http://schemas.microsoft.com/SMI/2011/WindowsSettings", "ws2011"
        m_nsLUT.Add "ws2011", m_nsLUT.Item(m_nsLUT.Count)
    m_nsLUT.Add "http://schemas.microsoft.com/SMI/2013/WindowsSettings", "ws2013"
        m_nsLUT.Add "ws2013", m_nsLUT.Item(m_nsLUT.Count)
    m_nsLUT.Add "http://schemas.microsoft.com/SMI/2014/WindowsSettings", "ws2014"
        m_nsLUT.Add "ws2014", m_nsLUT.Item(m_nsLUT.Count)
    m_nsLUT.Add "http://schemas.microsoft.com/SMI/2016/WindowsSettings", "ws2016"
        m_nsLUT.Add "ws2016", m_nsLUT.Item(m_nsLUT.Count)
    m_nsLUT.Add "http://schemas.microsoft.com/SMI/2017/WindowsSettings", "ws2017"
        m_nsLUT.Add "ws2017", m_nsLUT.Item(m_nsLUT.Count)
    
End Sub

Private Function pvImportManifest(Source As Variant, Appended As Boolean) As Boolean

    ' imports an external manifest and merges it with our internal manifest
    ' namespaces are validated during the process as are required element attributes

    Dim doc As DOMDocument60, docSrc As DOMDocument60
    Dim xmlElement As IXMLDOMElement, xmlParent As IXMLDOMElement
    Dim tElements As New Collection, tLUT As Collection
    Dim ns As String, sPrefix As String, bAppended As Boolean
    Dim n As Long, l As Long, bAbort As Boolean
    Dim xmlList As IXMLDOMNodeList, cNode As IXMLDOMElement
    
    If VarType(Source) = vbString Then              ' must be a file
        Set docSrc = New DOMDocument60
        docSrc.Load CStr(Source)
    ElseIf VarType(Source) = (vbArray Or vbByte) Then ' must be a byte array (retrieved from resource file)
        Set docSrc = New DOMDocument60
        docSrc.loadXML StrConv(Source, vbUnicode)
        Erase Source
    ElseIf IsObject(Source) Then
        If Source Is Nothing Then Exit Function
        If Source Is Clipboard Then
            Set docSrc = New DOMDocument60
            docSrc.loadXML Clipboard.GetText
        Else                                        ' called from AddItem
            On Error Resume Next
            Set cNode = Source
            If Err Then
                MsgBox "The provided source does not appear to be a valid XML manifest", vbExclamation + vbOKOnly, "No Action Taken"
                Err.Clear: Exit Function
            End If
        End If
    End If
    
    If Not docSrc Is Nothing Then                   ' else direct insert; guaranteed to use qualified names
        If Not docSrc.parseError.errorCode = 0& Then
            MsgBox "Failed to load the selected file." & vbCrLf & _
                docSrc.parseError.reason, vbExclamation + vbOKOnly, "No Action Taken"
            Set docSrc = Nothing
            Exit Function
        Else
            With docSrc.documentElement
                If Not .getAttribute("manifestVersion") & vbNullString = "1.0" Then
                    MsgBox "The provided source does not contain the required maninfestVersion attribute and/or value", vbExclamation + vbOKOnly, "No Action Taken"
                    Set docSrc = Nothing
                    Exit Function
                ElseIf Trim$(.namespaceURI) = vbNullString Then
                    MsgBox "The provided source's root is not in the required name space.", vbExclamation + vbOKOnly, "No Action Taken"
                    Set docSrc = Nothing
                    Exit Function
                End If
            End With
        End If
        
        Set doc = m_XML                     ' back these up so we can undo if needed
        Set tElements = m_ElementProps
        Set tLUT = m_nsLUT
        If Appended = False Then
            Set m_ElementProps = New Collection ' create new versions
            Call pvCreateLUT
            Call pvCreateDefaultManifest
        End If
        Set cNode = docSrc.documentElement
    End If
    
    On Error GoTo ErrHandler
    With cNode.selectNodes("//*[not(*)]")   ' focus on lowest level elements
        For l = 0& To .length - 1&
            Set xmlElement = .Item(l)
            If xmlElement.Attributes.getNamedItem("mcIIid") Is Nothing Then   ' else already done
                Set xmlParent = m_XML.documentElement
                With xmlElement.selectNodes("ancestor::*")
                    For n = 1& To .length - 1&
                        sPrefix = Me.GetNameSpacePrefix(.Item(n).namespaceURI, .Item(n).prefix)
                        Set xmlElement = xmlParent.selectSingleNode(sPrefix & .Item(n).baseName)
                        If xmlElement Is Nothing Then
                            Set xmlElement = .Item(n)
                            Exit For  ' insert from this point
                        End If
                        Set xmlParent = xmlElement
                    Next
                    If n = .length Then Set xmlElement = Nothing
                End With
                If xmlElement Is Nothing Then Set xmlElement = .Item(l)
                pvMergeElement xmlElement, xmlParent, (doc Is Nothing)
            End If
        Next
    End With
    On Error GoTo 0
    
ErrHandler:
    If Err Then
        Err.Clear
        MsgBox "Error importing source document." & Err.Description, vbExclamation + vbOKOnly, "No Action Taken"
        bAbort = True
    End If
    pvImportManifest = Not bAbort
    If Not doc Is Nothing Then
        Set docSrc = Nothing
        If bAbort Then
            Set m_XML = doc
            Set m_ElementProps = tElements
            Set m_nsLUT = tLUT
        End If
        Set tElements = Nothing
        Set doc = Nothing
        Set tLUT = Nothing
    End If
    
End Function

Private Sub pvMergeElement(srcElement As IXMLDOMElement, parentElement As IXMLDOMElement, AlwaysAppend As Boolean)

    ' overwrites existing elements if those elements are not yet in-use
    ' some elements will be allowed to be duplicated, most are only allowed to exist once

    Dim dstElement As IXMLDOMElement
    Dim bNotLoaded As Boolean, bMultiInstance As Boolean
    Dim bLoadFromRoot As Boolean
    Dim a As Long, sPrefix As String, ns As String
    
    sPrefix = Me.GetNameSpacePrefix(srcElement.namespaceURI)
    Set dstElement = parentElement.selectSingleNode(sPrefix & srcElement.baseName)
    
    Select Case parentElement.baseName
    Case "assembly"
        Select Case srcElement.baseName
        Case "assemblyIdentity"
        Case "description"
            dstElement.Text = srcElement.Text
        Case "file"
            If AlwaysAppend Then
                bNotLoaded = True
            Else
                bMultiInstance = True
            End If
        Case Else   ' other root children not associated with typical app manifests
            bNotLoaded = True
        End Select
        
    Case "dependentAssembly"
        If AlwaysAppend Then
            bNotLoaded = True: bLoadFromRoot = True
        Else
            On Error Resume Next
            Select Case LCase$(srcElement.getAttribute("name"))
            Case "microsoft.windows.common-controls"
                Set dstElement = m_XML.documentElement.selectSingleNode("//" & sPrefix & srcElement.baseName & "[@name='Microsoft.Windows.Common-Controls']")
            Case "microsoft.windows.gdiplus"
                Set dstElement = m_XML.documentElement.selectSingleNode("//" & sPrefix & srcElement.baseName & "[@name='Microsoft.Windows.GdiPlus']")
            Case Else   ' append from the root
                bNotLoaded = True: bLoadFromRoot = True
            End Select
            On Error GoTo 0
        End If
    Case "application"
        If srcElement.baseName = "supportedOS" Then
            On Error Resume Next
            Set dstElement = parentElement.selectSingleNode(sPrefix & srcElement.baseName & "[@Id='" & LCase$(srcElement.getAttribute("Id") & "']"))
            If Err.Number <> 0 Or dstElement Is Nothing Then
                If IsNull(srcElement.getAttribute("Id")) = True Then Exit Sub
                bNotLoaded = True
            End If
            On Error GoTo 0
        Else
            bNotLoaded = True
        End If
    Case "windowsSettings"
        If dstElement Is Nothing Then
            bNotLoaded = True
        Else
            dstElement.Text = srcElement.Text
        End If
    Case "requestedPrivileges"
    Case Else
        bNotLoaded = True: bLoadFromRoot = True
    End Select
    
    If bNotLoaded = False Then
        For a = 1& To m_ElementProps.Count
            Set m_Element = m_ElementProps.Item(a)
            If m_Element.ManifestElement Is dstElement Then
                If m_Element.isActive = 0& Then
                    m_Element.isActive = 1&
                ElseIf bMultiInstance = False Then
                    Exit Sub
                Else
                    bNotLoaded = True
                End If
                Exit For
            End If
        Next
    End If
    
    If bNotLoaded Then
        If bLoadFromRoot Then
            Set srcElement = srcElement.selectNodes("ancestor::*")(1)
            Set parentElement = m_XML.documentElement
        End If
        pvAppendElements srcElement, parentElement, vbNullString, True
    Else
        With srcElement.Attributes
            For a = 0& To .length - 1&
                If StrComp(.Item(a).baseName, "xmlns", vbTextCompare) Then
                    If .Item(a).prefix = vbNullString Then
                        dstElement.setAttribute .Item(a).nodeName, .Item(a).nodeValue
                    ElseIf StrComp(.Item(a).prefix, "xmlns", vbTextCompare) Then
                        ns = .Item(a).namespaceURI
                        sPrefix = Me.GetNameSpacePrefix(ns)
                        dstElement.Attributes.setNamedItem(m_XML.createNode(NODE_ATTRIBUTE, sPrefix & .Item(a).baseName, ns)).nodeValue = .Item(a).nodeValue
                    End If
                End If
            Next
        End With
        sPrefix = m_Element.Key
        With srcElement.selectNodes("*")
            For a = 0& To .length - 1&
                pvAppendElements .Item(a), dstElement, sPrefix, True
            Next
        End With
    End If

End Sub

Private Sub pvAppendElements(srcElement As IXMLDOMNode, parentElement As IXMLDOMElement, _
                            ByVal pKey As String, ByVal isBaseLevel As Boolean)

    ' recursive routine to append elements to the manifest
    ' items added to the m_ElementProps collection must meet one of two conditions
    '   1) has attributes (user should be able to edit these)
    '   2) an element's lowest child element(s). These generally include attributes or text elements

    Dim a As Long, tElement As IXMLDOMElement, bHasAttrs As Boolean
    Dim ns As String, sPrefix As String, bAdd As Boolean

    Do Until srcElement Is Nothing
        If srcElement.nodeType = NODE_ELEMENT Then
            ns = srcElement.namespaceURI
            sPrefix = Me.GetNameSpacePrefix(ns)
            ' place the inserted node relative to others in same namespace
            Set tElement = parentElement.insertBefore( _
                    m_XML.createNode(NODE_ELEMENT, sPrefix & srcElement.baseName, ns), _
                    pvGetInsertPos(parentElement, sPrefix & srcElement.baseName))
            With srcElement.Attributes              ' append any attrs
                For a = 0& To .length - 1&          ' skipping namespace attrs and/or applying namespace as needed
                    If StrComp(.Item(a).baseName, "xmlns", vbTextCompare) Then
                        If .Item(a).prefix = vbNullString Then
                            tElement.setAttribute .Item(a).nodeName, .Item(a).nodeValue
                            bAdd = True
                        ElseIf StrComp(.Item(a).prefix, "xmlns", vbTextCompare) Then
                            ns = .Item(a).namespaceURI: sPrefix = Me.GetNameSpacePrefix(ns)
                            tElement.Attributes.setNamedItem(m_XML.createNode(NODE_ATTRIBUTE, sPrefix & .Item(a).baseName, ns)).nodeValue = .Item(a).nodeValue
                            bAdd = True
                        End If
                    End If
                Next
            End With
            If bAdd = False Then bAdd = (srcElement.selectSingleNode("*") Is Nothing)
            If bAdd Then
                srcElement.Attributes.setNamedItem(srcElement.ownerDocument.createAttribute("mcIIid")).nodeValue = "done"
                Set m_Element = New cManifestEntryEx
                m_Element.Attach tElement, 1&
                m_Element.parentKey = pKey
                sPrefix = m_Element.Key
                m_ElementProps.Add m_Element, sPrefix
                a = m_ElementProps.Count
            Else
                sPrefix = vbNullString
            End If
            If srcElement.hasChildNodes Then
                pvAppendElements srcElement.firstChild, tElement, sPrefix, False
            End If
            If bAdd Then bAdd = False
        ElseIf srcElement.nodeType = NODE_TEXT Then
            parentElement.Text = srcElement.nodeValue
        End If
        If isBaseLevel Then Exit Do
        Set srcElement = srcElement.nextSibling
    Loop

End Sub

Private Function pvPrettyPrintXML(domSource As DOMDocument60, Indented As Boolean, IncludeDeclaration As Boolean) As String
    
    ' exports the manifest with options for indentation
    
    Dim rdrCompact As MSXML2.SAXXMLReader60
    Dim wrtFormatted As MSXML2.MXXMLWriter60
    Dim doc As DOMDocument60, IStream As IUnknown
    Dim vPtr(0 To 2) As Long, vType(0 To 2) As Integer
    Dim vParams(0 To 3) As Variant
    Dim hGlobal As Long
    
    ' creates a self-releasing IStream on hGlobal.
    hGlobal = GlobalAlloc(&H2&, 0&)
    If Not hGlobal = 0& Then
        Call CreateStreamOnHGlobal(hGlobal, 1&, IStream)
        If IStream Is Nothing Then
            GlobalFree hGlobal ' not good
            Exit Function
        End If
    End If
    
    Set wrtFormatted = New MSXML2.MXXMLWriter60
    Set rdrCompact = New MSXML2.SAXXMLReader60
    With wrtFormatted
        .encoding = "UTF-8"     ' manifest specs do not indicate UTF-16 is valid
        .byteOrderMark = True   ' will likely not be written anyway if xml is pure ansi
        .omitXMLDeclaration = Not IncludeDeclaration
        .StandAlone = True
        .indent = Indented
        .output = IStream
    End With
    With rdrCompact
        Set .contentHandler = wrtFormatted
        Set .dtdHandler = wrtFormatted
        Set .errorHandler = wrtFormatted
        .putProperty "http://xml.org/sax/properties/lexical-handler", wrtFormatted
        .putProperty "http://xml.org/sax/properties/declaration-handler", wrtFormatted
        .parse domSource
    End With
    If Not IStream Is Nothing Then
        Set doc = New DOMDocument60
        ' when reading from a stream, ensure the seek pointer is back to 0 position, else failure
        vParams(0) = 0@: vParams(1) = 0&: vParams(2) = 0&
        vPtr(0) = VarPtr(vParams(0)): vPtr(1) = VarPtr(vParams(1)): vPtr(2) = VarPtr(vParams(2))
        vType(0) = vbCurrency: vType(1) = vbLong: vType(2) = vbLong
        DispCallFunc ObjPtr(IStream), 20&, 4&, vbLong, 3&, VarPtr(vType(0)), VarPtr(vPtr(0)), vParams(3)
        
        doc.Load IStream
        Set domSource = doc
        Set IStream = Nothing
        Set doc = Nothing
    End If
End Function

Public Sub pvRemovePrefixes(domSource As DOMDocument60)

    Dim docXLST As DOMDocument60
    Dim xmlElem As IXMLDOMElement
    Dim sPrefix As String, ns As String
    Dim n As Long
   
    ' create an XLST transformation that removes namespace prefixes
    ' This will still preserve namespaces and does have two requirements:
    '   1. no prefix can be used for more than one NS
    '   2. the relationship between a prefix and its NS is known in advance
    ' Since we control this relationship via our little namespace manager, above two are done
   
    ' Basically, we remove the NS from each node and specifically assign it the NS as an attribute
    ' After the entire document is transformed, standard namespace inheritance applies
   
    ' The primary template in the xslt looks like the following, one template per named namespace
   
    Set docXLST = New DOMDocument60
    docXLST.loadXML "<?xml version='1.0'?><xsl:stylesheet version='1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'><xsl:output omit-xml-declaration='no'/></xsl:stylesheet>"
    With docXLST.documentElement
        Set xmlElem = .appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:output", .namespaceURI))
            xmlElem.setAttribute "method", "xml"
            
        ' create the catch-all templates for null-namespace if it exists in the manifest
        Set xmlElem = .appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:template", .namespaceURI))
            xmlElem.setAttribute "match", "*"
        Set xmlElem = xmlElem.appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:element", .namespaceURI))
            xmlElem.setAttribute "name", "{local-name()}"
        Set xmlElem = xmlElem.appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:apply-templates", .namespaceURI))
            xmlElem.setAttribute "select", "node()|@*"
   
        Set xmlElem = .appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:template", .namespaceURI))
            xmlElem.setAttribute "match", "@*"
        Set xmlElem = xmlElem.appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:copy", .namespaceURI))
        Set xmlElem = xmlElem.appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:apply-templates", .namespaceURI))
   
        ' create filters for specific namespace conversion
        For n = 0& To domSource.namespaces.length - 1&
            ns = domSource.namespaces(n): sPrefix = Me.GetNameSpacePrefix(ns)
            If Right$(sPrefix, 1) = ":" Then
                Set xmlElem = .appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:template", .namespaceURI))
                    xmlElem.setAttribute "match", sPrefix & "*"
                    xmlElem.setAttribute "xmlns:" & Left$(sPrefix, Len(sPrefix) - 1&), ns
                Set xmlElem = xmlElem.appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:element", .namespaceURI))
                    xmlElem.setAttribute "name", "{local-name()}"
                    xmlElem.setAttribute "namespace", ns
                Set xmlElem = xmlElem.appendChild(docXLST.createNode(NODE_ELEMENT, "xsl:apply-templates", .namespaceURI))
                    xmlElem.setAttribute "select", "node()|@*"
            End If
        Next
    End With
    domSource.transformNodeToObject docXLST, domSource
    Set docXLST = Nothing
    
End Sub

Private Sub pvValidateElement(isImportedElement As Boolean)

    ' routine applies properties of the element to the active cManifestEntryEx class (m_Element)
    ' note: elem_Fixed applied below means the user cannot change the element's value
    ' enumeration comes from cManifestEntryEx
    
    Dim xmlElement As IXMLDOMElement, sPrefix As String, sKey As String
    Dim bNotUpdated As Boolean
    
    Set xmlElement = m_Element.ManifestElement
    pvValidateAttributes xmlElement, isImportedElement
    
    With xmlElement
        Select Case .parentNode.baseName
        Case "assembly"
            Select Case .baseName
            Case "assemblyIdentity"
                m_Element.AddElementValueList elem_Fixed Or elem_Required
            Case "file"
                If .hasChildNodes = False Then
                    m_Element.AddElementValueList elem_Fixed
                Else
                    m_Element.AddElementValueList elem_ManualEntry Or elem_HasTextElement
                End If
            Case "description"
                m_Element.AddElementValueList elem_HasTextElement Or elem_ManualEntry Or elem_CanBeBlank
            Case "noInherit", "noInheritable"
                m_Element.AddElementValueList elem_Fixed
            Case "windowClass"
                m_Element.AddElementValueList elem_HasTextElement Or elem_ManualEntry
            Case "comInterfaceExternalProxyStub"
                m_Element.AddElementValueList elem_Fixed
            Case Else
                bNotUpdated = True
            End Select
        Case "dependentAssembly"
            If .baseName = "assemblyIdentity" Then
                m_Element.AddElementValueList elem_Fixed
            Else
                bNotUpdated = True
            End If
        Case "requestedPrivileges"
            m_Element.AddElementValueList elem_Fixed
        Case "file"
            If .baseName = "windowClass" Then
                m_Element.AddElementValueList elem_HasTextElement Or elem_ManualEntry
            Else
                m_Element.AddElementValueList elem_Fixed
            End If
        Case "requestedPrivileges"
            If .baseName = "requestedExecutionLevel" Then m_Element.AddElementValueList elem_Fixed
        Case "windowsSettings"
            Select Case .baseName
            Case "dpiAware" ' offer user-friendly values
                m_Element.AddElementValueList elem_Fixed Or elem_HasTextElement Or elem_HasValueList, "true", "System Aware Only", "false", "Unaware", "true/pm", "Win 8.1/10 per-monitor aware, lower O/S system aware", "per monitor", "Win 8.1/10 per-monitor aware, lower O/S unaware"
            Case "dpiAwareness" ' offer user-friendly values
                m_Element.AddElementValueList elem_Fixed Or elem_ManualEntry Or elem_HasTextElement Or elem_HasValueList Or elem_CanBeBlank, "system", "System Aware", "permonitor", "Per-Monitor Aware", "permonitorv2", "Per-Monitor Aware (v1703 only)", "unaware", "Unaware"
            Case Else
                m_Element.AddElementValueList elem_Fixed Or elem_HasTextElement
            End Select
        Case "windowClass"
            m_Element.AddElementValueList elem_HasTextElement Or elem_ManualEntry
        Case "comClass"
            If .baseName = "progid" Then
                m_Element.AddElementValueList elem_HasTextElement Or elem_ManualEntry
                If .Text = vbNullString Then .Text = "Program.ID"
            End If
        Case Else
            bNotUpdated = True
        End Select
        If bNotUpdated Then
            If isImportedElement Then                               ' allow updating value(s)
                m_Element.AddElementValueList elem_ManualEntry Or elem_CanBeBlank
            Else
                m_Element.AddElementValueList elem_Fixed            ' prevent updating value(s)
            End If
        End If
    End With
    
End Sub

Private Sub pvValidateAttributes(xmlElement As IXMLDOMElement, isImportedElement As Boolean)

    ' routine applies properties of the passed element's attributes
    ' note: attr_Fixed applied below means the user cannot change the element's value
    ' enumeration comes from cManifestEntryEx

    Dim a As Long
    With xmlElement
        Select Case .baseName
        Case "assemblyIdentity" ' name, version & type are required attributes; ensure they exist & not null
            If isImportedElement Then
                If IsNull(.getAttribute("name")) Then
                    .setAttribute "name", "Required.Entry"
                ElseIf .getAttribute("name") = vbNullString Then
                    .setAttribute "name", "Required.Entry"
                End If
                If IsNull(.getAttribute("version")) Then
                    .setAttribute "version", "0.0.0.0"
                ElseIf .getAttribute("version") = vbNullString Then
                    .setAttribute "version", "0.0.0.0"
                End If
                If IsNull(.getAttribute("type")) Then
                    .setAttribute "type", "win32"
                ElseIf .getAttribute("type") = vbNullString Then
                    .setAttribute "type", "win32"
                End If
                If IsNull(.getAttribute("language")) Then .setAttribute "language", ""
                If IsNull(.getAttribute("processorArchitecture")) Then .setAttribute "processorArchitecture", "x86"
                If IsNull(.getAttribute("publicKeyToken")) Then .setAttribute "publicKeyToken", ""
            End If
            With m_Element
                .AddAttributeValueList "name", attr_ManualEntry Or attr_Required
                .AddAttributeValueList "version", attr_ManualEntry Or attr_Required
                .AddAttributeValueList "type", attr_Fixed
                .AddAttributeValueList "language", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "processorArchitecture", attr_HasValueList Or attr_CanBeBlank, "x86", "32 Bit", "ia64", "64 Bit"
                .AddAttributeValueList "publicKeyToken", attr_ManualEntry Or attr_CanBeBlank
            End With
        Case "supportedOS"  ' Id attribute is required, ensure it exists & not null
            If IsNull(.getAttribute("Id")) Then
                .setAttribute "Id", "{00000000-0000-0000-0000-000000000000}"
                m_Element.AddAttributeValueList "Id", attr_ManualEntry Or attr_Required
            ElseIf .getAttribute("Id") = vbNullString Then
                .setAttribute "Id", "{00000000-0000-0000-0000-000000000000}"
                m_Element.AddAttributeValueList "Id", attr_ManualEntry Or attr_Required
            Else
                m_Element.AddAttributeValueList "Id", attr_Fixed
                Select Case .getAttribute("Id")
                    Case "{e2011457-1546-43c5-a5fe-008deee3d3f0}"   ' Vista
                    Case "{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"   ' Win 7
                    Case "{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"   ' Win 8
                    Case "{1f676c76-80e1-4239-95bb-83d0f6d0da78}"   ' Win 8.1
                    Case "{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"   ' Win 10
                    Case Else: m_Element.AddAttributeValueList "Id", attr_ManualEntry Or attr_Required
                End Select
            End If
        Case "file"     ' name attribute is required, ensure it exists & is not null
            If isImportedElement Then
                If IsNull(.getAttribute("name")) Then
                    .setAttribute "name", "FileName.ext"
                ElseIf .getAttribute("name") = vbNullString Then
                    .setAttribute "name", "FileName.ext"
                End If
                If IsNull(.getAttribute("hashalg")) Then .setAttribute "hashalg", ""
                If IsNull(.getAttribute("hash")) Then .setAttribute "hash", ""
            End If
            With m_Element
                .AddAttributeValueList "name", attr_Required Or attr_ManualEntry
                .AddAttributeValueList "hashalg", attr_CanBeBlank Or attr_HasValueList Or attr_ManualEntry, "SHA1", "Secure Hash Algorithm 1"
                .AddAttributeValueList "hash", attr_CanBeBlank Or attr_ManualEntry
            End With
        Case "requestedExecutionLevel"
            If isImportedElement Then   ' not definitive, but ensure these two attributes exist & not null
                If IsNull(.getAttribute("level")) Then
                    .setAttribute "level", "asInvoker"
                ElseIf .getAttribute("level") = vbNullString Then
                    .setAttribute "level", "asInvoker"
                End If
                If IsNull(.getAttribute("uiAccess")) Then
                    .setAttribute "uiAccess", "false"
                ElseIf .getAttribute("uiAccess") = vbNullString Then
                    .setAttribute "uiAccess", "false"
                End If
            End If
            With m_Element
                .AddAttributeValueList "level", attr_HasValueList, "asInvoker", "User Standard Rights", "highestAvailable", "User's Highest Rights", "requireAdministrator", "Admin Rights Required"
                .AddAttributeValueList "uiAccess", attr_HasValueList, "true", "True", "false", "False"
            End With
        Case "typelib"
            If IsNull(.getAttribute("tlbid")) Then
                .setAttribute "tlbid", "{00000000-0000-0000-0000-000000000000}"
            ElseIf .getAttribute("tlbid") = vbNullString Then
                .setAttribute "tlbid", "{00000000-0000-0000-0000-000000000000}"
            End If
            If IsNull(.getAttribute("helpdir")) Then
                .setAttribute "helpdir", ""
            ElseIf .getAttribute("helpdir") = vbNullString Then
                .setAttribute "helpdir", ""
            End If
            If IsNull(.getAttribute("version")) Then
                .setAttribute "version", "1.0"
            ElseIf .getAttribute("version") = vbNullString Then
                .setAttribute "version", "1.0"
            End If
            If IsNull(.getAttribute("resourceid")) Then .setAttribute "resourceid", ""
            If IsNull(.getAttribute("flags")) Then .setAttribute "flags", ""
            With m_Element
                .AddAttributeValueList "tlbid", attr_ManualEntry Or attr_Required
                .AddAttributeValueList "helpdir", attr_ManualEntry Or attr_Required Or attr_CanBeBlank
                .AddAttributeValueList "version", attr_ManualEntry Or attr_Required Or attr_CanBeBlank
                .AddAttributeValueList "resourceid", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "flags", attr_CanBeBlank Or attr_HasValueList Or attr_ManualEntry, "RESTRICTED", "", "CONTROL", "", "HIDDEN", "", "HASDISKIMAGE", ""
            End With
        Case "comClass"
            If IsNull(.getAttribute("clsid")) Then
                .setAttribute "clsid", "{00000000-0000-0000-0000-000000000000}"
            ElseIf .getAttribute("clsid") = vbNullString Then
                .setAttribute "clsid", "{00000000-0000-0000-0000-000000000000}"
            End If
            If IsNull(.getAttribute("tlbid")) Then .setAttribute "tlbid", ""
            If IsNull(.getAttribute("progid")) Then .setAttribute "progid", ""
            If IsNull(.getAttribute("threadingModel")) Then .setAttribute "threadingModel", ""
            If IsNull(.getAttribute("description")) Then .setAttribute "description", ""
            If IsNull(.getAttribute("miscStatus")) Then .setAttribute "miscStatus", ""
            If IsNull(.getAttribute("miscStatusIcon")) Then .setAttribute "miscStatusIcon", ""
            If IsNull(.getAttribute("miscStatusContent")) Then .setAttribute "miscStatusContent", ""
            If IsNull(.getAttribute("miscStatusDocprint")) Then .setAttribute "miscStatusDocprint", ""
            If IsNull(.getAttribute("miscStatusThumbnail")) Then .setAttribute "miscStatusThumbnail", ""
            With m_Element
                .AddAttributeValueList "clsid", attr_ManualEntry Or attr_Required
                .AddAttributeValueList "threadingModel", attr_HasValueList Or attr_CanBeBlank, "Apartment", "", "Free", "", "Both", "", "Neutral", ""
                .AddAttributeValueList "tlbid", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "progid", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "description", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "miscStatus", attr_ManualEntry Or attr_CanBeBlank Or attr_HasValueList, _
                    "activatewhenvisible", "OLEMISC_ACTIVATEWHENVISIBLE", "actslikebutton", "OLEMISC_ACTSLIKEBUTTON", "actslikelabel", "OLEMISC_ACTSLIKELABEL", "alignable", "OLEMISC_ALIGNABLE", "alwaysrun", "OLEMISC_ALWAYSRUN", "canlinkbyole1", "OLEMISC_CANLINKBYOLE1", "cantlinkinside", "OLEMISC_CANTLINKINSIDE", "ignoreativatewhenvisible", "OLEMISC_IGNOREACTIVATEWHENVISIBLE", "insertnotreplace", "OLEMISC_INSERTNOTREPLACE", "insideout", "OLEMISC_INSIDEOUT", "invisibleatruntime", "OLEMISC_INVISIBLEATRUNTIME", "islinkobject", "OLEMISC_ISLINKOBJECT", "nouiactivate", "OLEMISC_NOUIACTIVATE", "onlyiconic", "OLEMISC_ONLYICONIC", "recomposeonresize", "OLEMISC_RECOMPOSEONRESIZE", "renderingisdeviceindependent", "OLEMISC_RENDERINGISDEVICEINDEPENDENT", "setclientsitefirst", "OLEMISC_SETCLIENTSITEFIRST", "simpleframe", "OLEMISC_SIMPLEFRAME", "static", "OLEMISC_STATIC", "supportsmultilevelundo", "OLEMISC_SUPPORTSMULTILEVELUNDO", "wantstomenumerge", "OLEMISC_WANTSTOMENUMERGE", "imemode", "TOLEMISC_IMEMODE"
                .AddAttributeValueList "miscStatusIcon", attr_ManualEntry Or attr_CanBeBlank Or attr_HasValueList, _
                    "activatewhenvisible", "OLEMISC_ACTIVATEWHENVISIBLE", "actslikebutton", "OLEMISC_ACTSLIKEBUTTON", "actslikelabel", "OLEMISC_ACTSLIKELABEL", "alignable", "OLEMISC_ALIGNABLE", "alwaysrun", "OLEMISC_ALWAYSRUN", "canlinkbyole1", "OLEMISC_CANLINKBYOLE1", "cantlinkinside", "OLEMISC_CANTLINKINSIDE", "ignoreativatewhenvisible", "OLEMISC_IGNOREACTIVATEWHENVISIBLE", "insertnotreplace", "OLEMISC_INSERTNOTREPLACE", "insideout", "OLEMISC_INSIDEOUT", "invisibleatruntime", "OLEMISC_INVISIBLEATRUNTIME", "islinkobject", "OLEMISC_ISLINKOBJECT", "nouiactivate", "OLEMISC_NOUIACTIVATE", "onlyiconic", "OLEMISC_ONLYICONIC", "recomposeonresize", "OLEMISC_RECOMPOSEONRESIZE", "renderingisdeviceindependent", "OLEMISC_RENDERINGISDEVICEINDEPENDENT", "setclientsitefirst", "OLEMISC_SETCLIENTSITEFIRST", "simpleframe", "OLEMISC_SIMPLEFRAME", "static", "OLEMISC_STATIC", "supportsmultilevelundo", "OLEMISC_SUPPORTSMULTILEVELUNDO", "wantstomenumerge", "OLEMISC_WANTSTOMENUMERGE", "imemode", "TOLEMISC_IMEMODE"
                .AddAttributeValueList "miscStatusContent", attr_ManualEntry Or attr_CanBeBlank Or attr_HasValueList, _
                    "activatewhenvisible", "OLEMISC_ACTIVATEWHENVISIBLE", "actslikebutton", "OLEMISC_ACTSLIKEBUTTON", "actslikelabel", "OLEMISC_ACTSLIKELABEL", "alignable", "OLEMISC_ALIGNABLE", "alwaysrun", "OLEMISC_ALWAYSRUN", "canlinkbyole1", "OLEMISC_CANLINKBYOLE1", "cantlinkinside", "OLEMISC_CANTLINKINSIDE", "ignoreativatewhenvisible", "OLEMISC_IGNOREACTIVATEWHENVISIBLE", "insertnotreplace", "OLEMISC_INSERTNOTREPLACE", "insideout", "OLEMISC_INSIDEOUT", "invisibleatruntime", "OLEMISC_INVISIBLEATRUNTIME", "islinkobject", "OLEMISC_ISLINKOBJECT", "nouiactivate", "OLEMISC_NOUIACTIVATE", "onlyiconic", "OLEMISC_ONLYICONIC", "recomposeonresize", "OLEMISC_RECOMPOSEONRESIZE", "renderingisdeviceindependent", "OLEMISC_RENDERINGISDEVICEINDEPENDENT", "setclientsitefirst", "OLEMISC_SETCLIENTSITEFIRST", "simpleframe", "OLEMISC_SIMPLEFRAME", "static", "OLEMISC_STATIC", "supportsmultilevelundo", "OLEMISC_SUPPORTSMULTILEVELUNDO", "wantstomenumerge", "OLEMISC_WANTSTOMENUMERGE", "imemode", "TOLEMISC_IMEMODE"
                .AddAttributeValueList "miscStatusDocprint", attr_ManualEntry Or attr_CanBeBlank Or attr_HasValueList, _
                    "activatewhenvisible", "OLEMISC_ACTIVATEWHENVISIBLE", "actslikebutton", "OLEMISC_ACTSLIKEBUTTON", "actslikelabel", "OLEMISC_ACTSLIKELABEL", "alignable", "OLEMISC_ALIGNABLE", "alwaysrun", "OLEMISC_ALWAYSRUN", "canlinkbyole1", "OLEMISC_CANLINKBYOLE1", "cantlinkinside", "OLEMISC_CANTLINKINSIDE", "ignoreativatewhenvisible", "OLEMISC_IGNOREACTIVATEWHENVISIBLE", "insertnotreplace", "OLEMISC_INSERTNOTREPLACE", "insideout", "OLEMISC_INSIDEOUT", "invisibleatruntime", "OLEMISC_INVISIBLEATRUNTIME", "islinkobject", "OLEMISC_ISLINKOBJECT", "nouiactivate", "OLEMISC_NOUIACTIVATE", "onlyiconic", "OLEMISC_ONLYICONIC", "recomposeonresize", "OLEMISC_RECOMPOSEONRESIZE", "renderingisdeviceindependent", "OLEMISC_RENDERINGISDEVICEINDEPENDENT", "setclientsitefirst", "OLEMISC_SETCLIENTSITEFIRST", "simpleframe", "OLEMISC_SIMPLEFRAME", "static", "OLEMISC_STATIC", "supportsmultilevelundo", "OLEMISC_SUPPORTSMULTILEVELUNDO", "wantstomenumerge", "OLEMISC_WANTSTOMENUMERGE", "imemode", "TOLEMISC_IMEMODE"
                .AddAttributeValueList "miscStatusThumbnail", attr_ManualEntry Or attr_CanBeBlank Or attr_HasValueList, _
                    "activatewhenvisible", "OLEMISC_ACTIVATEWHENVISIBLE", "actslikebutton", "OLEMISC_ACTSLIKEBUTTON", "actslikelabel", "OLEMISC_ACTSLIKELABEL", "alignable", "OLEMISC_ALIGNABLE", "alwaysrun", "OLEMISC_ALWAYSRUN", "canlinkbyole1", "OLEMISC_CANLINKBYOLE1", "cantlinkinside", "OLEMISC_CANTLINKINSIDE", "ignoreativatewhenvisible", "OLEMISC_IGNOREACTIVATEWHENVISIBLE", "insertnotreplace", "OLEMISC_INSERTNOTREPLACE", "insideout", "OLEMISC_INSIDEOUT", "invisibleatruntime", "OLEMISC_INVISIBLEATRUNTIME", "islinkobject", "OLEMISC_ISLINKOBJECT", "nouiactivate", "OLEMISC_NOUIACTIVATE", "onlyiconic", "OLEMISC_ONLYICONIC", "recomposeonresize", "OLEMISC_RECOMPOSEONRESIZE", "renderingisdeviceindependent", "OLEMISC_RENDERINGISDEVICEINDEPENDENT", "setclientsitefirst", "OLEMISC_SETCLIENTSITEFIRST", "simpleframe", "OLEMISC_SIMPLEFRAME", "static", "OLEMISC_STATIC", "supportsmultilevelundo", "OLEMISC_SUPPORTSMULTILEVELUNDO", "wantstomenumerge", "OLEMISC_WANTSTOMENUMERGE", "imemode", "TOLEMISC_IMEMODE"
            End With
        Case "comInterfaceProxyStub"
            If IsNull(.getAttribute("iid")) Then
                .setAttribute "iid", "{00000000-0000-0000-0000-000000000000}"
            ElseIf .getAttribute("iid") = vbNullString Then
                .setAttribute "iid", "{00000000-0000-0000-0000-000000000000}"
            End If
            If IsNull(.getAttribute("name")) Then
                .setAttribute "name", "Name"
            ElseIf .getAttribute("name") = vbNullString Then
                .setAttribute "name", "Name"
            End If
            If IsNull(.getAttribute("threadingModel")) Then .setAttribute "threadingModel", ""
            If IsNull(.getAttribute("threadingModel")) Then .setAttribute "threadingModel", ""
            If IsNull(.getAttribute("baseInterface")) Then .setAttribute "baseInterface", ""
            If IsNull(.getAttribute("numMethods")) Then .setAttribute "numMethods", ""
            If IsNull(.getAttribute("proxyStubClsid32")) Then .setAttribute "proxyStubClsid32", ""
            With m_Element
                .AddAttributeValueList "iid", attr_ManualEntry Or attr_Required
                .AddAttributeValueList "threadingModel", attr_HasValueList Or attr_CanBeBlank, "Apartment", "", "Free", "", "Both", "", "Neutral", ""
                .AddAttributeValueList "name", attr_ManualEntry Or attr_Required
                .AddAttributeValueList "baseInterface", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "numMethods", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "proxyStubClsid32", attr_ManualEntry Or attr_CanBeBlank
            End With
        Case "comInterfaceExternalProxyStub"
            If IsNull(.getAttribute("iid")) Then
                .setAttribute "iid", "{00000000-0000-0000-0000-000000000000}"
            ElseIf .getAttribute("iid") = vbNullString Then
                .setAttribute "iid", "{00000000-0000-0000-0000-000000000000}"
            End If
            If IsNull(.getAttribute("name")) Then
                .setAttribute "name", "Name"
            ElseIf .getAttribute("name") = vbNullString Then
                .setAttribute "name", "Name"
            End If
            If IsNull(.getAttribute("tlbid")) Then .setAttribute "tlbid", ""
            If IsNull(.getAttribute("numMethods")) Then .setAttribute "numMethods", ""
            If IsNull(.getAttribute("threadingModel")) Then .setAttribute "threadingModel", ""
            If IsNull(.getAttribute("baseInterface")) Then .setAttribute "baseInterface", ""
            If IsNull(.getAttribute("proxyStubClsid32")) Then .setAttribute "proxyStubClsid32", ""
            With m_Element
                .AddAttributeValueList "iid", attr_ManualEntry Or attr_Required
                .AddAttributeValueList "name", attr_ManualEntry Or attr_Required
                .AddAttributeValueList "tlbid", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "threadingModel", attr_HasValueList Or attr_CanBeBlank, "Apartment", "", "Free", "", "Both", "", "Neutral", ""
                .AddAttributeValueList "baseInterface", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "numMethods", attr_ManualEntry Or attr_CanBeBlank
                .AddAttributeValueList "proxyStubClsid32", attr_ManualEntry Or attr_CanBeBlank
            End With
        Case "windowClass"
            If .Text = vbNullString Then .Text = "Class Name"
            If IsNull(.getAttribute("versioned")) Then .setAttribute "versioned", ""
            With m_Element
                .AddAttributeValueList "versioned", attr_CanBeBlank Or attr_HasValueList, "yes", "Yes", "no", "No"
            End With
        Case Else   ' for any other element attributes (unknown); allow user to update them on demand
            For a = 0& To xmlElement.Attributes.length - 1&
                m_Element.AddAttributeValueList a, attr_ManualEntry Or attr_CanBeBlank
            Next
        End Select
    End With

End Sub

Private Function pvGetInsertPos(pElement As IXMLDOMElement, nodeName As String) As IXMLDOMNode

    ' insert position could help reduce namespace declarations if manifest is exported without qualified names
    ' in any case, prefer to keep the manifest exported with contiguous namespace entries if possible

    Dim xmlElement As IXMLDOMNode, ns As String
    ' find the exact element name, if it exists
    Set xmlElement = pElement.selectSingleNode(nodeName & "[last()]")
    If xmlElement Is Nothing Then
        ' find the last element with same namespace
        If InStr(nodeName, ":") Then ns = Me.GetNameSpace(Left$(nodeName, InStr(nodeName, ":") - 1))
        Set xmlElement = pElement.selectSingleNode("*[namespace-uri()='" & ns & "'][last()]")
        If Not xmlElement Is Nothing Then Set pvGetInsertPos = xmlElement.nextSibling
    Else    ' use that element as the insert position
        Set pvGetInsertPos = xmlElement.nextSibling
    End If

End Function

Private Sub Class_Terminate()
    Set m_Element = Nothing
    Set m_ElementProps = Nothing
    Set m_XML = Nothing
    Set m_nsLUT = Nothing
End Sub
