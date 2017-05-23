        ��  ��                  h      �� ��     0 	        <?xml version="1.0" standalone="yes"?>
<asmv1:assembly manifestVersion="1.0" xmlns:asmv1="urn:schemas-microsoft-com:asm.v1" xmlns:asmv2="urn:schemas-microsoft-com:asm.v2" xmlns:comp="urn:schemas-microsoft-com:compatibility.v1" xmlns:asmv3="urn:schemas-microsoft-com:asm.v3" xmlns:ws2005="http://schemas.microsoft.com/SMI/2005/WindowsSettings" xmlns:ws2016="http://schemas.microsoft.com/SMI/2016/WindowsSettings"><asmv1:assemblyIdentity name="Manifest.Creator.II" version="2.0.0.0" type="win32" processorArchitecture="x86"/><asmv1:description>Application-Manifest Creation Tool</asmv1:description><asmv1:dependency><asmv1:dependentAssembly><asmv1:assemblyIdentity name="Microsoft.Windows.Common-Controls" version="6.0.0.0" type="win32" processorArchitecture="x86" publicKeyToken="6595b64144ccf1df" language="*"/></asmv1:dependentAssembly></asmv1:dependency><asmv2:trustInfo><asmv2:security><asmv2:requestedPrivileges><asmv2:requestedExecutionLevel level="highestAvailable" uiAccess="false"/></asmv2:requestedPrivileges></asmv2:security></asmv2:trustInfo><comp:compatibility><comp:application><comp:supportedOS Id="{e2011457-1546-43c5-a5fe-008deee3d3f0}"/><comp:supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"/><comp:supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"/><comp:supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/><comp:supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/></comp:application></comp:compatibility><asmv3:application><asmv3:windowsSettings><ws2005:dpiAware>true</ws2005:dpiAware><ws2016:dpiAwareness>system</ws2016:dpiAwareness></asmv3:windowsSettings></asmv3:application></asmv1:assembly>
  8   C U S T O M   S U B M A I N         0	        Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type
Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean

Private Sub Main()

    Dim iccex As InitCommonControlsExStruct, hMod As Long
    Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all known values
    ' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx

    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_ALL_CLASSES    ' you really should customize this value from the available constants
    End With
    On Error Resume Next ' error? Requires IEv3 or above
    hMod = LoadLibrary("shell32.dll")
    InitCommonControlsEx iccex
    If Err Then
        InitCommonControls ' try Win9x version
        Err.Clear
    End If
    On Error GoTo 0
    '... show your main form next (i.e., Form1.Show)
    frmMain.Show
    If hMod Then FreeLibrary hMod


'** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'          In place of VB Frames, use pictureboxes instead.
'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'          Doing so will prevent them from being themed.

End Sub i  @   C U S T O M   X S L T - S U B L S T         0	        <xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output omit-xml-declaration="no"/>
<xsl:strip-space elements="*"/>
<xsl:param name="subSelectionXPath" select="//*[@mcIIid='107018568']" />
<xsl:template match="*">
<xsl:choose>
<xsl:when test="descendant::node()[count(.|$subSelectionXPath)=count($subSelectionXPath)]">
<xsl:copy><xsl:copy-of select="@*"/>
<xsl:apply-templates select="*"/>
</xsl:copy></xsl:when>
<xsl:when test="count(.|$subSelectionXPath)=count($subSelectionXPath)">
<xsl:copy-of select="."/></xsl:when>
</xsl:choose></xsl:template>
</xsl:stylesheet>   �
  0   C U S T O M   B A S E       0	        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<asmv1:assembly manifestVersion="1.0" xmlns:asmv1="urn:schemas-microsoft-com:asm.v1" xmlns:asmv2="urn:schemas-microsoft-com:asm.v2" xmlns:comp="urn:schemas-microsoft-com:compatibility.v1" xmlns:asmv3="urn:schemas-microsoft-com:asm.v3" xmlns:ws2005="http://schemas.microsoft.com/SMI/2005/WindowsSettings" xmlns:ws2011="http://schemas.microsoft.com/SMI/2011/WindowsSettings" xmlns:ws2013="http://schemas.microsoft.com/SMI/2013/WindowsSettings" xmlns:ws2016="http://schemas.microsoft.com/SMI/2016/WindowsSettings" xmlns:ws2017="http://schemas.microsoft.com/SMI/2017/WindowsSettings">
<asmv1:noInherit/>
<asmv1:assemblyIdentity name="My.Cool.New.Application" version="1.0.0.0" type="win32" processorArchitecture="x86" publicKeyToken="" language=""/>
<asmv1:description/>
<asmv1:dependency>
<asmv1:dependentAssembly>
<asmv1:assemblyIdentity name="Microsoft.Windows.Common-Controls" version="6.0.0.0" type="win32" processorArchitecture="x86" publicKeyToken="6595b64144ccf1df" language="*"/>
</asmv1:dependentAssembly>
</asmv1:dependency>
<asmv1:dependency>
<asmv1:dependentAssembly>
<asmv1:assemblyIdentity name="Microsoft.Windows.GdiPlus" version="1.1.0.0" type="win32" processorArchitecture="x86" publicKeyToken="6595b64144ccf1df" language="*"/>
</asmv1:dependentAssembly>
</asmv1:dependency>
<asmv1:file name="FileName" hashalg="" hash=""/>
<asmv2:trustInfo>
<asmv2:security>
<asmv2:requestedPrivileges>
<asmv2:requestedExecutionLevel level="asInvoker" uiAccess="false"/>
</asmv2:requestedPrivileges>
</asmv2:security>
</asmv2:trustInfo>
<comp:compatibility>
<comp:application>
<comp:supportedOS Id="{e2011457-1546-43c5-a5fe-008deee3d3f0}"/>
<comp:supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"/>
<comp:supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"/>
<comp:supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/>
<comp:supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
</comp:application>
</comp:compatibility>
<asmv3:application>
<asmv3:windowsSettings>
<ws2005:dpiAware>true</ws2005:dpiAware>
<ws2005:autoElevate>true</ws2005:autoElevate>
<ws2005:disableTheming>true</ws2005:disableTheming>
<ws2011:highResolutionScrollingAware>true</ws2011:highResolutionScrollingAware>
<ws2011:disableWindowFiltering>true</ws2011:disableWindowFiltering>
<ws2013:ultraHighResolutionScrollingAware>true</ws2013:ultraHighResolutionScrollingAware>
<ws2013:printerDriverIsolation>true</ws2013:printerDriverIsolation>
<ws2016:dpiAwareness>system</ws2016:dpiAwareness>
<ws2016:longPathAware>true</ws2016:longPathAware>
<ws2017:gdiScaling>true</ws2017:gdiScaling>
</asmv3:windowsSettings>
</asmv3:application>
</asmv1:assembly>  