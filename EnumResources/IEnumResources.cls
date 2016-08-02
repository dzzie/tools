VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©2000 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

Public Sub EnumResourceSink(ByVal hModule As Long, ByVal ResName As String, ByVal ResType As String, Continue As Boolean)
   ' This routine is called by the MEnumResources module
   ' once for each resource type found and once for each
   ' named resource of that type.
   '
   ' Allows enumeration to be handled in class or form
   ' modules.  Enumeration continues until all resources
   ' have been found or Continue is set to False.
End Sub

