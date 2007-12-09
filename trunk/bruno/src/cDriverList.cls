VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDriverList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

 ' Daisy 2.02 Validator, Daisy 2.02 Regenerator, Bruno
 ' The Daisy Visual Basic Tool Suite
 ' Copyright (C) 2003,2004,2005,2006,2007,2008 Daisy Consortium
 '
 ' This library is free software; you can redistribute it and/or
 ' modify it under the terms of the GNU Lesser General Public
 ' License as published by the Free Software Foundation; either
 ' version 2.1 of the License, or (at your option) any later version.
 '
 ' This library is distributed in the hope that it will be useful,
 ' but WITHOUT ANY WARRANTY; without even the implied warranty of
 ' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 ' Lesser General Public License for more details.
 '
 ' You should have received a copy of the GNU Lesser General Public
 ' License along with this library; if not, write to the Free Software
 ' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

'contains a listing of filters available for selection
'the actual selected Driver sits in cDriver

Private aDrivers() As cDriverInfo
Public DriverCount As Long

Property Get Driver(i As Long) As cDriverInfo
  Set Driver = aDrivers(i)
End Property

Private Sub Class_Initialize()
Dim oDom As New MSXML2.DOMDocument40
    oDom.async = False
    oDom.validateOnParse = True
    oDom.resolveExternals = True
    oDom.preserveWhiteSpace = False
    oDom.setProperty "SelectionLanguage", "XPath"
    oDom.setProperty "SelectionNamespaces", "xmlns:bruno='http://www.daisy.org/2004/bruno'"
    oDom.setProperty "NewParser", False
Dim oFso As Object
Dim oFolder As Object, oFiles As Object, oFile As Object
Dim oName As IXMLDOMNode, oDesc As IXMLDOMNode
Dim sXmlParseMessage As String
    
  Set oFso = CreateObject("scripting.fileSystemObject")
  
  Set oFolder = oFso.GetFolder(oBruno.oPaths.DriverPath)
  If Not oFolder Is Nothing Then
    Set oFiles = oFolder.Files
    For Each oFile In oFolder.Files
      If oFso.getExtensionName(oFile.Path) = "drvr" Then
        If oBruno.oCmn.oDomCmn.fncParseFile(oFile.Path, oDom, sXmlParseMessage) Then
          Set oName = oDom.selectSingleNode("//bruno:name")
          Set oDesc = oDom.selectSingleNode("//bruno:desc")
          If Not oName Is Nothing Then
             ReDim Preserve aDrivers(DriverCount)
             Set aDrivers(DriverCount) = New cDriverInfo
             With aDrivers(DriverCount)
              .sName = oName.Text
              .sFullPath = oFile.Path
              If Not oDesc Is Nothing Then .sDesc = oDesc.Text
             End With
          End If
          DriverCount = DriverCount + 1
        Else
          frmMain.fncAddMessage "Driver file " & oFile.Path & " could not be loaded." & sXmlParseMessage
        End If
      End If
    Next
  End If
  Set oFso = Nothing
  Set oDom = Nothing
End Sub

