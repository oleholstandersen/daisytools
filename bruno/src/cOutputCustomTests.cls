VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutputCustomTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Public oTests As New MSXML2.DOMDocument40
Public bHasTests As Boolean

Private Sub Class_Initialize()
 'oTests exposes all customtests in smil presentation
 oTests.async = False
 oTests.validateOnParse = False
 oTests.resolveExternals = False
 oTests.preserveWhiteSpace = False
 oTests.setProperty "SelectionLanguage", "XPath"
 oTests.setProperty "NewParser", True
 oTests.loadXML ("<customTests></customTests>")
End Sub

Public Function fncAddCustomTest(oSmilCustomTestElement As IXMLDOMElement) As Boolean
Dim oCustomTest As IXMLDOMNode, oCustomTestId As IXMLDOMAttribute
Dim bAddbookStruct As Boolean, sBookStructValue As String
  'adds the test to oTests if not already existing
  'returns true if test was added
  If Not (fncTestExists(oSmilCustomTestElement)) Then
    Set oCustomTest = oTests.documentElement.appendChild(oBruno.oCmn.oDomCmn.fncRenameElement(oSmilCustomTestElement, "smilCustomTest", ""))
    
    'add the bookStruct attribute if in z2004 enum list
    Set oCustomTestId = oCustomTest.selectSingleNode("@id")
    If Not oCustomTestId Is Nothing Then
      bAddbookStruct = True
      Select Case oCustomTestId.Text
        Case "pagenum"
          sBookStructValue = "PAGE_NUMBER"
        Case "note"
          sBookStructValue = "NOTE"
        Case "noteref"
          sBookStructValue = "NOTE_REFERENCE"
        Case "annotation"
          sBookStructValue = "ANNOTATION"
        Case "linenum"
          sBookStructValue = "LINE_NUMBER"
        Case "sidebar"
          sBookStructValue = "OPTIONAL_SIDEBAR"
        Case "prodnote"
          sBookStructValue = "OPTIONAL_PRODUCER_NOTE"
        Case Else
          'dont add bookStruct
          bAddbookStruct = False
      End Select
      If bAddbookStruct Then oBruno.oCmn.oDomCmn.fncAppendAttribute oCustomTest, "bookStruct", sBookStructValue
    End If
    
    Me.bHasTests = True
    fncAddCustomTest = True
  Else
    fncAddCustomTest = False
  End If
End Function

Private Function fncTestExists(oInSmilCustomTestElement As IXMLDOMElement) As Boolean
Dim oNode As IXMLDOMNode
  fncTestExists = True
  For Each oNode In oTests.documentElement.childNodes
    If oNode.selectSingleNode("@id").Text = _
      oInSmilCustomTestElement.selectSingleNode("@id").Text Then
      Exit Function
    End If
  Next
  fncTestExists = False
End Function
