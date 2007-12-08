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
