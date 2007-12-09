Attribute VB_Name = "mPrettyPrint"
' Daisy 2.02 Regenerator DLL
' Copyright (C) 2003 Daisy Consortium
'
'    This file is part of Daisy 2.02 Regenerator.
'
'    Daisy 2.02 Regenerator is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Daisy 2.02 Regenerator is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Daisy 2.02 Regenerator; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


Option Explicit

Public Function fncPrettyPrint( _
    ByRef objOwner As oRegenerator, _
    ByRef bolPb2kLayoutFix As Boolean _
    ) As Boolean

Dim i As Long
Dim oDom As New MSXML2.DOMDocument40
    oDom.async = False
    oDom.validateOnParse = False
    oDom.resolveExternals = False
    oDom.preserveWhiteSpace = False
    oDom.setProperty "SelectionLanguage", "XPath"
    oDom.setProperty "NewParser", True
  
  On Error GoTo ErrHandler
  fncPrettyPrint = False
  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_NCC Or _
       objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1 Or _
       objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_CONTENT Or _
       objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_MASTER Then
      If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(i).sDomData, _
        oDom, objOwner) Then GoTo ErrHandler
            
      If Not fncXmlIndentSax(oDom, objOwner) Then GoTo ErrHandler
      
      'pb2k fix - makes ncc items occur on one line
      If bolPb2kLayoutFix Then
        If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_NCC Then
          If Not fncRemoveCrLf(oDom, "//a", objOwner) Then GoTo ErrHandler
          If Not fncRemoveCrLf(oDom, "//head/title", objOwner) Then GoTo ErrHandler
        End If
      End If
                  
      objOwner.objFileSetHandler.aOutFileSet(i).sDomData = oDom.xml
      
    End If 'aOutFileSet(i).eType =
  Next i
  fncPrettyPrint = True
ErrHandler:
  If Not fncPrettyPrint Then objOwner.addlog "<errH in='fncPrettyPrint'>fncPrettyPrint ErrH" & Err.Description & "</errH>"
End Function

Public Function fncXmlIndentSax( _
  ByRef oDom As MSXML2.DOMDocument40, _
  ByVal objOwner As oRegenerator _
  ) As Boolean

Dim rdr As New SAXXMLReader40
Dim wrt As New MXXMLWriter40
  
  On Error GoTo ErrHandler
  fncXmlIndentSax = False
  Set rdr.contentHandler = wrt
  Set rdr.dtdHandler = wrt
  Set rdr.errorHandler = wrt
  rdr.putFeature "preserve-system-identifiers", True
  rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt
  rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
  rdr.putFeature "schema-validation", False
  
  wrt.output = ""
  wrt.byteOrderMark = False

  wrt.standalone = True
  wrt.indent = True
  wrt.omitXMLDeclaration = False

  rdr.parse oDom
  
  If Not fncParseString( _
    Replace(wrt.output, "standalone=" & Chr(34) & "yes" & Chr(34), ""), oDom, objOwner) Then GoTo ErrHandler
  
  fncXmlIndentSax = True
ErrHandler:
  If Not fncXmlIndentSax Then objOwner.addlog "<errH in='fncXmlIndentSax'>fncXmlIndentSax ErrH</errH>"
End Function

Public Function fncRemoveCrLf( _
    ByRef oDom As IXMLDOMDocument, _
    ByVal sXpath As String, _
    ByVal objOwner As oRegenerator _
    ) As Boolean

Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim oNew As IXMLDOMText

  On Error GoTo ErrHandler
  fncRemoveCrLf = False

  'in the ncc, select all anchors
  Set oNodes = oDom.documentElement.selectNodes(sXpath)
  
  'remove the crlf before the node open
  If Not oNodes Is Nothing Then
    If sXpath <> "//head/title" Then
        For Each oNode In oNodes
          If Not oNode.previousSibling Is Nothing Then
            If oNode.previousSibling.nodeType = NODE_TEXT Then
              'only del the node if its only whitespace
              If (Trim$(oNode.previousSibling.Text) = "") Then
                oNode.previousSibling.nodeValue = vbNullString
              End If
            End If
          Else 'if previoussbling is nothing insert vbnullstring
            Set oNew = oDom.createTextNode("")
            Set oNew = oNode.parentNode.insertBefore(oNew, oNode)
            Set oNew = Nothing
          End If 'Not oNode.previousSibling is nothing
        Next
        
        'remove the crlf after the node close
        For Each oNode In oNodes
          If Not oNode.nextSibling Is Nothing Then
            If oNode.nextSibling.nodeType = NODE_TEXT Then
              'only del the node if its only whitespace
              If Trim$(oNode.nextSibling.Text) = "" Then
                oNode.nextSibling.nodeValue = vbNullString
              End If
            End If
          Else 'if nextsibling is nothing insert vbnullstring
            Set oNew = oDom.createTextNode("")
            Set oNew = oNode.parentNode.appendChild(oNew)
            Set oNew = Nothing
          End If 'Not oNode.nextSibling is nothing
        Next
    End If 'sXpath <> "//head/title"
        
    'trim the textnode
     For Each oNode In oNodes
        oNode.nodeTypedValue = Trim$(oNode.nodeTypedValue)
     Next
        
  End If 'not onodes is nothing
  'Debug.Print oDom.xml
  fncRemoveCrLf = True
ErrHandler:
  If Not fncRemoveCrLf Then objOwner.addlog "<errH in='fncRemoveCrLf'>fncRemoveCrLf ErrH</errH>"
End Function

Public Function fncPrettyInstance( _
  ByRef objDom As IXMLDOMNode, _
  ByVal strIndent As String, _
  ByVal objOwner As oRegenerator _
  ) As Boolean
Dim objChild As IXMLDOMNode
Dim objNew As IXMLDOMNode
Static indentOrigLen As Integer
  
  fncPrettyInstance = False
  If indentOrigLen = 0 Then indentOrigLen = Len(strIndent)
  
  If objDom.childNodes.length > 0 Then
    For Each objChild In objDom.childNodes
      fncPrettyInstance objChild, strIndent & Left$(strIndent, indentOrigLen), _
        objOwner
      If objDom.nodeType = NODE_ELEMENT Then
        Set objNew = objDom.ownerDocument.createNode(NODE_TEXT, vbNullString, vbNullString)
        objNew.nodeValue = vbCrLf & strIndent
        Set objNew = objDom.insertBefore(objNew, objChild)
        Set objNew = Nothing
      End If
    Next
    
    If objDom.nodeType = NODE_ELEMENT Then
      Set objNew = objDom.ownerDocument.createNode(NODE_TEXT, vbNullString, vbNullString)
      objNew.nodeValue = vbCrLf & Left(strIndent, Len(strIndent) - 1)
      Set objNew = objDom.appendChild(objNew)
      Set objNew = Nothing
    End If
  End If
  fncPrettyInstance = True
ErrHandler:
  If Not fncPrettyInstance Then objOwner.addlog "<errH in='fncPrettyInstance'>fncPrettyInstance ErrH" & Err.Description & "</errH>"
End Function

Private Function fncRemoveAnchorCrLf( _
    ByRef oDom As IXMLDOMDocument, _
    ByVal objOwner As oRegenerator _
    ) As Boolean

Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
  On Error GoTo ErrHandler
  fncRemoveAnchorCrLf = False

  'in the ncc, select all anchors
  Set oNodes = oDom.documentElement.selectNodes("//a")
  
  'remove the crlf before the a open
  For Each oNode In oNodes
    If oNode.previousSibling.nodeType = NODE_TEXT Then
      'only del the node if its only whitespace
      If Trim$(oNode.previousSibling.Text) = "" Then
        oNode.previousSibling.nodeValue = vbNullString
      End If
    End If
  Next
  
  'remove the crlf after the a close
  For Each oNode In oNodes
    If oNode.nextSibling.nodeType = NODE_TEXT Then
     'only del the node if its only whitespace
      If Trim$(oNode.nextSibling.Text) = "" Then
        oNode.nextSibling.nodeValue = vbNullString
      End If
    End If
  Next
  
  'trim the anchor textnode
  For Each oNode In oNodes
    oNode.nodeTypedValue = Trim$(oNode.nodeTypedValue)
  Next
  
  'Debug.Print oDom.xml
  fncRemoveAnchorCrLf = True
ErrHandler:
  If Not fncRemoveAnchorCrLf Then objOwner.addlog "<errH in='fncRemoveAnchorCrLf'>fncRemoveAnchorCrLf ErrH</errH>"
End Function
