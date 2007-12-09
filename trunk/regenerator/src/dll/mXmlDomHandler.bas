Attribute VB_Name = "mXmlDomHandler"
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

Public Function fncRenameNodes( _
    ByVal sOldNameNodesAsXpath As String, _
    ByVal sOldNameAsString As String, _
    ByVal sNewName As String, _
    ByRef ioDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oNodes As IXMLDOMNodeList
Dim oFoundNode As IXMLDOMNode
Dim oFoundNodeParent As IXMLDOMNode
Dim oNewNode As IXMLDOMNode
Dim sNodeValue As String
    
    On Error GoTo ErrHandler
    fncRenameNodes = False
    
    Set oNodes = ioDom.selectNodes(sOldNameNodesAsXpath)
    If Not oNodes Is Nothing Then
      If oNodes.length > 0 Then
        For Each oFoundNode In oNodes
          sNodeValue = oFoundNode.nodeTypedValue
          If oFoundNode.nodeType = NODE_ATTRIBUTE Then
            Set oFoundNodeParent = oFoundNode.selectSingleNode("..")
            oFoundNodeParent.Attributes.removeNamedItem (sOldNameAsString)
            If Not fncAppendAttribute(oFoundNodeParent, sNewName, sNodeValue, objOwner) Then GoTo ErrHandler
          ElseIf oFoundNode.nodeType = NODE_ELEMENT Then
            'revisit not done
          End If
        Next
      End If
    End If
    
    fncRenameNodes = True
    
ErrHandler:
  If Not fncRenameNodes Then objOwner.addlog "<errH in='fncRenameNodes' sOldNameNodesAsXpath='" & sOldNameNodesAsXpath & "' sNewName='" & sNewName & "'>fncRenameNodes ErrHandler</errH>"
End Function

Public Function fncRenameElementNode( _
    ByRef oOldElementNode As IXMLDOMNode, _
    ByVal sNewElementNodeName As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oNewElementNode As IXMLDOMElement
Dim oChildNode As IXMLDOMNode
Dim oClonedNode As IXMLDOMNode
Dim oNamedNodeMap As IXMLDOMNamedNodeMap
Dim oAttribute As IXMLDOMAttribute
Dim i As Long
    ' this function will loose the byref oldnode pointer
    ' so dont reuse the same object (oOldElementNode) without setting it again
    On Error GoTo ErrHandler
    fncRenameElementNode = False
      If Not oOldElementNode Is Nothing Then
        Set oNewElementNode = oOldElementNode.ownerDocument.createElement(sNewElementNodeName)
        
        Set oNamedNodeMap = oOldElementNode.Attributes
        For i = 0 To oNamedNodeMap.length - 1
          Set oClonedNode = oNamedNodeMap.Item(i).cloneNode(True)
          Set oClonedNode = oNewElementNode.Attributes.setNamedItem(oClonedNode)
        Next i
        
        For Each oChildNode In oOldElementNode.childNodes
          Set oClonedNode = oChildNode.cloneNode(True)
          Set oClonedNode = oNewElementNode.appendChild(oClonedNode)
        Next
        
        'Debug.Print oOldElementNode.xml
        'Debug.Print oNewElementNode.xml
                
        Set oOldElementNode = oOldElementNode.parentNode.replaceChild(oNewElementNode, oOldElementNode)
                
      End If
    fncRenameElementNode = True
    Exit Function
ErrHandler:
    If Not fncRenameElementNode Then objOwner.addlog "<errH in='fncRenameElementNode'>fncRenameElementNode ErrHandler</errH>"
End Function


Public Function fncRemoveNodes( _
    ByRef oOwnerDoc As MSXML2.DOMDocument40, _
    ByVal sXpath As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oNode As IXMLDOMNode
Dim oNodes As IXMLDOMNodeList

  On Error GoTo ErrHandler
  fncRemoveNodes = False
  
  Set oNodes = oOwnerDoc.selectNodes(sXpath)
  
  If Not (oNodes.length = 0) Or (oNodes Is Nothing) Then
    For Each oNode In oNodes
      Set oNode = oNode.parentNode.removeChild(oNode)
    Next
  End If
  
  fncRemoveNodes = True
  
ErrHandler:
  If Not fncRemoveNodes Then objOwner.addlog "<errH in='fncRemoveNodes'>fncRemoveNodes ErrHandler</errH>"
End Function

Public Function fncAppendAttribute( _
    ByRef oParentElem As IXMLDOMElement, _
    ByVal sAttrName As String, _
    ByVal sAttrValue As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oNewAttr As IXMLDOMNode
Dim oNamedNodeMap As IXMLDOMNamedNodeMap

  On Error GoTo ErrHandler
  fncAppendAttribute = False
  
  If sAttrName <> "" Then
        'Set oNewAttr = oDom.createNode(NODE_ATTRIBUTE, sAttrName, "")
        Set oNewAttr = oParentElem.ownerDocument.createNode(NODE_ATTRIBUTE, sAttrName, "")
        oNewAttr.nodeTypedValue = sAttrValue
        Set oNamedNodeMap = oParentElem.Attributes
        oNamedNodeMap.setNamedItem oNewAttr
  End If
  
  fncAppendAttribute = True
ErrHandler:
  If Not fncAppendAttribute Then objOwner.addlog "<errH in='fncAppendAttribute'>fncAppendAttribute ErrH</errH>"
End Function

Public Function fncRemoveAttribute( _
    ByRef oParentElem As IXMLDOMElement, _
    ByVal sAttrName As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean

Dim oRemoveAttr As IXMLDOMNode
Dim oNamedNodeMap As IXMLDOMNamedNodeMap

  On Error GoTo ErrHandler
  fncRemoveAttribute = False
  
  Set oNamedNodeMap = oParentElem.Attributes
  If Not oNamedNodeMap Is Nothing Then
    Set oRemoveAttr = oNamedNodeMap.removeNamedItem(sAttrName)
  End If
      
  fncRemoveAttribute = True
ErrHandler:
  If Not fncRemoveAttribute Then objOwner.addlog "<errH in='fncRemoveAttribute'>fncRemoveAttribute ErrH</errH>"
End Function

Public Function fncAppendChild( _
    ByRef oParentNode As IXMLDOMElement, _
    ByVal sName As String, _
    ByRef objOwner As oRegenerator, _
    Optional ByVal sTextNode As String, _
    Optional ByVal sAttr1Name As String, _
    Optional ByVal sAttr1Value As String, _
    Optional ByVal sAttr2Name As String, _
    Optional ByVal sAttr2Value As String, _
    Optional ByVal sAttr3Name As String, _
    Optional ByVal sAttr3Value As String, _
    Optional ByVal sAttr4Name As String, _
    Optional ByVal sAttr4Value As String _
    ) As Boolean

Dim oNewNode As IXMLDOMNode
Dim oNewAttr As IXMLDOMNode
Dim oNamedNodeMap As IXMLDOMNamedNodeMap
Dim oDom As New MSXML2.DOMDocument40
    oDom.async = False
    oDom.validateOnParse = False
    oDom.resolveExternals = False
    oDom.preserveWhiteSpace = True
    oDom.setProperty "SelectionLanguage", "XPath"
    oDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
    oDom.setProperty "NewParser", True
    
    On Error GoTo ErrHandler
    fncAppendChild = False

    'Set oNewNode = oDom.createNode(NODE_ELEMENT, sName, "http://www.w3.org/1999/xhtml") '!xht
    Set oNewNode = oDom.createNode(NODE_ELEMENT, sName, "") '!xht
    
    If sTextNode <> "" Then oNewNode.nodeTypedValue = sTextNode
    
    If sAttr1Name <> "" Then
        Set oNewAttr = oDom.createNode(NODE_ATTRIBUTE, sAttr1Name, "")
        oNewAttr.nodeTypedValue = sAttr1Value
        Set oNamedNodeMap = oNewNode.Attributes
        oNamedNodeMap.setNamedItem oNewAttr
    End If
    
    If sAttr2Name <> "" Then
        Set oNewAttr = oDom.createNode(NODE_ATTRIBUTE, sAttr2Name, "")
        oNewAttr.nodeTypedValue = sAttr2Value
        Set oNamedNodeMap = oNewNode.Attributes
        oNamedNodeMap.setNamedItem oNewAttr
    End If
    
    If sAttr3Name <> "" Then
        Set oNewAttr = oDom.createNode(NODE_ATTRIBUTE, sAttr3Name, "")
        oNewAttr.nodeTypedValue = sAttr3Value
        Set oNamedNodeMap = oNewNode.Attributes
        oNamedNodeMap.setNamedItem oNewAttr
    End If
    
    If sAttr4Name <> "" Then
        Set oNewAttr = oDom.createNode(NODE_ATTRIBUTE, sAttr4Name, "")
        oNewAttr.nodeTypedValue = sAttr4Value
        Set oNamedNodeMap = oNewNode.Attributes
        oNamedNodeMap.setNamedItem oNewAttr
    End If
    
    If oParentNode.childNodes.length > 2 Then
      'Set oNewNode = oParentNode.insertBefore(oNewNode, oParentNode.selectSingleNode("/*[last()]"))
      Set oNewNode = oParentNode.insertBefore(oNewNode, oParentNode.lastChild.previousSibling)
    Else
      Set oNewNode = oParentNode.appendChild(oNewNode)
    End If
    fncAppendChild = True
ErrHandler:
   Set oDom = Nothing
   If Not fncAppendChild Then objOwner.addlog "<errH in='fncAppendChild'>fncAppendChild ErrHandler</errH>"
End Function

Public Function fncCountNodes( _
    oParentNode As IXMLDOMNode, _
    sXpath As String _
    ) As Long
Dim oNodes As IXMLDOMNodeList
  On Error GoTo ErrHandler

  Set oNodes = oParentNode.selectNodes(sXpath)
  
  If Not oNodes Is Nothing Then
    fncCountNodes = oNodes.length
  Else
    fncCountNodes = 0
  End If

ErrHandler:

End Function

Public Function fncNodeExists( _
    ByRef oParentNode As IXMLDOMNode, _
    ByVal sXpath As String, _
    ByRef sNodeName As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
' should be faster than using fncCountNodes
' since it doesnt traverse the whole tree
' sNodeName is a byref string that can be used to return the nodename of the node

Dim oNode As IXMLDOMNode
  On Error GoTo ErrHandler
  fncNodeExists = False
  
  Set oNode = oParentNode.selectSingleNode(sXpath)
  
  If Not oNode Is Nothing Then
    fncNodeExists = True
    sNodeName = oNode.nodeName
  Else
   sNodeName = ""
  End If
  Exit Function
ErrHandler:
   objOwner.addlog "<errH in='fncNodeExists'>fncNodeExists ErrHandler</errH>"
End Function

Public Function fncAddId( _
    ByRef oNodeList As IXMLDOMNodeList, _
    ByVal sPrefix As String, _
    ByRef objOwner As oRegenerator, _
    Optional ByVal sFormatType As String, _
    Optional bSkipIfExists As Boolean _
    ) As Boolean
    
Dim oDom As New MSXML2.DOMDocument40
    oDom.async = False
    oDom.validateOnParse = False
    oDom.resolveExternals = False
    oDom.setProperty "NewParser", True
Dim k As Long
Dim oNode As IXMLDOMNode
Dim oAttrNode As IXMLDOMNode
Dim oNamedNodeMap As IXMLDOMNamedNodeMap

  On Error GoTo ErrHandler
  
  fncAddId = False

  k = 0
  For Each oNode In oNodeList
    Set oAttrNode = oNode.selectSingleNode("@id")
    'if (if there is no id) or (id exists and should be replaced)
    If (oAttrNode Is Nothing) Or ((Not oAttrNode Is Nothing) And (Not bSkipIfExists)) Then
        Set oAttrNode = oDom.createNode(NODE_ATTRIBUTE, "id", "")
        'Set oNamedNodeMap = oNodeList.Item(k).Attributes
        Set oNamedNodeMap = oNode.Attributes
        oAttrNode.nodeValue = fncGenerateSequId(sPrefix, k + 1, sFormatType)
        oNamedNodeMap.setNamedItem oAttrNode
        k = k + 1
    End If
  Next

  fncAddId = True
  
ErrHandler:
  Set oDom = Nothing
  If Not fncAddId Then objOwner.addlog ("<errH in='fncAddId'>fncAddId ErrHandler</errH>")
End Function

Public Function fncGenerateSequId( _
    sPrefix As String, _
    i As Long, _
    Optional sFormatType As String _
    ) As String
    
    If sFormatType <> "" Then
        fncGenerateSequId = sPrefix & Format(i, sFormatType)
    Else
        fncGenerateSequId = sPrefix & Format(i, "0000")
    End If
End Function

Public Function fncAddNamesSpaces(objOwner As oRegenerator) As Boolean
Dim sDomData As String, i As Long
  
  On Error GoTo ErrHandler
  fncAddNamesSpaces = False
  
  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_NCC Or _
      objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_CONTENT Then
         sDomData = objOwner.objFileSetHandler.aOutFileSet(i).sDomData
         If Not fncAddNameSpace(sDomData, objOwner) Then GoTo ErrHandler
         objOwner.objFileSetHandler.aOutFileSet(i).sDomData = sDomData
         sDomData = ""
    End If
  Next i
  
  fncAddNamesSpaces = True
  
ErrHandler:
  If Not fncAddNamesSpaces Then objOwner.addlog ("<errH in='fncAddNamesSpaces'>fncAddNamesSpaces ErrHandler</errH>")
End Function
  
Private Function fncAddNameSpace( _
    ByRef sStringToHack As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
    On Error GoTo ErrHandler
    fncAddNameSpace = False
    sStringToHack = Replace$(sStringToHack, "<html", "<html xmlns=""http://www.w3.org/1999/xhtml""")
    fncAddNameSpace = True
ErrHandler:
  If Not fncAddNameSpace Then objOwner.addlog "<errH in='fncAddNameSpace'>fncAddNameSpace ErrH</errH>"
End Function

Public Function fncMoveSiblingToTop( _
    oOwnerDom As MSXML2.DOMDocument40, _
    sXpath As String, _
    objOwner As oRegenerator _
    ) As Boolean
Dim oSibling As IXMLDOMNode

 'this function moves a certain sibling to the first pos in sibling set
 fncMoveSiblingToTop = False
 On Error GoTo ErrHandler
 
 Set oSibling = oOwnerDom.selectSingleNode(sXpath)
 If Not oSibling Is Nothing Then
   Set oSibling = oSibling.parentNode.insertBefore(oSibling, oSibling.parentNode.firstChild)
 End If 'oSibling Is Nothing
 fncMoveSiblingToTop = True
 
ErrHandler:
  If Not fncMoveSiblingToTop Then objOwner.addlog "<errH in='fncMoveSiblingToTop'>fncMoveSiblingToTop ErrH</errH>"
End Function

Public Function fncEscapeQuotes( _
  ByVal sStringToEscape As String _
  ) As String
  'mg20030807, delete them for now, amp gets reescaped later
  'this is a hack of immense proportion
  'but handles weirdness (&x0a;     )from tidyLib?
  sStringToEscape = Replace(Replace(Replace(Replace(sStringToEscape, Chr(39), ""), Chr(34), ""), Chr(10), ""), "      ", " ")
  fncEscapeQuotes = sStringToEscape
  'Debug.Print sStringToEscape
End Function

