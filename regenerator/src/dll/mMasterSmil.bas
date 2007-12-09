Attribute VB_Name = "mMasterSmil"
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


Option Explicit

Public Function fncGenerateMasterSmil( _
    sOutCharsetName As String, _
    objOwner As oRegenerator _
    ) As Boolean

Dim oMasterSmilDom As New MSXML2.DOMDocument40
    oMasterSmilDom.async = False
    oMasterSmilDom.validateOnParse = False
    oMasterSmilDom.resolveExternals = False
    oMasterSmilDom.preserveWhiteSpace = False
    oMasterSmilDom.setProperty "SelectionLanguage", "XPath"
    oMasterSmilDom.setProperty "NewParser", True
Dim strSkeleton As String
Dim oNode As IXMLDOMNode, dNode As IXMLDOMNode
Dim oNodeList As IXMLDOMNodeList
Dim oAttrs As IXMLDOMNamedNodeMap
'Dim oAttrNode As IXMLDOMAttribute
'Dim oNodeValue As IXMLDOMAttribute
Dim oAttr As IXMLDOMAttribute
Dim i As Long

  On Error GoTo ErrHandler
  fncGenerateMasterSmil = False

  strSkeleton = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " " & "encoding=" & Chr(34) & sOutCharsetName & Chr(34) & "?>" & _
       "<!DOCTYPE smil PUBLIC " & Chr(34) & "-//W3C//DTD smil 1.0//EN" & Chr(34) & " " & Chr(34) & "http://www.w3.org/TR/REC-smil/smil10.dtd" & Chr(34) & ">" & _
       "<smil>" & _
       "<head>" & _
       "<layout>" & _
       "<region id=" & Chr(34) & "txtView" & Chr(34) & " />" & _
       "</layout>" & _
       "</head>" & _
       "<body>" & _
       "</body>" & _
       "</smil>"

  'initiate mastersmil as oMasterSmilDoc
  If Not fncParseString(strSkeleton, oMasterSmilDom, objOwner) Then GoTo ErrHandler
  
  If Not fncDoNonNccBiblioMetaData( _
        oMasterSmilDom, _
        TYPE_SMIL_MASTER, _
        False, _
        objOwner _
        ) Then GoTo ErrHandler

  'add ncc:timeInThissmil
  If Not fncAppendChild(oMasterSmilDom.selectSingleNode("//head"), "meta", objOwner, , "name", "ncc:timeInThisSmil", "content", fncConvertMS2SmilClockVal(objOwner.lTotalTimeMs, 0)) Then GoTo ErrHandler
  
  'create ref children
    For i = 1 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
      If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1 Then
        'create a new ref child
        Set oNode = oMasterSmilDom.createNode(NODE_ELEMENT, "ref", "")
        'get localname into src attr
        Set oAttr = oMasterSmilDom.createAttribute("src")
        oAttr.nodeValue = objOwner.objFileSetHandler.aOutFileSet(i).sFileName
        Set oAttrs = oNode.Attributes
        oAttrs.setNamedItem oAttr
        'get title into title attr
        Set oAttr = oMasterSmilDom.createAttribute("title")
        oAttr.nodeValue = objOwner.objFileSetHandler.aOutFileSet(i).sSmilTitle
        Set oAttrs = oNode.Attributes
        oAttrs.setNamedItem oAttr
        ' append to document
        Set oNode = oMasterSmilDom.documentElement.lastChild.appendChild(oNode)
      End If
    Next i
    
    'now omastersmil contains all ref children, add id to each
    Set oNodeList = oMasterSmilDom.selectNodes("//ref")

    If Not fncAddId(oNodeList, "rgn_smil_", objOwner, "0000", False) Then Exit Function
    
        
    If Not fncRemoveSchemeAttrs(oMasterSmilDom, objOwner) Then GoTo ErrHandler
        
    objOwner.objFileSetHandler.fncAddObjectToOutputArray TYPE_SMIL_MASTER, "master.smil", oMasterSmilDom.xml
   
   fncGenerateMasterSmil = True
ErrHandler:
    Set oMasterSmilDom = Nothing
    DoEvents
    If Not fncGenerateMasterSmil Then objOwner.addlog "<errH in='fncGeneratemastersmil'>fncGeneratemastersmil ErrHandler</errH>"
End Function
