VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oCommonMeta"
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


Option Explicit

Private oDcIdentifierNode As IXMLDOMNode
Private oDcCreatorNode As IXMLDOMNode
Private oDcTitleNode As IXMLDOMNode
Private oDcFormatNode As IXMLDOMNode
Private oDcLanguageNode As IXMLDOMNode
Private oNccGeneratorNode As IXMLDOMNode
Private oHttpEquivNode As IXMLDOMNode

'************************************************
'* this class is instantiated after ncc meta creation has been done
'* smil/content/mastersmil use its properties to clone commonly used meta nodes
'************************************************
Public Function fncSetCommonMetaNodes( _
    oSourceDom As MSXML2.DOMDocument40, _
    objOwner As oRegenerator _
    ) As Boolean
  On Error GoTo ErrHandler
  fncSetCommonMetaNodes = False
  Set oDcIdentifierNode = oSourceDom.selectSingleNode("//head/meta[@name='dc:identifier']") '!xht
  Set oDcCreatorNode = oSourceDom.selectSingleNode("//head/meta[@name='dc:creator']") '!xht
  Set oDcTitleNode = oSourceDom.selectSingleNode("//head/meta[@name='dc:title']") '!xht
  Set oDcFormatNode = oSourceDom.selectSingleNode("//head/meta[@name='dc:format']") '!xht
  Set oDcLanguageNode = oSourceDom.selectSingleNode("//head/meta[@name='dc:language']") '!xht
  Set oNccGeneratorNode = oSourceDom.selectSingleNode("//head/meta[@name='ncc:generator']") '!xht
  Set oHttpEquivNode = oSourceDom.selectSingleNode("//head/meta[@http-equiv='Content-type']") '!xht
  
  objOwner.addlog "<identifiers>"
  If Not oDcIdentifierNode Is Nothing Then objOwner.addlog oDcIdentifierNode.xml
  If Not oDcTitleNode Is Nothing Then objOwner.addlog oDcTitleNode.xml
  objOwner.addlog "</identifiers>"
  
  fncSetCommonMetaNodes = True
ErrHandler:
  If Not fncSetCommonMetaNodes Then objOwner.addlog "<errH in='fncSetCommonMetaNodes'>fncSetCommonMetaNodes ErrH</errH>"
End Function

Public Property Get DcIdentifier() As IXMLDOMNode
  If Not oDcIdentifierNode Is Nothing Then Set DcIdentifier = oDcIdentifierNode
End Property

Public Property Get DcCreator() As IXMLDOMNode
  If Not oDcCreatorNode Is Nothing Then Set DcCreator = oDcCreatorNode
End Property

Public Property Get DcTitle() As IXMLDOMNode
  If Not oDcTitleNode Is Nothing Then Set DcTitle = oDcTitleNode
End Property

Public Property Get DcFormat() As IXMLDOMNode
  If Not oDcFormatNode Is Nothing Then Set DcFormat = oDcFormatNode
End Property

Public Property Get DcLanguage() As IXMLDOMNode
  If Not oDcLanguageNode Is Nothing Then Set DcLanguage = oDcLanguageNode
End Property

Public Property Get NccGenerator() As IXMLDOMNode
  If Not oNccGeneratorNode Is Nothing Then Set NccGenerator = oNccGeneratorNode
End Property

Public Property Get HttpEquiv() As IXMLDOMNode
  If Not oHttpEquivNode Is Nothing Then Set HttpEquiv = oHttpEquivNode
End Property

Public Function fncTerminateObject() As Boolean
  Dim bolResult As Boolean
  
  Set oDcIdentifierNode = Nothing
  Set oDcCreatorNode = Nothing
  Set oDcTitleNode = Nothing
  Set oDcFormatNode = Nothing
  Set oDcLanguageNode = Nothing
  Set oNccGeneratorNode = Nothing
  Set oHttpEquivNode = Nothing
  
  If Not ( _
       oDcIdentifierNode Is Nothing Or _
       oDcCreatorNode Is Nothing Or _
       oDcTitleNode Is Nothing Or _
       oDcFormatNode Is Nothing Or _
       oDcLanguageNode Is Nothing Or _
       oNccGeneratorNode Is Nothing Or _
       oHttpEquivNode Is Nothing) Then bolResult = False Else bolResult = True

  fncTerminateObject = bolResult
End Function

