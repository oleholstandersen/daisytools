VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oXhtmlEntities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private oPrivEntityDom As New MSXML2.DOMDocument40
Private oPrivAllEntities As IXMLDOMNodeList

Private Sub Class_Initialize()
    oPrivEntityDom.async = False
    oPrivEntityDom.validateOnParse = False
    oPrivEntityDom.resolveExternals = False
    oPrivEntityDom.preserveWhiteSpace = False
    oPrivEntityDom.setProperty "SelectionLanguage", "XPath"
    oPrivEntityDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
    oPrivEntityDom.setProperty "NewParser", True
End Sub

Public Function fncInitialize(objOwner As oRegenerator) As Boolean
    fncInitialize = False
    If Not fncParseFile(sResourcePath & "xhtml_entities.xml", oPrivEntityDom, objOwner) Then GoTo ErrH
    Set oPrivAllEntities = oPrivEntityDom.selectNodes("//@*")
    If oPrivAllEntities Is Nothing Then GoTo ErrH
    fncInitialize = True
ErrH:
End Function

Public Function fncTerminateObject() As Boolean
Dim bolResult As Boolean
  Set oPrivAllEntities = Nothing
  Set oPrivEntityDom = Nothing
  If Not ( _
       oPrivAllEntities Is Nothing Or _
       oPrivEntityDom Is Nothing _
        ) Then bolResult = False Else bolResult = True
  fncTerminateObject = bolResult
End Function

Public Property Get oEntityDom() As MSXML2.DOMDocument40
  Set oEntityDom = oPrivEntityDom
End Property

Public Property Get oAllEntities() As IXMLDOMNodeList
  Set oAllEntities = oPrivAllEntities
End Property

Private Sub Class_Terminate()
  Set oPrivEntityDom = Nothing
  Set oPrivAllEntities = Nothing
End Sub
