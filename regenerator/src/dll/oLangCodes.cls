VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oLangCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private oPrivLangCodeDom As New MSXML2.DOMDocument40
Private oPrivLangCodes As IXMLDOMNodeList

Private Sub Class_Initialize()
    oPrivLangCodeDom.async = False
    oPrivLangCodeDom.validateOnParse = False
    oPrivLangCodeDom.resolveExternals = False
    oPrivLangCodeDom.preserveWhiteSpace = False
    oPrivLangCodeDom.setProperty "SelectionLanguage", "XPath"
    oPrivLangCodeDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
    oPrivLangCodeDom.setProperty "NewParser", True
End Sub

Public Function fncInitialize(objOwner As oRegenerator) As Boolean
    fncInitialize = False
    If Not fncParseFile(sResourcePath & "iso639.xml", oPrivLangCodeDom, objOwner) Then GoTo ErrHandler
    Set oPrivLangCodes = oPrivLangCodeDom.selectNodes("//code")
    If oPrivLangCodes Is Nothing Then GoTo ErrHandler
    fncInitialize = True
ErrHandler:
End Function

Public Function fncTerminateObject() As Boolean
Dim bolResult As Boolean
  Set oPrivLangCodes = Nothing
  Set oPrivLangCodeDom = Nothing
  If Not ( _
       oPrivLangCodes Is Nothing Or _
       oPrivLangCodeDom Is Nothing _
        ) Then bolResult = False Else bolResult = True
  fncTerminateObject = bolResult
End Function

Public Property Get oLangCodeDom() As MSXML2.DOMDocument40
  Set oLangCodeDom = oPrivLangCodeDom
End Property

Public Property Get oLangCodes() As IXMLDOMNodeList
  Set oLangCodes = oPrivLangCodes
End Property

Private Sub Class_Terminate()
  Set oPrivLangCodeDom = Nothing
  Set oPrivLangCodes = Nothing
End Sub

