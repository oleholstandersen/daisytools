VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFilter"
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

Private oFilterDom As New msxml2.DOMDocument40

Public sName As String
Public sFullPath As String
Public sDefaultNameSpace As String
Public sDefaultNameSpaceUri As String
Public sLocalDtdPath As String
Public sOutFileSet As String

Public sXpathStartElem As String
Public sXpathRemove As String
Public sXpathSyncOmit As String
Public sXpathSyncForce As String
Public sXpathSmilRefOmit As String
Public sXpathSmilSequence As String
Public sXpathSmilSplit As String
Public sXpathNavInclude As String
Public sXpathNavIncludeNoDeep As String
Public sXpathMixedContentHandler As String

Public bMixedContentHandlerActive As Boolean
Public sItmtAttValue As String
Public sItmtAttName As String
Public sItmtElemName As String

Private Sub Class_Initialize()
  'debug.Print "cFilter.initialize"
  oFilterDom.async = False
  oFilterDom.validateOnParse = True
  oFilterDom.resolveExternals = True
  oFilterDom.preserveWhiteSpace = False
  oFilterDom.setProperty "SelectionLanguage", "XPath"
  oFilterDom.setProperty "SelectionNamespaces", "xmlns:bruno='http://www.daisy.org/2004/bruno'"
  oFilterDom.setProperty "NewParser", False
End Sub

Private Sub Class_Terminate()
  'debug.Print "cFilter.terminate"
  Set oFilterDom = Nothing
End Sub

Public Function fncInstantiate(sFilterFullPath As String) As Boolean
Dim oItemNode As IXMLDOMNode
Dim oItemNodes As IXMLDOMNodeList
    
  fncInstantiate = False
   
  If Not oBruno.oCmn.oDomCmn.fncParseFile(sFilterFullPath, oFilterDom, "") Then GoTo errH
  
  Me.sFullPath = sFilterFullPath
  Me.sName = oFilterDom.selectSingleNode("//bruno:name").Text
    
  sOutFileSet = oFilterDom.selectSingleNode("//bruno:bruno-filter/@outfileset").Text
  
  sDefaultNameSpace = oFilterDom.selectSingleNode("//bruno:bruno-filter/@target-namespace").Text
  sDefaultNameSpaceUri = oBruno.oCmn.oUriParser.fncGetAttrValueFromString(sDefaultNameSpace, Mid(sDefaultNameSpace, 1, InStr(1, sDefaultNameSpace, "=")))
  
  sLocalDtdPath = oFilterDom.selectSingleNode("//bruno:bruno-filter/@local-dtd").Text
  If Not oBruno.oCmn.oFsoCmn.fncFileExists(sLocalDtdPath) Then
    'the user didnt provide a full path to the dtd, fallback to Bruno default DTD dir:
    sLocalDtdPath = oBruno.oPaths.DtdPath & sLocalDtdPath
  End If
  
  Set oItemNodes = oFilterDom.selectNodes("//bruno:start/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathStartElem) Then GoTo errH
  
  Set oItemNodes = oFilterDom.selectNodes("//bruno:remove/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathRemove) Then GoTo errH
    
  Set oItemNodes = oFilterDom.selectNodes("//bruno:sync-omit/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSyncOmit) Then GoTo errH
    
  Set oItemNodes = oFilterDom.selectNodes("//bruno:sync-force/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSyncForce) Then GoTo errH
    
  Set oItemNodes = oFilterDom.selectNodes("//bruno:smilref-omit/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSmilRefOmit) Then GoTo errH
    
  Set oItemNodes = oFilterDom.selectNodes("//bruno:smil-sequence/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSmilSequence) Then GoTo errH
    
  Set oItemNodes = oFilterDom.selectNodes("//bruno:smil-split/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSmilSplit) Then GoTo errH
    
  Set oItemNodes = oFilterDom.selectNodes("//bruno:nav-include/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathNavInclude) Then GoTo errH
  
  Set oItemNodes = oFilterDom.selectNodes("//bruno:nav-include/bruno:item[@deep='false']")
  If Not fncBuildXpath(oItemNodes, sXpathNavIncludeNoDeep) Then GoTo errH
      
  Set oItemNodes = oFilterDom.selectNodes("//bruno:mixed-content-handler/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathMixedContentHandler) Then GoTo errH
            
  Set oItemNode = oFilterDom.selectSingleNode("//bruno:mixed-content-handler[@active='true']")
  If Not oItemNode Is Nothing Then
    bMixedContentHandlerActive = True
  End If
  
  Set oItemNode = oFilterDom.selectSingleNode("//bruno:mixed-content-handler/@elemName")
  If Not oItemNode Is Nothing Then
    sItmtElemName = oItemNode.Text
  End If
        
  Set oItemNode = oFilterDom.selectSingleNode("//bruno:mixed-content-handler/@attName")
  If Not oItemNode Is Nothing Then
    sItmtAttName = oItemNode.Text
  End If
        
  Set oItemNode = oFilterDom.selectSingleNode("//bruno:mixed-content-handler/@attValue")
  If Not oItemNode Is Nothing Then
    sItmtAttValue = oItemNode.Text
  End If
    
  fncInstantiate = True
errH:

 Set oItemNode = Nothing
 Set oItemNodes = Nothing
End Function

Private Function fncBuildXpath(ByRef oItemNodes As IXMLDOMNodeList, ByRef sXpathString As String) As Boolean
Dim i As Long
  fncBuildXpath = False
  
  sXpathString = ""
  If Not oItemNodes Is Nothing Then
    For i = 0 To oItemNodes.length - 2
      sXpathString = sXpathString & oItemNodes(i).selectSingleNode("@xpath").Text & " | "
    Next
    'add last iterat without ending pipe
    If Not oItemNodes.length = 0 Then sXpathString = sXpathString & oItemNodes(i).selectSingleNode("@xpath").Text
  End If
  If sXpathString = "" Then sXpathString = "/nothing"
  
  fncBuildXpath = True
End Function
