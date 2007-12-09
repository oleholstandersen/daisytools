VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oDocumentSmilRelations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Daisy 2.02 Validator Engine
' Copyright (C) 2002 Daisy Consortium
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
'
' For information about the Daisy Consortium, visit www.daisy.org or contact
' info@mail.daisy.org. For development issues, contact markus.gylling@tpb.se or
' karl.ekdahl@tpb.se.

' This function checks the linkback to SMIL files from a content document file.
' The procedure is, take all <anchor> elements, if they're pointing to a SMIL file,
' check if that smil file exists and check if the ID (if any) included in the
' reference exists in the SMIL file.
'
Public Function fncCheckLinkback( _
  iobjReport As oReport, isDocFile As String, Optional ibolMultiVolume As Variant _
  ) As Boolean
  
  Dim objFileIntegrity As New oFileIntegrity, objXmlIntegrity As New oXmlIntegrity
' ke20030527
  Dim objBogusReport As oReport, objDocDom As Object
  Dim objSmilDom As Object, objNodeList As Object
  Dim objNode As Object, objNode2 As Object
  Dim sCurrentSmil As String, sLastSmil As String

  'fncInsertTime "oDocumentSmilRelations.fncCheckLinkback"

  Set objSmilDom = CreateObject("Msxml2.DOMDocument.4.0")
  Set objDocDom = CreateObject("Msxml2.DOMDocument.4.0")
' ke20030527
  Set objBogusReport = New oReport
  
  Dim bolMultiVolume As Boolean
  bolMultiVolume = False
  If Not IsMissing(ibolMultiVolume) Then bolMultiVolume = ibolMultiVolume
  
  objXmlIntegrity.fncIsWellformedXML objBogusReport, isDocFile, objDocDom
  If objDocDom Is Nothing Then Exit Function
  
  Dim objNode3 As Object
  
  Set objNodeList = objDocDom.selectNodes("//a") '/@href")
  For Each objNode In objNodeList
  
    Set objNode3 = objNode.selectSingleNode("@href")
    If objNode3 Is Nothing Then GoTo Skip
    
    sCurrentSmil = fncStripIdAddPath(objNode3.nodeValue, isDocFile) 'node3
    If sCurrentSmil = "" Or (Not (LCase$(Right$(sCurrentSmil, 3)) = "sml" Or LCase$(Right$(sCurrentSmil, 4)) = "smil")) Then GoTo Skip
    
    If Not sCurrentSmil = sLastSmil Then
      sLastSmil = sCurrentSmil
      Set objSmilDom = Nothing
' ke20030527
      Set objBogusReport = Nothing
      Set objBogusReport = New oReport
' If the refered file exist but is invalid, the 'objBogusReport' will contain
' at least one error item.
      objXmlIntegrity.fncIsWellformedXML objBogusReport, _
        sCurrentSmil, objSmilDom
    End If
    
    'If ((objSmilDom Is Nothing) Or (objSmilDom.url = "")) Then
    'REVISIT mg20030326 orig above
    'If (objSmilDom Is Nothing) Or (Not objSmilDom Is Nothing And objSmilDom.url = "") Then
    If (objSmilDom Is Nothing) Then
        
' ke20030527
' If 'objBogusReport' contains more than 0 items, the file is invalid but exist;
' don't report it as missing.
        If (Not bolMultiVolume) And (objBogusReport.lFailedTestCount < 1) Then _
          fncInsFail2Report iobjReport, objNode, "documentSmilRelations.fileExist", _
          isDocFile, "file doesn't exist: " & sCurrentSmil
    Else
      iobjReport.subInsertSucceededTest
      
      Set objNode2 = objSmilDom.selectSingleNode("//*[@id='" & _
        fncGetId(objNode3.nodeValue) & "']")
      If objNode2 Is Nothing Then
        fncInsFail2Report iobjReport, objNode, "documentSmilRelations.fragmentExist", _
          isDocFile, "id doesn't exist: " & fncGetId(objNode3.nodeValue)
      Else
        iobjReport.subInsertSucceededTest
      End If
    End If

Skip:
    DoEvents '**BugTrace
  Next objNode
  
  Set objDocDom = Nothing
  Set objSmilDom = Nothing
  
  fncCheckReferences iobjReport, isDocFile
  
  'fncInsertTime "oDocumentSmilRelations.fncCheckLinkback"
End Function

' This function checks that refered non-smil files exists
Private Function fncCheckReferences(iobjReport As oReport, isDocFile As String) _
  As Boolean
  
  Dim objFileIntegrity As New oFileIntegrity, objXmlIntegrity As New oXmlIntegrity
  Dim sCurrentFile As String
  Dim objBogusReport As New oReport, objDocDom As Object
  Dim objNodeList As Object
  
  Set objDocDom = CreateObject("Msxml2.DOMDocument.4.0")
  
  Dim objNode As Object
  
  'fncInsertTime "oDocumentSmilRelations.fncCheckReferences"
  
  objXmlIntegrity.fncIsWellformedXML objBogusReport, isDocFile, objDocDom
  If objDocDom Is Nothing Then Exit Function
  
  Dim objNode3 As Object
  
  Set objNodeList = objDocDom.selectNodes("//link[@rel='stylesheet']" + _
    "| //link[@type='text/css'] | //img")
  For Each objNode In objNodeList
    
    If objNode.nodeName = "link" Then
      Set objNode3 = objNode.selectSingleNode("@href")
    ElseIf objNode.nodeName = "img" Then
      Set objNode3 = objNode.selectSingleNode("@src")
    End If
    If objNode3 Is Nothing Then GoTo Skip
    
    sCurrentFile = fncStripIdAddPath(objNode3.nodeValue, isDocFile)
    'mg20030316: disabled the above
    'ke20030604: reenabled
    'this wont work when nodevalue is "images/ping.jpg"
    'REVISIT: support for relative paths in nodevalues should be added to fncParseURI)
    
    'sCurrentFile = fncGetPathName(isDocFile) & objNode3.nodeValue
    'sCurrentFile = fncGetAbsolutePathName(sCurrentFile)
    
    If sCurrentFile = "" Then GoTo Skip
    
    Set objBogusReport = New oReport
    objFileIntegrity.fncFileExists objBogusReport, sCurrentFile, isDocFile, objNode
    
    If objBogusReport.lFailedTestCount > 0 Then
      fncInsFail2Report iobjReport, objNode, "documentSmilRelations.fileExist", _
        isDocFile, "file doesn't exist: " & sCurrentFile
    End If
Skip:
  Next objNode
  
  'fncInsertTime "oDocumentSmilRelations.fncCheckReferences"
End Function
