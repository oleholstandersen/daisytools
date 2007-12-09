Attribute VB_Name = "mMetaHandler"
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

Public Function fncDoNonNccBiblioMetaData( _
    ByRef oMemberDom As MSXML2.DOMDocument40, _
    ByRef lFileType As Long, _
    ByRef bolPreserveBiblioMeta As Boolean, _
    ByRef objOwner As oRegenerator, _
    Optional ByRef lCurrentOutArrayItem As Long _
    ) As Boolean
Dim i As Long
Dim oTitleNode As IXMLDOMNode
Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim oNodeAtt As IXMLDOMNode
  
  '***********************************************************
  '* this is a generisized function for removing/reading common
  '* bibliographic metadata to contentdocs and smilfiles
  '* it uses the cCommonmeta class which is instantiated
  '* after ncc meta has been done
  '***********************************************************
  
  On Error GoTo ErrHandler
  fncDoNonNccBiblioMetaData = False
    
  '**********************************************************
  '* remove preexisting elements
  '**********************************************************
  
  Set oNodes = oMemberDom.selectNodes("//head/meta") '!xht
  
  If Not (oNodes.length = 0) Or (oNodes Is Nothing) Then
    For Each oNode In oNodes
      If bolPreserveBiblioMeta And lFileType = TYPE_SMIL_CONTENT Then
        Set oNodeAtt = oNode.selectSingleNode("@name")
        If Not oNodeAtt Is Nothing Then
          Dim sTmp As String: sTmp = LCase$(Trim(oNodeAtt.Text))
          If sTmp = "dc:identifier" Or _
             sTmp = "dc:title" Or _
             sTmp = "dc:creator" Or _
             sTmp = "dc:format" Or _
             sTmp = "ncc:generator" Then
            Set oNode = oNode.parentNode.removeChild(oNode)
            
          End If 'sTmp=
        End If 'Not oNodeAtt Is Nothing Then
        
        Set oNodeAtt = oNode.selectSingleNode("@http-equiv")
        If Not oNodeAtt Is Nothing Then
          sTmp = LCase$(Trim(oNodeAtt.Text))
          If sTmp = "content-type" Then
            Set oNode = oNode.parentNode.removeChild(oNode)
          End If
        End If
        
      ElseIf (Not bolPreserveBiblioMeta And lFileType = TYPE_SMIL_CONTENT) _
        Or (lFileType = TYPE_SMIL_1) Then
          Set oNode = oNode.parentNode.removeChild(oNode)
          
      End If 'if bWantsPreserveMeta
    Next
  End If 'Not (oNodes.length = 0) Or (oNodes Is Nothing)
      
  If (lFileType = TYPE_SMIL_CONTENT) And (Not bolPreserveBiblioMeta) Then
    Set oTitleNode = oMemberDom.selectSingleNode("//head/title") '!xht
  ElseIf lFileType = TYPE_SMIL_1 Then
    Set oTitleNode = oMemberDom.selectSingleNode("//head/title")
  End If
  If Not oTitleNode Is Nothing Then Set oTitleNode = oTitleNode.parentNode.removeChild(oTitleNode)
  
  Set oNode = Nothing
      
  '*******************************************'
  '* then append the nodes from cCommonMeta
  '*******************************************'
  
  Dim oClone As IXMLDOMNode
  Dim oMemberHeadNode As IXMLDOMNode
  
  Set oMemberHeadNode = oMemberDom.selectSingleNode("//head") '!xht this xpath is the same for all filetypes
  
  'this function is called from smil, content and mastersmil.
  'start with adding those that go into all these types:
  
  Set oClone = objOwner.objCommonMeta.NccGenerator.cloneNode(True)
  If Not oClone Is Nothing Then Set oClone = oMemberHeadNode.insertBefore(oClone, oMemberHeadNode.firstChild)
  
  
  Set oClone = objOwner.objCommonMeta.DcFormat.cloneNode(True)
  If Not oClone Is Nothing Then Set oClone = oMemberHeadNode.insertBefore(oClone, oMemberHeadNode.firstChild)
  
  If Not objOwner.objCommonMeta.DcTitle Is Nothing Then
    Set oClone = objOwner.objCommonMeta.DcTitle.cloneNode(True)
    '20041005 fix for victor prob:
    If lFileType = TYPE_SMIL_1 Then
      Dim oTempClone As IXMLDOMNode
      Set oTempClone = oClone.selectSingleNode("@content")
      If Not oTempClone Is Nothing Then
        If InStr(1, oTempClone.Text, "id", vbTextCompare) > 0 Then
          oTempClone.Text = oTempClone.Text & " "
        End If
      End If
    End If
    Set oClone = oMemberHeadNode.insertBefore(oClone, oMemberHeadNode.firstChild)
  Else
    'objowner.addlog "warning: no book title available"
  End If
  
  If Not objOwner.objCommonMeta.DcIdentifier Is Nothing Then
    Set oClone = objOwner.objCommonMeta.DcIdentifier.cloneNode(True)
    Set oClone = oMemberHeadNode.insertBefore(oClone, oMemberHeadNode.firstChild)
  Else
    'objowner.addlog "warning: no book identifier available"
  End If
  
  'then add specifics:
          
  If lFileType = TYPE_SMIL_CONTENT Then
    If Not objOwner.objCommonMeta.DcCreator Is Nothing Then
      Set oClone = objOwner.objCommonMeta.DcCreator.cloneNode(True)
      Set oClone = oMemberHeadNode.insertBefore(oClone, oMemberHeadNode.firstChild)
    Else
      'objowner.addlog "warning: no book creator available"
    End If
    Set oClone = objOwner.objCommonMeta.HttpEquiv.cloneNode(True)
    Set oClone = oMemberHeadNode.insertBefore(oClone, oMemberHeadNode.firstChild)
  End If
  
  If lFileType = TYPE_SMIL_1 Then
    Dim sTempTitle As String
    sTempTitle = objOwner.objFileSetHandler.aOutFileSet(lCurrentOutArrayItem).sSmilTitle
    'mg20041005: fix for victor prob when "id" in smil title: append space to end of value
    If InStr(1, sTempTitle, "id", vbTextCompare) > 0 Then
      sTempTitle = sTempTitle & " "
    End If
    If Not fncAppendChild(oMemberHeadNode, "meta", objOwner, , "name", "title", "content", _
    sTempTitle) Then
      'objowner.addlog "warning: smil title elem set fail in fncDoNonNccBiblioMetaData"
    End If
  End If
    
  If (lFileType = TYPE_SMIL_CONTENT) And (Not bolPreserveBiblioMeta) Then
    If Not objOwner.objCommonMeta.DcTitle Is Nothing Then
      If Not fncAppendChild(oMemberHeadNode, "title", objOwner, objOwner.objCommonMeta.DcTitle.selectSingleNode("@content").Text) Then
        'objowner.addlog "warning: content doc title elem set fail in fncDoNonNccBiblioMetaData"
      End If
    End If
  End If
    
  'note that timeinthissmil and totalelapsed time etc are still missing in the files,
  'but these are not added here since they are not static fields; not necessarily computed yet
    
  fncDoNonNccBiblioMetaData = True
  
ErrHandler:
  If Not fncDoNonNccBiblioMetaData Then objOwner.addlog "<errH in='fncDoNonNccBiblioMetaData'>fncDoNonNccBiblioMetaData ErrHandler</errH>"
End Function
