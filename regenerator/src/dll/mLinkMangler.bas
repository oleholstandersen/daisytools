Attribute VB_Name = "mLinkMangler"
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

Private TextSrcArray() As eTextSrcData
Private lTextSrcArrayItems As Long

Private sNoteBodyUriArray() As eNoteBodyData
Private lNoteBodyUriArrayItems As Long


'Public bolRebuildLinkStructure As Boolean

'mLinkMangler
'logic references to smilfiles:
'assumes that base id is id of text element inside par

'For each smilfile
'  For each text@id
'    Give the text id a new value
'      traverse ncc and modify any references to old value
'      (if existing) traverse contentdocs and modify any references to old value
'      if neither ncc nor contentdoc referenced this target, issue warning
'  next 'for each text@id
'next 'for each smilfile

'now there may be ncc and contentdoc smilrefs
'that point to void, these are easily isolated because of naming conventions
'(all regenerator ids begin with "rgn_")

'logic references to text media object:

'For each smilfile
'  for each text@src
'    if uri resolves ok
'      modify target id value
'      give same value to text src fragment part
'      check if this was pointing to ncc document
'    else
'      issue warning
'    end If
'  next 'for each text@src
'Next 'for each smilfile

'If no text/src where pointing to ncc
'rename ids in ncc

'now it is possible to find smilrefs that point to void and make a reasonable guess on position
'if ncc smilref x points to void
'find first previous ncc smilref y that does not point to void
'if points to same smilfile, repoint x to point to next par after y
'etc

Public Function fncRebuildLinkStructure(objOwner As oRegenerator, bolEstimateBrokenXhtmlLinks As Boolean) As Boolean
  On Error GoTo ErrHandler
  objOwner.addlog "<status>rebuilding link structure...</status>"
  
  fncRebuildLinkStructure = False
  
  lTextSrcArrayItems = 0
  lNoteBodyUriArrayItems = 0
  
  If Not fncLinkFixSmilRef(objOwner) Then GoTo ErrHandler
  If Not fncLinkFixTextRef(objOwner, bolEstimateBrokenXhtmlLinks) Then GoTo ErrHandler
  
  'kill the private arrays
  lTextSrcArrayItems = 0: ReDim Preserve TextSrcArray(lTextSrcArrayItems)
  lNoteBodyUriArrayItems = 0: ReDim Preserve TextSrcArray(lNoteBodyUriArrayItems)
    
  fncRebuildLinkStructure = True

ErrHandler:
  If Not fncRebuildLinkStructure Then objOwner.addlog "<errH in='fncRebuildLinkStructure'>fncRebuildLinkStructure ErrH</errH>"
End Function

Private Function fncLinkFixSmilRef(objOwner As oRegenerator) As Boolean
Dim i As Long, k As Long
Dim lSmilFileCount As Long
Dim oSmilDom As New MSXML2.DOMDocument40
    oSmilDom.async = False
    oSmilDom.validateOnParse = False
    oSmilDom.resolveExternals = False
    oSmilDom.preserveWhiteSpace = False
    oSmilDom.setProperty "SelectionLanguage", "XPath"
    oSmilDom.setProperty "NewParser", True
Dim oNccDom As New MSXML2.DOMDocument40
    oNccDom.async = False
    oNccDom.validateOnParse = False
    oNccDom.resolveExternals = False
    oNccDom.preserveWhiteSpace = False
    oNccDom.setProperty "SelectionLanguage", "XPath"
    oNccDom.setProperty "NewParser", True
Dim oSmilTextIdNodes As IXMLDOMNodeList
Dim oSmilTextIdNode As IXMLDOMNode
Dim sOldUri As String, sNewUri As String
Dim sSmilId As String
  'this function assumes that all doms are closed and sit in array
  
  On Error GoTo ErrHandler
  fncLinkFixSmilRef = False
  
  'parse the ncc
  If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(0).sDomData, oNccDom, objOwner) Then GoTo ErrHandler
  
' For each smilfile
  lSmilFileCount = 0
  For i = 1 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1 Then
      lSmilFileCount = lSmilFileCount + 1
      
      sSmilId = CStr(Format(lSmilFileCount, "0000"))
      If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(i).sDomData, oSmilDom, objOwner) Then GoTo ErrHandler
'     For each text@id
      Set oSmilTextIdNodes = oSmilDom.selectNodes("//text/@id")
      k = 0
      For Each oSmilTextIdNode In oSmilTextIdNodes
        k = k + 1
        sOldUri = LCase$(objOwner.objFileSetHandler.aOutFileSet(i).sFileName & "#" & oSmilTextIdNode.Text)
'       Give the text id a new value
        oSmilTextIdNode.Text = "rgn_txt_" & sSmilId & "_" & CStr(Format(k, "0000"))
        sNewUri = objOwner.objFileSetHandler.aOutFileSet(i).sFileName & "#" & oSmilTextIdNode.Text
'      traverse ncc and modify any references to old value
'      (if existing) traverse contentdocs and modify any references to old value
        If Not fncChangeSmilUriReference(sOldUri, sNewUri, oNccDom, objOwner) Then GoTo ErrHandler
      Next 'For Each oSmilTextIdNode
      'save the modified smil back to array
      objOwner.objFileSetHandler.aOutFileSet(i).sDomData = oSmilDom.xml
    End If
    DoEvents
  Next 'For i = 1 To aOutFileSetMembers
  
 'save the modified ncc back to array
  objOwner.objFileSetHandler.aOutFileSet(0).sDomData = oNccDom.xml
    
  fncLinkFixSmilRef = True
ErrHandler:
  If Not fncLinkFixSmilRef Then objOwner.addlog "<errH in='fncLinkFixSmilRef' arrayItem='" & CStr(i) & "'>fncLinkFixSmilRef ErrH</errH>"
End Function

Private Function fncLinkFixTextRef(objOwner As oRegenerator, bolEstimateBrokenXhtmlLinks As Boolean) As Boolean
Dim i As Long, k As Long
Dim oSmilDom As New MSXML2.DOMDocument40
    oSmilDom.async = False
    oSmilDom.validateOnParse = False
    oSmilDom.resolveExternals = False
    oSmilDom.preserveWhiteSpace = False
    oSmilDom.setProperty "SelectionLanguage", "XPath"
    oSmilDom.setProperty "NewParser", True
Dim oNccDom As New MSXML2.DOMDocument40
    oNccDom.async = False
    oNccDom.validateOnParse = False
    oNccDom.resolveExternals = False
    oNccDom.preserveWhiteSpace = False
    oNccDom.setProperty "SelectionLanguage", "XPath"
    oNccDom.setProperty "NewParser", True
Dim oSmilTextSrcNodes As IXMLDOMNodeList
Dim oSmilTextSrcNode As IXMLDOMNode
Dim sSmilId As String
Dim sNewId As String
Dim lSmilFileCount As Long
Dim bNccIsContentCarrier As Boolean
Dim oNccNodes As IXMLDOMNodeList
Dim objIdGetter As oIdGetter
'below dimmed here for optimization in fncModifyTargetId
Dim oTargetDom As New MSXML2.DOMDocument40
    oTargetDom.async = False
    oTargetDom.validateOnParse = False
    oTargetDom.resolveExternals = False
    oTargetDom.preserveWhiteSpace = False
    oTargetDom.setProperty "SelectionLanguage", "XPath"
    oTargetDom.setProperty "NewParser", True
Dim sLastUsedFileName As String
Dim lArrayItem As Long
'mg20030330: for support of notepairs in skippable dtbs:
Dim bolHasNoteBodies As Boolean
Dim sPreviousNiceLink As String
  'this function assumes that all doms are closed and sit in array
  
  On Error GoTo ErrHandler
  
  Set objIdGetter = New oIdGetter
  
  fncLinkFixTextRef = False
  lSmilFileCount = 0
'  lTextSrcArrayItems = 0
' For each smilfile
  For i = 1 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1 Then
      lSmilFileCount = lSmilFileCount + 1
      sSmilId = CStr(Format(lSmilFileCount, "0000"))
      If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(i).sDomData, oSmilDom, objOwner) Then GoTo ErrHandler
'     for each text@src
      Set oSmilTextSrcNodes = oSmilDom.selectNodes("//text/@src")
      If oSmilTextSrcNodes Is Nothing Then
        objOwner.addlog "<error in='fncLinkFixTextRef'>oSmilTextSrcNodes is nothing in fncLinkFixTextRef</error>"
        GoTo ErrHandler
      End If
      For Each oSmilTextSrcNode In oSmilTextSrcNodes
        k = k + 1
'       create a new id value
        'sNewId = "rgn_cnt_" & sSmilId & "_" & CStr(Format(k, "0000"))
        sNewId = "rgn_cnt_" & objIdGetter.Getid
'       (if uri resolves ok) modify target id value
        
        '@@mg20030320:
        'if several TextSrc nodes point to the same destination
        'subsequent tries to locate will be unsuccessfull
        'therefore store TextSrcNode.Text in an array and replace if used before
        Dim lTextSrcArrayItem As Long
        Dim sUriToUse As String
        
        'check if current text@src has been active before
        lTextSrcArrayItem = fncIsObjectInTextSrcArray(oSmilTextSrcNode.Text)
        
        If lTextSrcArrayItem = -1 Then
          'not active before
          sUriToUse = oSmilTextSrcNode.Text
          If Not fncAddObjectToTextSrcArray(oSmilTextSrcNode.Text, sNewId) Then GoTo ErrHandler
        Else
          'active before, lTextSrcArrayItem is array item number
          sUriToUse = fncStripId(TextSrcArray(lTextSrcArrayItem).sTextSrc) & "#" & TextSrcArray(lTextSrcArrayItem).sNewId
          sNewId = TextSrcArray(lTextSrcArrayItem).sNewId
        End If
                
        If fncModifyTargetId(sUriToUse, sNewId, oTargetDom, sLastUsedFileName, lArrayItem, bolHasNoteBodies, objOwner) Then
'         give same value to text src fragment part
          oSmilTextSrcNode.Text = fncStripId(oSmilTextSrcNode.Text) & "#" & sNewId
          sPreviousNiceLink = oSmilTextSrcNode.Text
'         check if this was pointing to ncc document
          If InStr(1, oSmilTextSrcNode.Text, "ncc.htm", vbTextCompare) Then
            bNccIsContentCarrier = True
            oSmilTextSrcNode.Text = "ncc.html" & "#" & sNewId
            sPreviousNiceLink = oSmilTextSrcNode.Text
          End If
        Else
          'objOwner.addlog "<warning in='fncLinkFixTextRef'>warning: " & oSmilTextSrcNode.Text & " did not resolve</warning>"
          'mg200401: this gives value of previous text@src to broken textsrc
          If bolEstimateBrokenXhtmlLinks Then
          'the bool above just taken as an indicator that estimates should be done at all
            If sPreviousNiceLink <> "" Then
              objOwner.addlog "<warning in='fncLinkFixTextRef'>warning: estimated new value " & sPreviousNiceLink & " for " & oSmilTextSrcNode.Text & " </warning>"
              oSmilTextSrcNode.Text = sPreviousNiceLink
            Else
              objOwner.addlog "<warning in='fncLinkFixTextRef'>warning: broken link " & oSmilTextSrcNode.Text & " could not be estimated </warning>"
            End If
          End If
        End If
      Next 'For Each oSmilTextSrcNode In
'     save the modified smil back to array
      objOwner.objFileSetHandler.aOutFileSet(i).sDomData = oSmilDom.xml
    End If
    DoEvents
  Next i
  'for the very last replace done in fncModifyTargetId
  objOwner.objFileSetHandler.aOutFileSet(lArrayItem).sDomData = oTargetDom.xml

' If no text/src where pointing to ncc
' rename ids in ncc
  If Not bNccIsContentCarrier Then
    If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(0).sDomData, oNccDom, objOwner) Then GoTo ErrHandler
    Set oNccNodes = oNccDom.selectNodes("//h1" & _
                                      "| //h2" & _
                                      "| //h3" & _
                                      "| //h4" & _
                                      "| //h5" & _
                                      "| //h6" & _
                                      "| //span" & _
                                      "| //div") '!xht
    If Not fncAddId(oNccNodes, "rgn_ncc_", objOwner, "0000", False) Then GoTo ErrHandler
    objOwner.objFileSetHandler.aOutFileSet(0).sDomData = oNccDom.xml
  End If 'Not bNccIsContentCarrier
    
  If bolHasNoteBodies Then
    'this bool was set byref when current func called fncModifyTargetId above
    'it means that this may be a skippable DTB with bodyref attributes
    'this if clause modifies the bodyref content
    'using the sNoteBodyUriArray that was filled in fncModifyTargetId
    Dim r As Long, s As Long
    Dim oBodyRefNodes As IXMLDOMNodeList
    Dim oBodyRefNode As IXMLDOMNode
    Dim bolModded As Boolean
    
    'go through contentdoc array and mod any bodyref attribute
    For r = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
      'for each contentdoc
      If objOwner.objFileSetHandler.aOutFileSet(r).eType = TYPE_SMIL_CONTENT Then
        If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(r).sDomData, oTargetDom, objOwner) Then GoTo ErrHandler
        Set oBodyRefNodes = oTargetDom.selectNodes("//@bodyref")
        If Not oBodyRefNodes Is Nothing Then
          For Each oBodyRefNode In oBodyRefNodes
            bolModded = False
            'go through sNoteBodyUriArray and look for a match
            For s = 0 To lNoteBodyUriArrayItems - 1
              If oBodyRefNode.Text = sNoteBodyUriArray(s).sTextSrc Then
                'bodyref attr value is whole uri (name#id)
                 oBodyRefNode.Text = fncGetFileName(oBodyRefNode.Text) & "#" & sNoteBodyUriArray(s).sNewId
                 bolModded = True
              ElseIf oBodyRefNode.Text = "#" & fncGetId(sNoteBodyUriArray(s).sTextSrc) Then
                'bodyref attr is id only (#id)
                'if it is the same document then rename
                If sNoteBodyUriArray(s).sOwnerDocName = objOwner.objFileSetHandler.aOutFileSet(r).sFileName Then
                  oBodyRefNode.Text = "#" & sNoteBodyUriArray(s).sNewId
                  bolModded = True
                End If
              End If
              
            Next
            If Not bolModded Then
              'it might be that the notebody elem is not refd from smil (if its just a container)
              'in this case, the orig references are maintained.
              'check that this is so
              Dim oTestNode As IXMLDOMNode
              Set oTestNode = oTargetDom.selectSingleNode("//*[@id='" & fncGetId(oBodyRefNode.Text) & "']")
              If oTestNode Is Nothing Then
                objOwner.addlog "<warning in='fncLinkFixTextRef'>" & oBodyRefNode.xml & " not modified.</warning>"
              End If
            End If
          Next
          'save mods back to array
          objOwner.objFileSetHandler.aOutFileSet(r).sDomData = oTargetDom.xml
        End If 'Not oBodyRefNodes Is Nothing
      End If 'objOwner.objFileSetHandler.aOutFileSet(r).eType = TYPE_SMIL_CONTENT
    Next r
  End If 'bolHasNoteBodies
  
  fncLinkFixTextRef = True
ErrHandler:
  Set objIdGetter = Nothing
  'lTextSrcArrayItems = 0: ReDim Preserve TextSrcArray(lTextSrcArrayItems)
  If Not fncLinkFixTextRef Then objOwner.addlog "<errH in='fncLinkFixTextRef' smilFile='" & CStr(i) & "' textNode='" & CStr(k) & "'>fncLinkFixTextRef ErrH</errH>"
End Function

Private Function fncModifyTargetId( _
    ByVal sUri As String, _
    ByVal sNewId As String, _
    ByRef oTargetDom As MSXML2.DOMDocument40, _
    ByRef sLastUsedFileName As String, _
    ByRef lArrayItem As Long, _
    ByRef bolHasNoteBodies As Boolean, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim sTargetFileName As String
Dim sTargetId As String
Dim oContextIdNode As IXMLDOMNode
Dim oContextIdNodeParentClass As IXMLDOMNode
  
  fncModifyTargetId = False
   
  sTargetFileName = fncStripId(sUri)
  sTargetId = fncGetId(sUri)
    
  'if the same target file is used, no need to reparse
  'oTargetDom and lArrayItem will remain since resides at caller
  If Not LCase$(sTargetFileName) = LCase$(sLastUsedFileName) Then
    'the target file to be modified is different than previous
    'save result of previous call into array (if there has been a previous call)
    If oTargetDom.xml <> "" Then objOwner.objFileSetHandler.aOutFileSet(lArrayItem).sDomData = oTargetDom.xml
    lArrayItem = objOwner.objFileSetHandler.fncGetArrayItemFromName(sTargetFileName)
    If lArrayItem > -1 Then
      If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(lArrayItem).sDomData, oTargetDom, objOwner) Then GoTo ErrHandler
      sLastUsedFileName = objOwner.objFileSetHandler.aOutFileSet(lArrayItem).sFileName
    Else
      'lArrayItem = -1
      objOwner.addlog "<error in='fncModifyTargetId'>" & sTargetFileName & " not found in array.</error>"
      GoTo ErrHandler
    End If
  End If

  'set contextid, this should work since fragment lcase has been done earlier
  
  Set oContextIdNode = oTargetDom.selectSingleNode("//*[@id='" & sTargetId & "']/@id")
  If oContextIdNode Is Nothing Then
    objOwner.addlog "<error in='fncModifyTargetId'>id value " & sTargetId & " not found in " & sTargetFileName & "</error>"
    Exit Function
  Else
    oContextIdNode.Text = sNewId
    'if this id sits on a class="notebody",
    'there might be a bodyref attr in any of the content docs
    'that references the old id
    Set oContextIdNodeParentClass = oContextIdNode.selectSingleNode("../@class")
    If Not oContextIdNodeParentClass Is Nothing Then
      If LCase$(oContextIdNodeParentClass.Text) = "notebody" Then
        'save the old id value and its ownerdoc URI
        bolHasNoteBodies = True
        ReDim Preserve sNoteBodyUriArray(lNoteBodyUriArrayItems)
        sNoteBodyUriArray(lNoteBodyUriArrayItems).sTextSrc = sTargetFileName & "#" & sTargetId
        sNoteBodyUriArray(lNoteBodyUriArrayItems).sNewId = sNewId
        sNoteBodyUriArray(lNoteBodyUriArrayItems).sOwnerDocName = sTargetFileName
        lNoteBodyUriArrayItems = lNoteBodyUriArrayItems + 1
      End If
    End If
  End If
            
  'note that the very last call to this fnc will not save back to array
  'therefore the fnc that calls this fnc must do that as the last thing
  
  fncModifyTargetId = True
   
ErrHandler:
  If Not fncModifyTargetId Then objOwner.addlog "<errH in='fncModifyTargetId' arrayItem='" & lArrayItem & "'>fncModifyTargetId ErrH</errH>"
End Function

Private Function fncAddObjectToTextSrcArray( _
  ByVal sTextSrc As String, _
  ByVal sNewId As String _
  ) As Boolean
        On Error GoTo ErrHandler
        fncAddObjectToTextSrcArray = False
            ReDim Preserve TextSrcArray(lTextSrcArrayItems)
            With TextSrcArray(lTextSrcArrayItems)
              .sTextSrc = sTextSrc
              .sNewId = sNewId
            End With
            lTextSrcArrayItems = lTextSrcArrayItems + 1
        fncAddObjectToTextSrcArray = True
ErrHandler:
       ' If Not fncAddObjectToTextSrcArray Then objOwner.addlog "<errH in='fncAddObjectToTextSrcArray'>fncAddObjectToTextSrcArray ErrH</errH>"
End Function

Private Function fncIsObjectInTextSrcArray( _
  ByVal sTextSrc As String _
  ) As Long
Dim i As Long
    'return item number, if no match return -1
    
    For i = 0 To lTextSrcArrayItems - 1
        If LCase$(TextSrcArray(i).sTextSrc) = LCase$(sTextSrc) Then
            fncIsObjectInTextSrcArray = i
            Exit Function
        End If
    Next i
    fncIsObjectInTextSrcArray = -1
End Function

Public Function fncMakeTrueNccOnly(lDtbType, objOwner As oRegenerator) As Boolean
Dim oNccDom As New MSXML2.DOMDocument40
    oNccDom.async = False
    oNccDom.validateOnParse = False
    oNccDom.resolveExternals = False
    oNccDom.preserveWhiteSpace = False
    oNccDom.setProperty "SelectionLanguage", "XPath"
    oNccDom.setProperty "NewParser", True
Dim oSmilDom As New MSXML2.DOMDocument40
    oSmilDom.async = False
    oSmilDom.validateOnParse = False
    oSmilDom.resolveExternals = False
    oSmilDom.preserveWhiteSpace = False
    oSmilDom.setProperty "SelectionLanguage", "XPath"
    oSmilDom.setProperty "NewParser", True
Dim oCntDom As New MSXML2.DOMDocument40
    oCntDom.async = False
    oCntDom.validateOnParse = False
    oCntDom.resolveExternals = False
    oCntDom.preserveWhiteSpace = False
    oCntDom.setProperty "SelectionLanguage", "XPath"
    oCntDom.setProperty "NewParser", True
Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode, oHrefTargetNode As IXMLDOMNode, oNccHrefNode As IXMLDOMNode
Dim i As Long, lArrayItem As Long, lTemp As Long
Dim oTextSrc As IXMLDOMNode
Dim sUri As String
  On Error GoTo ErrHandler
  fncMakeTrueNccOnly = False
  
' check that dtbtype inparam is nccOnly, else warn
' but fnc should work on full or partial text dtbs as well

' for each node in ncc
'  resolve its smil href URI (may be a par or a text)
'  pos at text@src
'  if this smil text src uri is not pointing to ncc (<text src="cnt.html#id001"/>)
'    if the text node is not the same, then warn
'    mod src content
'  end if
' next node in ncc

' For each smil src uri
'   if this smil text src uri is pointing to ncc, save the src value in tmpvar
'   if this smil text src uri is not pointing to ncc (<text src="cnt.html#id001"/>)
'     mod src value to tmpvar
'   end if
' Next smil src uri

  objOwner.addlog "<status>converting to true ncc only...</status>"
' check that dtbtype inparam is nccOnly, else warn
' but fnc should work on full or partial text dtbs as well
  If (lDtbType <> DTB_AUDIOONLY) And (lDtbType <> DTB_AUDIONCC) Then
    objOwner.addlog "<warning in='fncMakeTrueNccOnly'>warning: 'make true ncc only' is being run on a DTB that is not marked as ncc only</warning>"
  End If

  If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(0).sDomData, oNccDom, objOwner) Then GoTo ErrHandler
  Set oNodes = oNccDom.selectNodes("//body/*")
  If Not oNodes Is Nothing Then
    lArrayItem = -1
    lTemp = -1
'   for each node in ncc
    For Each oNode In oNodes
'     resolve its smil href URI (oHrefTargetNode may be a par or a text)
      Set oNccHrefNode = oNode.selectSingleNode("a/@href")
      If Not oNccHrefNode Is Nothing Then
        sUri = oNccHrefNode.Text
      Else
       objOwner.addlog "<warning in='fncMakeTrueNccOnly'>warning: ncc node without a/@href</warning> "
       GoTo nextNode
      End If
      lArrayItem = objOwner.objFileSetHandler.fncGetArrayItemFromName(fncStripId(sUri))
      If lArrayItem < 0 Then
        objOwner.addlog "<error in='fncGetUriTargetNode'>filename " & fncStripId(sUri) & " not found in array</errH>"
        GoTo ErrHandler
      End If
      
      If (oNode Is oNodes.Item(0)) Then
        'if this is the first iterat
        If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(lArrayItem).sDomData, oSmilDom, objOwner) Then GoTo ErrHandler
      End If
      If (lTemp > -1) And (lArrayItem <> lTemp) Then
        'save and reparse only if doc changed and not if its first iterat
        objOwner.objFileSetHandler.aOutFileSet(lTemp).sDomData = oSmilDom.xml
        'If lArrayItem = 131 Then Stop
        If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(lArrayItem).sDomData, oSmilDom, objOwner) Then GoTo ErrHandler
      End If
      Set oHrefTargetNode = oSmilDom.selectSingleNode("//*[@id='" & fncGetId(sUri) & "']")
      If oHrefTargetNode Is Nothing Then
        lTemp = lArrayItem
        GoTo nextNode
      End If
      If oHrefTargetNode.nodeName = "par" Then Set oHrefTargetNode = oHrefTargetNode.selectSingleNode("text")
      If oHrefTargetNode Is Nothing Then
        lTemp = lArrayItem
        GoTo nextNode
      End If
'     go to at text@src
      Set oTextSrc = oHrefTargetNode.selectSingleNode("@src")
'     if this smil text src uri is not pointing to ncc (<text src="cnt.html#id001"/>)
      If InStr(1, oTextSrc.Text, "ncc.html") < 1 Then

'       if the text node is not the same, then warn
          'not done: prob is not known if linkback in contentdoc
'       mod src content
        oTextSrc.Text = "ncc.html#" & oNode.selectSingleNode("@id").Text
        'objOwner.addlog ("modded " & sUri & " to " & oTextSrc.Text)
      End If 'InStr(1, oTextSrc.Text, "ncc.html") < 1
      lTemp = lArrayItem
nextNode:
    Next oNode
    'save the very last iterat
        
    objOwner.objFileSetHandler.aOutFileSet(lTemp).sDomData = oSmilDom.xml
  End If 'Not oNodes Is Nothing

  'give sUri a default value for fallback below
  sUri = "ncc.html#" & oNccDom.selectSingleNode("//body/*/@id").Text
  
  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1 Then
      If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(i).sDomData, oSmilDom, objOwner) Then GoTo ErrHandler
      Set oNodes = oSmilDom.selectNodes("//text/@src")
        For Each oNode In oNodes
'         if this smil text src uri is pointing to ncc, save the src value in tmpvar
          If InStr(1, oNode.Text, "ncc.html") > 0 Then sUri = oNode.Text
'         if this smil text src uri is not pointing to ncc: mod src value to tmpvar
          If InStr(1, oNode.Text, "ncc.html") < 1 Then oNode.Text = sUri
        Next oNode '
      objOwner.objFileSetHandler.aOutFileSet(i).sDomData = oSmilDom.xml
    End If
' Next smil src uri
  Next

  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    'make the content docs obsolete
    If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_CONTENT Then
      objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_DELETED
    End If
    'make any aux file refd by content doc obsolete
    If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_OTHER And _
     objOwner.objFileSetHandler.aOutFileSet(i).lOwnerType = TYPE_SMIL_CONTENT Then
     objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_DELETED
    End If
  Next
  
  fncMakeTrueNccOnly = True
ErrHandler:
  Set oNccDom = Nothing
  Set oSmilDom = Nothing
  Set oCntDom = Nothing

  If Not fncMakeTrueNccOnly Then objOwner.addlog "<errH in='fncMakeTrueNccOnly'>fncMakeTrueNccOnly ErrH</errH>"
End Function

'      If Not fncGetUriTargetNode(oNccNode.selectSingleNode("a/@href").Text, oSmilDom, oHrefTargetNode, lArrayItem, objOwner) Then GoTo ErrHandler

'Public Function fncGetUriTargetNode( _
'  ByVal sUri As String, _
'  ByRef oDom As MSXML2.DOMDocument40, _
'  ByRef oTargetNode As IXMLDOMNode, _
'  ByRef lArrayItem As Long, _
'  ByRef objOwner As oRegenerator _
'  ) As Boolean
'Dim lTmp As Long
'Dim lPrevItem As Long
'  On Error GoTo ErrHandler
'  fncGetUriTargetNode = False
'
'  lPrevItem = lArrayItem
'  'check if input larrayitem matches sUri filename, if so assume that doc is already parsed
'  lTmp = objOwner.objFileSetHandler.fncGetArrayItemFromName(fncStripId(sUri))
'  If lArrayItem <> lTmp Then
'    lArrayItem = lTmp
'    If lArrayItem < 0 Then
'      objOwner.addlog "<error in='fncGetUriTargetNode'>filename " & fncStripId(sUri) & " not found in array</errH>"
'    End If
'    objOwner.objFileSetHandler.aOutFileSet(lPrevItem).sDomData = oDom.xml
'    If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(lArrayItem).sDomData, oDom, objOwner) Then GoTo ErrHandler
'  End If
'  Set oTargetNode = oDom.selectSingleNode("//*[@id='" & fncGetId(sUri) & "']")
'
'  fncGetUriTargetNode = True
'
'ErrHandler:
'  If Not fncGetUriTargetNode Then objOwner.addlog "<errH in='fncGetUriTargetNode'>fncGetUriTargetNode ErrH</errH>"
'End Function

