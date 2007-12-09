Attribute VB_Name = "mInternalXhtml"
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

'contains functions that apply to both ncc and content

Public Function fncSetLangAttrs( _
    ByRef oDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator, _
    ByVal lCurrentArrayItem As Long _
    ) As Boolean
Dim i As Long

Dim oHtmlNode As IXMLDOMNode
Dim oLangNodes As IXMLDOMNodeList
Dim oLangNode As IXMLDOMNode
Dim oNode As IXMLDOMNode
Dim sLangString As String
Dim bValidLangNodes As Boolean
    
  'this function sets the lang attr on <html elem of ncc and content doc
  On Error GoTo ErrHandler
  fncSetLangAttrs = False
    
  Set oHtmlNode = oDom.selectSingleNode("//html")
  If oHtmlNode Is Nothing Then
    objOwner.addlog "<error in='fncSetLangAttrs'>oHtmlNode is nothing in fncSetLangAttrs</error>"
    Exit Function
  End If
            
  'check if dc:language is available
  'if so, get the first two chars from it
   If (Not objOwner.objCommonMeta.DcLanguage Is Nothing) Then
     If (objOwner.objCommonMeta.DcLanguage.selectSingleNode("@content").Text <> "") Then
         sLangString = Mid(Trim$(objOwner.objCommonMeta.DcLanguage.selectSingleNode("@content").Text), 1, 2)
       'check that the value found is valid
       If Not fncIsValidLangCode(sLangString, objOwner.objLangCodes.oLangCodes, objOwner) Then
         objOwner.addlog "<warning>no valid dc:language found in ncc to substitute lang attr</warning>"
         'either sLangstring was never set,
         'or sLangString contains unrecognized value
         '//REVISIT use userlcid or similar systemvariable to guess on value??
         'sLangString=systemvariable (current windows language)
       Else
         If Not fncAppendAttribute(oHtmlNode, "xml:lang", sLangString, objOwner) Then GoTo ErrHandler
         If Not fncAppendAttribute(oHtmlNode, "lang", sLangString, objOwner) Then GoTo ErrHandler
         fncSetLangAttrs = True
         Exit Function
       End If 'Not fncIsValidLangCode
     Else
       objOwner.addlog "<warning>empty dc:language in ncc is not a substitute for lang attr</warning>"
     End If '(oCommonmeta.DcLanguage.Text <> "")
   Else
     objOwner.addlog "<warning>no dc:language found in ncc to substitute for lang attr</warning>"
   End If '(Not oCommonmeta.DcLanguage Is Nothing)
      
   'if we are here, there was no valid dc:language
   'go and see if preexisting lang attributes
      
   Set oLangNodes = oHtmlNode.selectNodes("@lang | @xml:lang")
   If (Not oLangNodes Is Nothing) Then
     If (oLangNodes.length > 0) Then
       bValidLangNodes = True
       For Each oLangNode In oLangNodes
         If fncIsValidLangCode(oLangNode.Text, objOwner.objLangCodes.oLangCodes, objOwner) Then
           'do nothing
         Else
           objOwner.addlog "<message>invalid lang code " & oLangNode.Text & " found.</message>"
           bValidLangNodes = False
         End If
       Next
     End If
   End If
   'only if all (if any) langnodes where valid is bValidLangNodes true
   If bValidLangNodes Then
     fncSetLangAttrs = True
     Exit Function
   Else
     If (Not oLangNodes Is Nothing) Then
       If (oLangNodes.length > 0) Then
         For Each oLangNode In oLangNodes
           'remove the invalid lang attrs
           If Not fncRemoveAttribute(oHtmlNode, oLangNode.nodeName, objOwner) Then GoTo ErrHandler
         Next
       End If
     End If
   End If
                 
   fncSetLangAttrs = True
  
ErrHandler:
  If Not fncSetLangAttrs Then objOwner.addlog "<errH in='fncSetLangAttrs' arrayItem='" & CStr(lCurrentArrayItem) & "'>fncSetLangAttrs ErrH</errH>"
End Function

Private Function fncIsValidLangCode( _
    sCandidate As String, _
    oLangCodes As IXMLDOMNodeList, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oLangCode As IXMLDOMNode

  On Error GoTo ErrHandler
  fncIsValidLangCode = False
  For Each oLangCode In objOwner.objLangCodes.oLangCodes
    If LCase$(oLangCode.Text) = LCase$(sCandidate) Then
      fncIsValidLangCode = True
      Exit For
    End If
  Next
  Exit Function
ErrHandler:
  objOwner.addlog "<errH in='fncIsValidLangCode'>fncIsValidLangCode ErrH</errH>"
End Function

Public Function fncFixAttrValueCase( _
  ByRef oDom As MSXML2.DOMDocument40, _
  ByRef objOwner As oRegenerator, _
  ByVal lDocumentType As Long, _
  ByRef lCurrentArrayItem As Long _
) As Boolean
'mg20030911: fix class attribute value case for specific daisy elements
'span.page-normal
'span.page-front
'span.page-special
'span.sidebar
'span.optional-prodnote
'span.noteref
'div.group
'div.notebody
Dim oClassAttrNodes As IXMLDOMNodeList
Dim oClassAttrNode As IXMLDOMNode

  On Error GoTo ErrHandler
  fncFixAttrValueCase = False
  If (lDocumentType = TYPE_NCC) Or (lDocumentType = TYPE_SMIL_CONTENT) Then
    Set oClassAttrNodes = oDom.selectNodes("//span/@class")
    If Not oClassAttrNodes Is Nothing Then
      For Each oClassAttrNode In oClassAttrNodes
        If (InStr(1, oClassAttrNode.Text, "page-normal", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "page-normal"
        ElseIf (InStr(1, oClassAttrNode.Text, "page-front", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "page-front"
        ElseIf (InStr(1, oClassAttrNode.Text, "page-special", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "page-special"
        ElseIf (InStr(1, oClassAttrNode.Text, "sidebar", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "sidebar"
        ElseIf (InStr(1, oClassAttrNode.Text, "optional-prodnote", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "optional-prodnote"
        ElseIf (InStr(1, oClassAttrNode.Text, "noteref", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "noteref"
        End If
      Next
    End If
  
    Set oClassAttrNodes = oDom.selectNodes("//div/@class")
    If Not oClassAttrNodes Is Nothing Then
      For Each oClassAttrNode In oClassAttrNodes
        If (InStr(1, oClassAttrNode.Text, "group", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "group"
        ElseIf (InStr(1, oClassAttrNode.Text, "notebody", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "notebody"
        End If
      Next
    End If
  
    Set oClassAttrNodes = oDom.selectNodes("//head/meta/@name")
    If Not oClassAttrNodes Is Nothing Then
      For Each oClassAttrNode In oClassAttrNodes
        If (InStr(1, oClassAttrNode.Text, "dc:title", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:title"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:identifier", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:identifier"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:contributor", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:contributor"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:coverage", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:coverage"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:creator", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:creator"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:date", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:date"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:description", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:description"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:format", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:format"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:language", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:language"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:narrator", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:narrator"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:producedDate", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:producedDate"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:producer", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:producer"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:publisher", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:publisher"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:relation", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:relation"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:revision", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:revision"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:revisionDate", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:revisionDate"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:rights", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:rights"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:source", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:source"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:sourceDate", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:sourceDate"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:sourceEdition", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:sourceEdition"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:sourcePublisher", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:sourcePublisher"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:sourceRights", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:sourceRights"
        ElseIf (InStr(1, oClassAttrNode.Text, "ncc:sourceTitle", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "ncc:sourceTitle"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:subject", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:subject"
        ElseIf (InStr(1, oClassAttrNode.Text, "dc:type", vbTextCompare) > 0) Then
          oClassAttrNode.Text = "dc:type"
        End If
      Next
    End If
  ElseIf (lDocumentType = TYPE_SMIL_1) Or (lDocumentType = TYPE_SMIL_MASTER) Then
    'this should not be needed: all smil metadata is redone anyway
  End If
    
  fncFixAttrValueCase = True

ErrHandler:
  Set oClassAttrNodes = Nothing
  Set oClassAttrNode = Nothing
  If Not fncFixAttrValueCase Then objOwner.addlog "<errH in='fncFixAttrValueCase' arrayItem='" & CStr(lCurrentArrayItem) & "'>fncFixAttrValueCase ErrH</errH>"
End Function


Public Function fncFixPageNums( _
    ByRef oDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator, _
    ByRef lCurrentArrayItem As Long _
    ) As Boolean
Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim oTextNode As IXMLDOMNode
  'this function:
  'trims any whitespace surrounding pagenum textnodes
  'if page-normal contains alphanum chars, convert to page-special
    
  On Error GoTo ErrHandler
  fncFixPageNums = False
  
  Set oNodeList = oDom.selectNodes _
    ("//span[@class='page-normal']/a" & _
    "| //span[@class='page-special']/a" & _
    "| //span[@class='page-front']/a")
    
  'always trim whitespace
  For Each oNode In oNodeList
    oNode.Text = Trim$(oNode.Text)
  Next
    
  Set oNodeList = oDom.selectNodes("//span[@class='page-normal']")
  
  'if page-normal contains alphanum chars, convert to page-special
  For Each oNode In oNodeList
    Set oTextNode = oNode.selectSingleNode("a")
    If oTextNode Is Nothing Then GoTo nextIterat
    If Not IsNumeric(oTextNode.Text) Then
      'remove class attr
      If Not fncRemoveAttribute(oNode, "class", objOwner) Then GoTo ErrHandler
      'add new
      If Not fncAppendAttribute(oNode, "class", "page-special", objOwner) Then GoTo ErrHandler
    End If
nextIterat:
  Next
    
  fncFixPageNums = True
ErrHandler:
  If Not fncFixPageNums Then objOwner.addlog "<errH in='fncFixPageNums' arrayItem='" & CStr(lCurrentArrayItem) & "'>fncFixPageNums ErrH</errH>"
End Function

Public Function fncChangeSmilUriReference( _
    ByVal sOldUri As String, _
    ByVal sNewUri As String, _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oContentDom As New MSXML2.DOMDocument40
    oContentDom.async = False
    oContentDom.validateOnParse = False
    oContentDom.resolveExternals = False
    oContentDom.preserveWhiteSpace = False
    oContentDom.setProperty "SelectionLanguage", "XPath"
    oContentDom.setProperty "NewParser", True
Dim i As Long
  
'this function modifies ncc and content smilrefs
'since ncc may be open in other funcs when this is called, it is sent in byref
'the function assumes that no contentdocs are opened in other funcs
'sOldUri comes in lcased
'sNewUri comes in without lcase

  On Error GoTo ErrHandler
  fncChangeSmilUriReference = False
  
  'first do the ncc dom
  If Not fncHrefReplace(oNccDom, sOldUri, sNewUri, objOwner) Then GoTo ErrHandler
  
  'then do the contentdocs, use the dedicated contentdoc pointer array
  If objOwner.objFileSetHandler.aOutFileSetContentDocMembers > 0 Then
    For i = 0 To objOwner.objFileSetHandler.aOutFileSetContentDocMembers - 1
      'If Not fncParseString(aOutFileSet(i).sDomData, oContentDom) Then GoTo ErrHandler
      'If Not fncHrefReplace(oContentDom, sOldUri, sNewUri) Then GoTo ErrHandler
      'aOutFileSet(i).sDomData = oContentDom.xml
      
      'do it dirty instead, saves time
       objOwner.objFileSetHandler.aOutFileSet(objOwner.objFileSetHandler.aOutFileSetContentDocs(i)).sDomData = Replace$( _
         objOwner.objFileSetHandler.aOutFileSet(objOwner.objFileSetHandler.aOutFileSetContentDocs(i)).sDomData, _
         """" & sOldUri, _
         """" & sNewUri, 1, -1, vbTextCompare)
         '"""" & sNewUri, 1, 1, vbTextCompare)
    Next i
  End If
    
  fncChangeSmilUriReference = True
ErrHandler:
  If Not fncChangeSmilUriReference Then objOwner.addlog "<errH in='fncChangeSmilUriReference'>fncChangeSmilUriReference ErrH</errH>"
End Function

Private Function fncHrefReplace( _
 ByRef oDom As MSXML2.DOMDocument40, _
 ByVal sOldUri As String, _
 ByVal sNewUri As String, _
 ByRef objOwner As oRegenerator _
 ) As Boolean
Dim oHrefList As IXMLDOMNodeList
Dim oHref As IXMLDOMNode
 
 On Error GoTo ErrHandler
 fncHrefReplace = False
 
 Set oHrefList = oDom.selectNodes("//@href")
 If Not oHrefList Is Nothing Then
   For Each oHref In oHrefList
     If Trim(LCase$(oHref.Text)) = sOldUri Then
       oHref.Text = sNewUri
     End If
   Next
 Else
   objOwner.addlog "<warning in='fncHrefReplace'>no href nodes in fncChangeSmilUriReference</warning>"
 End If

 fncHrefReplace = True
ErrHandler:
  If Not fncHrefReplace Then objOwner.addlog "<errH in='fncHrefReplace'>fncHrefReplace ErrH</errH>"
End Function

Public Function fncStripEmptyElem( _
    ByRef ioDom As MSXML2.DOMDocument40, _
    ByVal sXpath As String, _
    objOwner As oRegenerator _
    ) As Boolean
Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim bolIsEmpty As Boolean
    
  On Error GoTo ErrHandler
  fncStripEmptyElem = False
  
  Set oNodes = ioDom.selectNodes(sXpath)
  If Not oNodes Is Nothing Then
    For Each oNode In oNodes
      bolIsEmpty = False
      'if the p has no childnode other than one text node
      If oNode.childNodes.length < 2 Then
        If (oNode.childNodes.length = 1) Then
          If (oNode.childNodes.Item(0).nodeType = NODE_TEXT) Then
            Dim sTemp As String
            sTemp = Trim$(oNode.nodeTypedValue)
            If sTemp = "" Then bolIsEmpty = True
            If Len(sTemp) = 1 And Asc(sTemp) = 160 Then bolIsEmpty = True
          End If
        Else
         'childnodes.length was 0
         bolIsEmpty = True
        End If 'oNode.childNodes.length = 1
      Else
       'there were more than one childnode so continue with next para
      End If 'oNode.childNodes.length > 1
      
      If bolIsEmpty Then
        objOwner.addlog "<message in='fncStripEmptyElem'>Stripped empty para:" & oNode.xml & "</message>"
        Set oNode = oNode.parentNode.removeChild(oNode)
      End If
    Next oNode
  End If 'Not oParNodes Is Nothing
  fncStripEmptyElem = True
ErrHandler:
  If Not fncStripEmptyElem Then objOwner.addlog "<errH in='fncStripEmptyElem'>fncStripEmptyElem ErrH</errH>"
End Function

Public Function fncAddCss( _
 oDom As MSXML2.DOMDocument40, _
 objOwner As oRegenerator _
 ) As Boolean
 On Error GoTo ErrHandler
 fncAddCss = False
 
 Dim oPreviousCssLink As IXMLDOMNode
 Dim oHeadNode As IXMLDOMNode
 Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
 Dim oFolder, oFiles As Object, oFile As Object
 
 Set oPreviousCssLink = oDom.selectSingleNode("//head/link[@rel='stylesheet'][@href][@type='text/css']")
 Set oHeadNode = oDom.selectSingleNode("//head")
 If oHeadNode Is Nothing Then
   objOwner.addlog "<error in='fncAddCss'>cant find headnode in fncAddCss: css not added</error>"
   fncAddCss = True
   Exit Function
 End If
 
 If oPreviousCssLink Is Nothing Then
   'only add a css if none is there before
   If oFSO.folderexists(sResourcePath) Then
     Set oFolder = oFSO.getfolder(oFSO.GetAbsolutePathName(sResourcePath))
     Set oFiles = oFolder.Files
     For Each oFile In oFiles
       If LCase$(fncGetExtension(oFile.Name)) = ".css" Then
         If Not objOwner.objFileSetHandler.fncIsObjectInOutputArray(oFile.Name) Then
           If Not objOwner.objFileSetHandler.fncAddObjectToOutputArray(TYPE_CSS_INSERTED, oFile.Name, "", "") Then GoTo ErrHandler
         End If
         If Not fncAppendChild(oHeadNode, "link", objOwner, "", "rel", "stylesheet", "href", oFile.Name, "type", "text/css") Then GoTo ErrHandler
         'objOwner.addlog "<message>added stylesheet " & oFile.Name & "</message>"
         Exit For
       End If
     Next
   End If
 Else
   'previous stylesheet left, new one not added
 End If 'oPreviousCssLink

 fncAddCss = True
ErrHandler:
  If Not fncAddCss Then objOwner.addlog "<errH in='fncAddCss'>fncAddCss ErrH</errH>"
End Function

