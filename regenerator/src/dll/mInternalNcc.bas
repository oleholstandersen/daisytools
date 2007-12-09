Attribute VB_Name = "mInternalNcc"
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

Public Function fncUpdateTotalTimeAndFiles( _
    lTotalTimeMs As Long, _
    ByRef objOwner As oRegenerator) _
    As Boolean
Dim oNccDom As New MSXML2.DOMDocument40
    oNccDom.async = False
    oNccDom.validateOnParse = False
    oNccDom.resolveExternals = False
    oNccDom.preserveWhiteSpace = False
    oNccDom.setProperty "SelectionLanguage", "XPath"
    oNccDom.setProperty "NewParser", True
Dim oNodeBase As IXMLDOMNode
Dim oNode As IXMLDOMNode
Dim lFileCount As Long, i As Long

  On Error GoTo ErrHandler
  fncUpdateTotalTimeAndFiles = False
  
  If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(0).sDomData, oNccDom, objOwner) Then GoTo ErrHandler
  Set oNodeBase = oNccDom.selectSingleNode("//head")
  If Not oNodeBase Is Nothing Then
    'this elem should not exist but remove if it does
    Set oNode = oNodeBase.selectSingleNode("/meta[name='ncc:totalTime']")
    If Not oNode Is Nothing Then Set oNode = oNodeBase.removeChild(oNode)
    'then add the whole thing
    If Not fncAppendChild(oNodeBase, "meta", objOwner, , "name", "ncc:totalTime", _
      "content", fncConvertMS2SmilClockVal(lTotalTimeMs, 0, False)) Then
        GoTo ErrHandler
    End If
    'update ncc:files (empty audio may have been inserted)
    'this elem should not exist but remove if it does
    Set oNode = oNodeBase.selectSingleNode("/meta[name='ncc:files']")
    If Not oNode Is Nothing Then Set oNode = oNodeBase.removeChild(oNode)
    
    For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
      If (objOwner.objFileSetHandler.aOutFileSet(i).eType <> TYPE_DELETED) Then
        'mg20050324 removed this inner clause since I dont understand it
        'If Not (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_OTHER And _
        'objOwner.objFileSetHandler.aOutFileSet(i).lOwnerType = TYPE_SMIL_CONTENT) Then
          lFileCount = lFileCount + 1
        'End If
      End If
    Next
    'then add the whole thing
    'mastersmil is created later so +1
    If Not fncAppendChild(oNodeBase, "meta", objOwner, , "name", "ncc:files", "content", lFileCount + 1) Then GoTo ErrHandler
  Else
    objOwner.addlog "<error>no head element found in ncc</error>"
  End If
    
  objOwner.objFileSetHandler.aOutFileSet(0).sDomData = oNccDom.xml
    
  fncUpdateTotalTimeAndFiles = True
  
ErrHandler:
  Set oNodeBase = Nothing
  Set oNode = Nothing
  If Not fncUpdateTotalTimeAndFiles Then objOwner.addlog "<errH in='fncUpdateTotalTimeAndFiles'>fncUpdateTotalTimeAndFiles ErrH" & Err.Number & " : " & Err.Description & "</errH>"
End Function
  
Public Function fncTidyNcc( _
  ByRef oNccDom As MSXML2.DOMDocument40, _
  sOrigNccPath As String, _
  lInputEncoding As Long, _
  sOutCharsetName As String, _
  objOwner As oRegenerator _
  ) As Boolean
Dim sNcc As String

  On Error GoTo ErrHandler
  fncTidyNcc = False
  
  objOwner.addlog "<status>tidying and parsing input ncc...</status>"
    
  'try to tidy the ncc, send in an empty string byref
  If TidyLib.fncRunTidy(sOrigNccPath, sNcc, lInputEncoding, TYPE_NCC, sOutCharsetName, objOwner) Then
    'if tidy returns something, try to parse it
    If sNcc <> "" Then
      If fncParseString(sNcc, oNccDom, objOwner) Then
        objOwner.objFileSetHandler.fncAddObjectToOutputArray TYPE_NCC, "ncc.html", oNccDom.xml
      Else
        objOwner.addlog "<error in='fncTidyNcc'>could not parse tidy processed ncc</error>"
        GoTo ErrHandler
      End If 'fncParseString(sNcc, oNccDom)
    Else
      objOwner.addlog "<error in='fncTidyNcc'>tidy process on input ncc returned null</error>"
      GoTo ErrHandler
    End If 'sNccXml <> ""
  Else
    GoTo ErrHandler
  End If 'fncRunTidy(sNccXml)
  
  fncTidyNcc = True
  
ErrHandler:
  If Not fncTidyNcc Then objOwner.addlog "<errH in='fncTidyNcc'>fncTidyNcc ErrH" & Err.Number & " : " & Err.Description & "</errH>"
End Function

Public Function fncInternalNcc( _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef bolPreserveBiblioMeta As Boolean, _
    ByRef sMetaPath As String, _
    ByRef sOutCharsetName As String, _
    ByRef bolIsMultiVolume As Boolean, _
    ByRef lDtbType As Long, _
    ByRef bolAddCss As Boolean, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim sDtbType As String

  On Error GoTo ErrHandler
  fncInternalNcc = False
    
  Select Case lDtbType
    Case DTB_AUDIOONLY
      sDtbType = "audioOnly"
    Case DTB_AUDIONCC
      sDtbType = "audioNcc"
    Case DTB_AUDIOPARTTEXT
      sDtbType = "audioPartText"
    Case DTB_AUDIOFULLTEXT
      sDtbType = "audioFullText"
    Case DTB_TEXTPARTAUDIO
      sDtbType = "textPartAudio"
    Case DTB_TEXTNCC
      sDtbType = "textNcc"
  End Select
  
  'mg20030911: fix pagenum case (class attr value)
  If Not fncFixAttrValueCase(oNccDom, objOwner, TYPE_NCC, 0) Then GoTo ErrHandler

  'redo pagenums before calculating meta
  If Not fncFixPageNums(oNccDom, objOwner, 0) Then GoTo ErrHandler
    
  If Not fncDoNccMetaData( _
    oNccDom, _
    bolPreserveBiblioMeta, _
    sMetaPath, _
    sOutCharsetName, _
    bolIsMultiVolume, _
    sDtbType, _
    objOwner _
    ) Then GoTo ErrHandler
      
  'create a class with metaelements that reoccur in other files
  If Not objOwner.objCommonMeta.fncSetCommonMetaNodes(oNccDom, objOwner) Then GoTo ErrHandler
  
  'set the dc identifier prop in oregenerator
  If Not objOwner.objCommonMeta.DcIdentifier Is Nothing Then
    objOwner.fncSetDcIdentifier (objOwner.objCommonMeta.DcIdentifier.selectSingleNode("@content").Text)
  Else
    objOwner.fncSetDcIdentifier ("unknown")
  End If
            
  If Not fncSetLangAttrs(oNccDom, objOwner, 0) Then GoTo ErrHandler
            
  If bolAddCss Then
    If Not fncAddCss(oNccDom, objOwner) Then GoTo ErrHandler
  End If
                
  If Not fncFixFragmentCase(oNccDom, TYPE_NCC, objOwner) Then GoTo ErrHandler
    
  'move http-equiv elem to first sibling pos
  If Not fncMoveSiblingToTop(oNccDom, "//head/meta[@http-equiv='Content-type']", objOwner) Then GoTo ErrHandler
    
  ' mg20030325 do what tidy should be doing; also causes illegal char in output
  ' unlikely there are p's in ncc but anyway
   If Not fncStripEmptyElem(oNccDom, "//p", objOwner) Then GoTo ErrHandler
   
  'mg 20030911: rename unallowed ncc elems to div.group
  'make sure that class attr value lcase is done (done in fncFixAttrValueCase)
  If Not fncRenameNonNccElem(oNccDom, objOwner) Then GoTo ErrHandler
    
  'mg20050324 fix h1.title
  If Not fncFixFirstH1Class(oNccDom, objOwner) Then GoTo ErrHandler
    
  'dont save oNccDom back to array here since it is still live in mGlobal
  fncInternalNcc = True

ErrHandler:
    If Not fncInternalNcc Then objOwner.addlog "<errH in='fncInternalNcc'>fncInternalNcc ErrH" & Err.Number & " : " & Err.Description & "</errH>"
End Function

Private Function fncDoNccMetaData( _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef bolPreserveBiblioMeta As Boolean, _
    ByRef sMetaPath As String, _
    ByRef sWantsThisEncoding As String, _
    ByRef bolIsMultiVolume As Boolean, _
    ByRef sDtbType As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oHeadNode As IXMLDOMNode
Dim oNode As IXMLDOMNode
Dim oNodeAtt As IXMLDOMNode
Dim oNodes As IXMLDOMNodeList
Dim oTempNode As IXMLDOMNode
Dim oForbiddenImportMetaDom As New MSXML2.DOMDocument40
    oForbiddenImportMetaDom.async = False
    oForbiddenImportMetaDom.validateOnParse = False
    oForbiddenImportMetaDom.resolveExternals = False
    oForbiddenImportMetaDom.preserveWhiteSpace = False
    oForbiddenImportMetaDom.setProperty "SelectionLanguage", "XPath"
    oForbiddenImportMetaDom.setProperty "NewParser", True

  On Error GoTo ErrHandler
  fncDoNccMetaData = False
  
  Select Case bolPreserveBiblioMeta
    Case True
      'dont touch the bibliographical metadata
      'start with removing identifyable dtb related metadata
      Set oHeadNode = oNccDom.selectSingleNode("//head")
      If oHeadNode Is Nothing Then
        objOwner.addlog "<error in='fncDoNccMetaData'>oHeadNode Is Nothing</error>"
        GoTo ErrHandler
      End If
      Set oNodes = oNccDom.selectNodes("//head/meta") '!xht
      If Not (oNodes.length = 0) Or (oNodes Is Nothing) Then
        For Each oNode In oNodes
          Set oNodeAtt = oNode.selectSingleNode("@name")
          If Not oNodeAtt Is Nothing Then
            Dim sTmp As String: sTmp = LCase$(Trim(oNodeAtt.Text))
            'get previous generator value before deleting it
            If sTmp = "ncc:generator" Then
              Dim sPreviousGenerator  As String
              Set oTempNode = oNode.selectSingleNode("@content")
              If Not oTempNode Is Nothing Then
                sPreviousGenerator = oTempNode.Text
                objOwner.addlog "<previousGenerator value='" & sPreviousGenerator & "'/>"
              End If
            End If
            'get previous revision number before deleting it
            If sTmp = "ncc:revision" Then
              Dim sPreviousRevision As String
              Set oTempNode = oNode.selectSingleNode("@content")
              If Not oTempNode Is Nothing Then sPreviousRevision = Trim$(oTempNode.Text)
            End If
            If sTmp = "ncc:identifier" Then
              Dim sNccIdentifier As String
              Set oTempNode = oNode.selectSingleNode("@content")
              If Not oTempNode Is Nothing Then sNccIdentifier = Trim$(oTempNode.Text)
            End If
            
            'mg20040218 special fix for colin garnham
            If sTmp = "dc:language" Then
              Set oTempNode = oNode.selectSingleNode("@content")
              If Not oTempNode Is Nothing Then
                If oTempNode.Text = "en English" Or oTempNode.Text = "en-English" Then
                  oTempNode.Text = "en"
                End If
              End If
            End If
            
            If sTmp = "ncc:page-normal" Or _
               sTmp = "ncc:pagenormal" Or _
               sTmp = "ncc:page-front" Or _
               sTmp = "ncc:pagefront" Or _
               sTmp = "ncc:page-special" Or _
               sTmp = "ncc:pagespecial" Or _
               sTmp = "ncc:totaltime" Or _
               sTmp = "ncc:tocitems" Or _
               sTmp = "ncc:generator" Or _
               sTmp = "ncc:charset" Or _
               sTmp = "ncc:depth" Or _
               sTmp = "ncc:files" Or _
               sTmp = "ncc:setinfo" Or _
               sTmp = "ncc:footnotes" Or _
               sTmp = "ncc:maxpagenormal" Or _
               sTmp = "ncc:multimediatype" Or _
               sTmp = "ncc:revision" Or _
               sTmp = "ncc:revisiondate" Or _
               sTmp = "ncc:prodnotes" Or _
               sTmp = "dc:format" Or _
               sTmp = "ncc:format" Or _
               sTmp = "ncc:sidebars" _
            Then
              Set oNode = oNode.parentNode.removeChild(oNode)
            End If
          End If 'Not oNodeAtt Is Nothing Then
            
          Set oNodeAtt = oNode.selectSingleNode("@http-equiv")
          If Not oNodeAtt Is Nothing Then
            sTmp = LCase$(Trim(oNodeAtt.Text))
            If sTmp = "content-type" Then
              Set oNode = oNode.parentNode.removeChild(oNode)
            End If
          End If
        Next
      End If 'Not (oNodes.length = 0) Or (oNodes Is Nothing)
    Case False 'remove all ncc meta elements
      'mg20050419, dont remove ncc:narrator
      If Not fncRemoveNodes(oNccDom, "//head/meta[not(@name='ncc:narrator')]", objOwner) Then
      'If Not fncRemoveNodes(oNccDom, "//head/meta", objOwner) Then '!xht
        objOwner.addlog "<error in='fncDoNccMetaData'>fncRemoveNodes oNccDom //head/meta fail</error>" '!xht
        GoTo ErrHandler
      End If
      
      'import bibliographical from external source and insert.
      'elements allowed in import document: dc:title dc:identifier dc:date dc:description dc:subject dc:language dc:publisher dc:source ncc:producedDate ncc:producer ncc:sourceDate ncc:sourceEdition ncc:sourcePublisher (Not ncc:revision ncc:revisiondate) ncc:narrator prod:*
      If sMetaPath <> "" Then  'if the optional input metapath parameter has been set
        Dim oMetaDom As New MSXML2.DOMDocument40
            oMetaDom.async = False
            oMetaDom.validateOnParse = False
            oMetaDom.resolveExternals = False
            oMetaDom.preserveWhiteSpace = False
            oMetaDom.setProperty "SelectionLanguage", "XPath"
            oMetaDom.setProperty "NewParser", True
            oMetaDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
        Dim oMetaNodeList As IXMLDOMNodeList
        Dim oMetaItem As IXMLDOMNode
        Dim oHead As IXMLDOMNode
        Dim oForbiddenMetaNodes As IXMLDOMNodeList
        
        If Not fncParseFile(sResourcePath & "forbiddenMeta.xml", oForbiddenImportMetaDom, objOwner) Then GoTo ErrHandler
        Set oForbiddenMetaNodes = oForbiddenImportMetaDom.selectNodes("//name")
        If oForbiddenMetaNodes Is Nothing Then GoTo ErrHandler
        
        If fncFileExists(sMetaPath, objOwner) Then
          If fncParseFile(sMetaPath, oMetaDom, objOwner) Then
            Set oHead = oNccDom.selectSingleNode("//head") '!xht
            Set oMetaNodeList = oMetaDom.selectNodes("//meta") '!xht
            If (oMetaNodeList.length <> 0) Or Not (oMetaNodeList Is Nothing) Then  'there was a meta doc pointed out but no docelement children
              For Each oMetaItem In oMetaNodeList
                'check that it is not a forbidden elem
                Dim oTestNode As IXMLDOMNode
                Set oTestNode = oMetaItem.selectSingleNode("@name")
                If Not oTestNode Is Nothing Then
                  If Not fncIsForbiddenImportMeta(oTestNode.Text, oForbiddenMetaNodes, objOwner) Then
                    If oTestNode.Text = "ncc:narrator" Then
                      'the import meta is narrator
                      'mg20050419
                      'remove narrator in ncc if existing
                      If Not fncRemoveNodes(oHead.ownerDocument, "//head/meta[@name='ncc:narrator']", objOwner) Then GoTo ErrHandler
                    End If
                    oHead.appendChild oMetaItem
                  Else
                    objOwner.addlog "<message in='fncDoNccMetaData'>meta element " & oTestNode.Text & " excluded from import</message>"
                  End If 'fncIsForbiddenImportMeta
                End If 'Not oTestNode Is Nothing
              Next
            Else
              objOwner.addlog ("<error in='fncDoNccMetaData'>no meta elements found in meta import document " & fncGetFileName(sMetaPath) & "</error>")
              GoTo ErrHandler
            End If 'oMetaNodeList.length <> 0
          Else
            objOwner.addlog "<error in='fncDoNccMetaData'>external metadata document could not be parsed</error>"
            GoTo ErrHandler
          End If 'fncParseFile(sMetaPath)
        Else
          objOwner.addlog "<error in='fncDoNccMetaData'>external metadata document could not be found</error>"
          GoTo ErrHandler
        End If 'fncFileExists(sMetaPath)
      Else
         objOwner.addlog "<error in='fncDoNccMetaData'>external metadata document path was empty</error>"
         GoTo ErrHandler
      End If 'sMetaPath <> "" Then
      Set oMetaDom = Nothing
      
      Set oHeadNode = oHead
  End Select
  
  'now insert previous generator elem as prod:
  If sPreviousGenerator <> "" Then
    If Not fncAppendChild(oHeadNode, "meta", objOwner, , "name", "prod:prevGenerator", "content", sPreviousGenerator) Then GoTo ErrHandler
  End If
  'insert revision number and revisiondate
  If IsNumeric(sPreviousRevision) Then
    If Not fncAppendChild(oHeadNode, "meta", objOwner, , "name", "ncc:revision", "content", CStr(CLng(sPreviousRevision) + 1)) Then GoTo ErrHandler
  Else
    If Not fncAppendChild(oHeadNode, "meta", objOwner, , "name", "ncc:revision", "content", "1") Then GoTo ErrHandler
  End If
  'insert revision date
  If Not fncAppendChild(oHeadNode, "meta", objOwner, , "name", "ncc:revisionDate", "content", _
    CStr(DatePart("yyyy", Date)) & "-" & Format(CStr(DatePart("m", Date)), "00") & "-" & Format(CStr(DatePart("d", Date)), "00")) Then GoTo ErrHandler
  'for both cases, issue a warning if dc:identifier is missing
  Set oNode = oNccDom.selectSingleNode("//head/meta[@name='dc:identifier']") '!xht
  If oNode Is Nothing Then
    objOwner.addlog ("<warning>warning: dc:identifier is missing</warning>")
    'insert ncc:identifier as dc:identifier if it existed
    If sNccIdentifier <> "" Then
      objOwner.addlog "<message>warning: using found ncc:identifier for dc:identifier</message>"
      If Not fncAppendChild(oHeadNode, "meta", objOwner, , "name", "dc:identifier", "content", sNccIdentifier) Then GoTo ErrHandler
    End If
  End If
        
  'for both cases, create new dtbrelated and insert
  Dim oNccHead As IXMLDOMNode
  Set oNccHead = oNccDom.selectSingleNode("//head") '!xht
  If Not fncDoNccDtbMeta(oNccDom, oNccHead, sWantsThisEncoding, bolIsMultiVolume, sDtbType, objOwner) Then
    objOwner.addlog "<error>fncDoNccDtbMeta(oNccDom, oNccHead, sWantsThisEncoding) fail</error>"
    GoTo ErrHandler
  End If
  
  fncDoNccMetaData = True
  
ErrHandler:
 If Not fncDoNccMetaData Then objOwner.addlog "<errH in='fncDoNccMetaData'>fncDoNccMetaData ErrH</errH>"
End Function
        
Private Function fncDoNccDtbMeta( _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef oHead As IXMLDOMNode, _
    ByRef sOutCharsetName As String, _
    ByRef bolIsMultiVolume As Boolean, _
    ByRef sDtbType As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oNode As IXMLDOMNode
Dim oNode2 As IXMLDOMNode
Dim oBody As IXMLDOMNode
  On Error GoTo ErrHandler
  
  fncDoNccDtbMeta = False
    
  'title
  Set oNode = oNccDom.selectSingleNode("//head/meta[@name='dc:title']/@content") '!xht
  If Not oNode Is Nothing Then
    Set oNode2 = oNccDom.selectSingleNode("//head/title") '!xht
    If Not oNode2 Is Nothing Then
      oNode2.Text = oNode.Text
    Else
      If Not fncAppendChild(oHead, "title", objOwner, oNode.Text) Then GoTo ErrHandler
    End If
  Else
    objOwner.addlog ("<warning in='fncDoNccDtbMeta'>warning: dc:title missing</warning>")
  End If
  
  'dc:format
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "dc:format", "content", "Daisy 2.02") Then GoTo ErrHandler
  
  'ncc:generator
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:generator", "content", "Regenerator " & sAppVersion) Then GoTo ErrHandler
  
  'ncc:tocItems
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:tocItems", "content", CStr(fncCountNodes(oNccDom, "//body/*"))) Then GoTo ErrHandler '!xht xxx
  
  'ncc:pageNormal
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:pageNormal", "content", CStr(fncCountNodes(oNccDom, "//body/span[@class='page-normal']"))) Then GoTo ErrHandler '!xht
  
  'ncc:pageFront
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:pageFront", "content", CStr(fncCountNodes(oNccDom, "//body/span[@class='page-front']"))) Then GoTo ErrHandler '!xht
  
  'ncc:pageSpecial
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:pageSpecial", "content", CStr(fncCountNodes(oNccDom, "//body/span[@class='page-special']"))) Then GoTo ErrHandler '!xht
  
  'ncc:sidebars
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:sidebars", "content", CStr(fncCountNodes(oNccDom, "//body/span[@class='sidebar']"))) Then GoTo ErrHandler '!xht
  
  'ncc:prodNotes
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:prodnotes", "content", CStr(fncCountNodes(oNccDom, "//body/span[@class='optional-prodnote']"))) Then GoTo ErrHandler '!xht
  
  'ncc:footnotes
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:footnotes", "content", CStr(fncCountNodes(oNccDom, "//body/span[@class='noteref']"))) Then GoTo ErrHandler '!xht
  
  'ncc:depth
  Dim r As Long
  For r = 6 To 1 Step -1
    If fncCountNodes(oHead, "//body/h" & CStr(r)) > 0 Then Exit For
  Next r
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:depth", "content", CStr(r)) Then GoTo ErrHandler
  
  'ncc:MaxPageNormal
  Set oNode = oNccDom.selectSingleNode("//body/span[@class='page-normal'][last()]/a") '!xht
  If Not oNode Is Nothing Then
    If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:maxPageNormal", "content", oNode.Text) Then GoTo ErrHandler
  Else
    If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:maxPageNormal", "content", "0") Then GoTo ErrHandler
  End If
  
  'ncc:charset
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:charset", "content", sOutCharsetName) Then GoTo ErrHandler
  
  'http-equiv
  If Not fncAppendChild(oHead, "meta", objOwner, , "http-equiv", "Content-type", "content", "text/html; charset=" & sOutCharsetName) Then GoTo ErrHandler
  
  'ncc:multiMediaType
  If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:multimediaType", "content", sDtbType) Then GoTo ErrHandler
  
  If Not bolIsMultiVolume Then
    'ncc:setinfo should be maintained as-if if bolIsMultiVolume
    If Not fncAppendChild(oHead, "meta", objOwner, , "name", "ncc:setInfo", "content", "1 of 1") Then GoTo ErrHandler
  End If
      
  fncDoNccDtbMeta = True
    
ErrHandler:
  If Not fncDoNccDtbMeta Then objOwner.addlog "<errH in='fncDoNccDtbMeta'>fncDoNccDtbMeta ErrHandler</errH>"
End Function

Private Function fncIsForbiddenImportMeta( _
  ByVal sMetaName As String, _
  ByRef oForbiddenNames As IXMLDOMNodeList, _
  ByRef objOwner As oRegenerator _
  ) As Boolean
Dim oNode As IXMLDOMNode
  On Error GoTo ErrHandler
  fncIsForbiddenImportMeta = True

  For Each oNode In oForbiddenNames
    If Trim$(LCase$(oNode.Text)) = Trim$(LCase$(sMetaName)) Then Exit Function
  Next

  fncIsForbiddenImportMeta = False
ErrHandler:
  If fncIsForbiddenImportMeta Then objOwner.addlog "<errH in='fncIsForbiddenImportMeta'>fncIsForbiddenImportMeta ErrH</errH>"
End Function

Private Function fncFixFirstH1Class( _
  ByRef oNccDom As MSXML2.DOMDocument40, _
  ByRef objOwner As oRegenerator) _
  As Boolean
'this function makes sure there is a class='title' on the first child of body
  Dim oBodyFirstChild As IXMLDOMElement
  Dim oAttr As IXMLDOMAttribute
  
  On Error GoTo ErrHandler
  fncFixFirstH1Class = False
  
  Set oBodyFirstChild = oNccDom.selectSingleNode("//body/h1")
  
  If Not oBodyFirstChild Is Nothing Then
    Set oAttr = oBodyFirstChild.Attributes.getNamedItem("class")
    If Not oAttr Is Nothing Then
      oAttr.nodeValue = "title"
    Else
      Set oAttr = oNccDom.createAttribute("class")
      Set oAttr = oBodyFirstChild.Attributes.setNamedItem(oAttr)
      oAttr.nodeValue = "title"
    End If
  End If
      
  fncFixFirstH1Class = True
ErrHandler:
  If Not fncFixFirstH1Class Then objOwner.addlog "<errH in='fncFixFirstH1Class'>fncFixFirstH1Class ErrH</errH>"
End Function

Private Function fncRenameNonNccElem( _
  ByRef oNccDom As MSXML2.DOMDocument40, _
  ByRef objOwner As oRegenerator) _
  As Boolean
'this function locates all elements that are not allowed in ncc
'and renames them to div.group

'span.page-normal
'span.page-front
'span.page-special
'span.sidebar
'span.optional-prodnote
'span.noteref
'div.group

Dim oBodyChildren As IXMLDOMNodeList
Dim oBodyChild As IXMLDOMNode
Dim oBodyChildClassAttr As IXMLDOMNode

  On Error GoTo ErrHandler
  fncRenameNonNccElem = False

' select all first level body children
  Set oBodyChildren = oNccDom.selectNodes("//body/*")
  If Not oBodyChildren Is Nothing Then
    For Each oBodyChild In oBodyChildren
      Set oBodyChildClassAttr = oBodyChild.selectSingleNode("@class")
      Select Case oBodyChild.nodeName
        Case "h1", "h2", "h3", "h4", "h5", "h6"
          'ok
        Case "span"
          If Not oBodyChildClassAttr Is Nothing Then
            Select Case oBodyChildClassAttr.Text
              Case "page-normal", "page-front", "page-special", "sidebar", "optional-prodnote", "noteref"
                'ok
              Case Else
                'not ok
                objOwner.addlog "<message>Renaming " & oBodyChild.xml & "to div.group</message>"
                If Not fncAppendAttribute(oBodyChild, "class", "group", objOwner) Then GoTo ErrHandler
                If Not fncRenameElementNode(oBodyChild, "div", objOwner) Then GoTo ErrHandler
            End Select
          Else
            'not ok
            objOwner.addlog "<message>Renaming " & oBodyChild.xml & "to div.group</message>"
            If Not fncAppendAttribute(oBodyChild, "class", "group", objOwner) Then GoTo ErrHandler
            If Not fncRenameElementNode(oBodyChild, "div", objOwner) Then GoTo ErrHandler
          End If
        Case "div"
          If Not oBodyChildClassAttr Is Nothing Then
            Select Case oBodyChildClassAttr.Text
              Case "group"
                'ok
              Case Else
                'not ok
                objOwner.addlog "<message>Renaming " & oBodyChild.xml & "to div.group</message>"
                If Not fncAppendAttribute(oBodyChild, "class", "group", objOwner) Then GoTo ErrHandler
                If Not fncRenameElementNode(oBodyChild, "div", objOwner) Then GoTo ErrHandler
            End Select
          Else
            'not ok
            objOwner.addlog "<message>Renaming " & oBodyChild.xml & "to div.group</message>"
            If Not fncAppendAttribute(oBodyChild, "class", "group", objOwner) Then GoTo ErrHandler
            If Not fncRenameElementNode(oBodyChild, "div", objOwner) Then GoTo ErrHandler
          End If
        Case Else
          'not ok
           objOwner.addlog "<message>Renaming &lt;" & oBodyChild.nodeName & " &gt; to div.group in ncc</message>"
           If Not fncAppendAttribute(oBodyChild, "class", "group", objOwner) Then GoTo ErrHandler
           If Not fncRenameElementNode(oBodyChild, "div", objOwner) Then GoTo ErrHandler
      End Select
    Next
  End If 'Not oBodyChildren Is Nothing

  fncRenameNonNccElem = True
  
ErrHandler:
  Set oBodyChildren = Nothing
  Set oBodyChild = Nothing
  If Not fncRenameNonNccElem Then objOwner.addlog "<errH in='fncRenameNonNccElem'>fncRenameNonNccElem ErrH</errH>"
End Function

'Private Function fncGetMultiMediaType( _
'    ByRef oNccDom As MSXML2.DOMDocument40, _
'    ByRef sResult As String _
'    ) As Boolean
    
  'Dim sResult As String
  'If Not fncGetMultiMediaType(oNccDom, sResult) Then objowner.addlog "unable to determine ncc:multiMediaType"
'types:
'type 1 "audioOnly"; Full audio with Title element only.
'type 2 "audioNcc"; Full audio with NCC only.
'type 3 "audioPartText"; Full audio with NCC and partial text: [structure and some additional text]
'type 4 "audioFullText"; Full audio and full text: [structure and complete text and audio]
'type 5 "textPartAudio"; Full text and some audio: [structure, complete text and limited audio]
'type 6 "textNcc"; Text and no audio: [structured electronic text only. No audio is present.]

'Since ncc-only books are done in two ways;
'a) all smil <text src pointing back to ncc.html
'b) smil <text src pointing to another xhtml doc but only containg a duplicate nodeset to that of the ncc body
'...it is not possible to use text src destination filetype to determine a basic type
'Instead the check sequence would be

'1. If only one document+fragment is pointed out by all smil <text src,
'and that one document contains as child of body only one element H1,
'then dtb is type 1 "audioOnly"
'(there is no rule saying that type 1 must be only one par, and in addition there are reasons to not produce type 1 dtbs like that (generic smil player navigation using pars)

'2. If one or several documents+fragments are pointed out by all smil <text src,
'and these document contains as child of body only the element set allowed in ncc,
'then dtb is type 2 "audioNcc"

'3. If one or several documents+fragments are pointed out by all smil <text src,
'and these document contains as child of body the element set allowed in ncc and additional xhtml elements,
'then dtb is type 3 "audioPartText" or type 4 "audioFullText".

'Problem: How to determine whether type 3 or type 4? This can only be determined by comparing with the original print?

'4. If one or several documents+fragments are pointed out by all smil <text src,
'and these document contains as child of body the element set allowed in ncc and additional xhtml elements, and n of the smil par elements does not contain an audio child, then
'then dtb is type 5 "textPartAudio"

'Problem: in some textFullAudio dtbs, pars are rendered like that although all audio of the original has actually been narrated.

'5.If one or several documents+fragments are pointed out by all smil <text src,
'and these document contains as child of body the element set allowed in ncc and additional xhtml elements, and none of the smil par elements contains an audio child, then
'then dtb is type6 "textNcc

'End Function
