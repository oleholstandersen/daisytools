Attribute VB_Name = "mInternalSmil"
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

Public Function fncInternalSmilTimeAndId( _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef lCurrentArrayItem As Long, _
    ByRef bolIsMultiVolume As Boolean, _
    ByRef lSmilNumber As Long, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oSmilDom As New MSXML2.DOMDocument40
    oSmilDom.async = False
    oSmilDom.validateOnParse = False
    oSmilDom.resolveExternals = False
    oSmilDom.preserveWhiteSpace = False
    oSmilDom.setProperty "SelectionLanguage", "XPath"
    oSmilDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
    oSmilDom.setProperty "NewParser", True

Dim oNode As IXMLDOMNode
Dim oAttNode As IXMLDOMNode
Dim oNodes As IXMLDOMNodeList
Dim sSmilId As String

  On Error GoTo ErrHandler
  fncInternalSmilTimeAndId = False
  
  sSmilId = CStr(Format(lSmilNumber, "0000"))
  
  If Not fncParseString( _
    objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sDomData, _
    oSmilDom, objOwner _
    ) Then GoTo ErrHandler

    'add id to all pars (if not id is already set)
    Set oNodes = oSmilDom.selectNodes("//body/seq/par")
    If Not fncAddId(oNodes, "rgn_par_" & sSmilId & "_", objOwner, "0000", True) Then GoTo ErrHandler
    
    'objOwner.addlog "id pars done"
    'objOwner.addlog CStr(lCurrentArrayItem)
    
    'add id to all smil audio (regardless of whether set)
    'this should not happen on mvb announce smils, and
    'since they are differently nested than the xpaths below, it is ok like this
    '//REVISIT kolla så att 2.0 och 2.02 MVB smilannounce nestar: //body/seq/audio

    Set oNodes = oSmilDom.selectNodes( _
      "//body/seq/par/seq/audio | //body/seq/par/audio")
    If Not fncAddId(oNodes, "rgn_aud_" & sSmilId & "_", objOwner, "0000", False) Then GoTo ErrHandler
    
    'mg 20030606
    'add id to all smil img, do same as audio above, just mod the xpath
    Set oNodes = oSmilDom.selectNodes( _
      "//body/seq/par/seq/img | //body/seq/par/img")
    If Not fncAddId(oNodes, "rgn_img_" & sSmilId & "_", objOwner, "0000", False) Then GoTo ErrHandler
    
    'objOwner.addlog "id audio done"

    'do totalelapsed time before timeinthissmil
    'since timeinthissmil updates lTotalTimeMs
    'and lTotalTimeMs while still looping through the filesetarray
    'describes time elapsed time in previous smils

    If Not bolIsMultiVolume Then
      If Not fncCreateTotalElapsedTime(oSmilDom, objOwner) Then GoTo ErrHandler
    End If

    'add smil timeinthissmil and seq.dur
    If Not fncCreateTimeInThisSmil(oSmilDom, objOwner) Then GoTo ErrHandler
    
    objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sDomData = oSmilDom.xml

  fncInternalSmilTimeAndId = True
ErrHandler:
  Set oSmilDom = Nothing
  DoEvents
  If Not fncInternalSmilTimeAndId Then objOwner.addlog "<errH in='fncInternalSmilTimeAndId' arrayItem='" & lCurrentArrayItem & "'>fncInternalSmil ErrH" & Err.Number & " : " & Err.Description & "</errH>"
End Function

Public Function fncInternalSmil( _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef lCurrentArrayItem As Long, _
    ByRef bolPreserveBiblioMeta As Boolean, _
    ByRef bolIsMultiVolume As Boolean, _
    ByRef bolFixMultiTextInPar As Boolean, _
    ByRef bolMergeShortPhrases As Boolean, _
    ByRef lClipSpan As Long, _
    ByRef lClipLessThan As Long, _
    ByRef lFirstClipLessThan As Long, _
    ByRef lNextClipLessThan As Long, _
    ByRef lSmilNumber As Long, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
    
Dim oSmilDom As New MSXML2.DOMDocument40
    oSmilDom.async = False
    oSmilDom.validateOnParse = False
    oSmilDom.resolveExternals = False
    oSmilDom.preserveWhiteSpace = False
    oSmilDom.setProperty "SelectionLanguage", "XPath"
    oSmilDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
    oSmilDom.setProperty "NewParser", True
Dim oNode As IXMLDOMNode
Dim oAttNode As IXMLDOMNode
Dim oNodes As IXMLDOMNodeList
Dim sSmilId As String

    On Error GoTo ErrHandler
    fncInternalSmil = False
      
    sSmilId = CStr(Format(lSmilNumber, "0000"))
          
    If Not fncParseString( _
        objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sDomData, _
        oSmilDom, objOwner _
        ) Then GoTo ErrHandler
    
    If Not fncDoNonNccBiblioMetaData( _
        oSmilDom, _
        TYPE_SMIL_1, _
        bolPreserveBiblioMeta, _
        objOwner, _
        lCurrentArrayItem _
        ) Then GoTo ErrHandler
        
    'fix smil layout attrval
    Set oNode = oSmilDom.selectSingleNode("//head/layout/region/@id")
    If Not oNode Is Nothing Then oNode.nodeValue = "txtView"
        
    'remove scheme attributes which are not allowed on meta in smil1dtd
    If Not fncRemoveSchemeAttrs(oSmilDom, objOwner) Then GoTo ErrHandler
    
    'correct first-empty-phrase problem in each par
    If bolMergeShortPhrases Then
      If Not fncCorrectFirstPhraseInPar(oSmilDom, objOwner, lCurrentArrayItem, lClipSpan, lClipLessThan, lFirstClipLessThan, lNextClipLessThan) Then GoTo ErrHandler
    End If
    
    'correct smil clip-times
    If Not fncCorrectLastClipTimeInSmil(oSmilDom, lCurrentArrayItem, objOwner) Then GoTo ErrHandler
    
    If bolFixMultiTextInPar Then
      'fix multiple text children of par
      If Not fncFixMultiTextInPar(oSmilDom, lCurrentArrayItem, lSmilNumber, oNccDom, objOwner) Then GoTo ErrHandler
      'fix first par
      If Not fncFixFirstPar(oSmilDom, lCurrentArrayItem, lSmilNumber, oNccDom, objOwner) Then GoTo ErrHandler
    End If
       
    'fix fragment case
    If Not fncFixFragmentCase(oSmilDom, TYPE_SMIL_1, objOwner) Then GoTo ErrHandler
    
    'mg20039328: added wips rename
    If Not fncRenameNodes("//@endsynch", "endsynch", "endsync", oSmilDom, objOwner) Then GoTo ErrHandler
    
    If Not fncFixAttrValueCase(oSmilDom, objOwner, TYPE_SMIL_1, 0) Then GoTo ErrHandler
    
    'set the modded dom back to array
    objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sDomData = oSmilDom.xml
    
    
    
    
  fncInternalSmil = True
  
ErrHandler:
  Set oSmilDom = Nothing
  DoEvents
  If Not fncInternalSmil Then objOwner.addlog "<errH in='fncInternalSmil' arrayItem='" & lCurrentArrayItem & "'>fncInternalSmil ErrH" & Err.Number & " : " & Err.Description & "</errH>"
End Function

Public Function fncCreateTimeInThisSmil( _
    ByRef oSmilDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim i As Long, k As Long, r As Long, lngTimeCount As Long
Dim oNodeList As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim oNodeMap As IXMLDOMNamedNodeMap
Dim oNodeBase As IXMLDOMElement
Dim oItem As IXMLDOMNode

  On Error GoTo ErrHandler
  fncCreateTimeInThisSmil = False
  lngTimeCount = 0
  Set oNodeList = oSmilDom.getElementsByTagName("audio")
  
  For Each oItem In oNodeList
    Set oNodeMap = oItem.Attributes
    lngTimeCount = lngTimeCount + fncConvertSmilClockVal2Ms(oNodeMap.getNamedItem("clip-end").Text) - fncConvertSmilClockVal2Ms(oNodeMap.getNamedItem("clip-begin").Text)
  Next
  
 '**************************************************
 '* lngTimeCount is now this smil totaltime in ms; *
 '**************************************************
 
 'add value to ncc:timeInThisSmil
 Set oNodeBase = oSmilDom.selectSingleNode("//head")
 If Not oNodeBase Is Nothing Then
   'this elem should not exist but remove if it does
   Set oNode = oNodeBase.selectSingleNode("/meta[name='ncc:timeInThisSmil']")
   If Not oNode Is Nothing Then Set oNode = oNodeBase.removeChild(oNode)
   'then add the whole thing
   If Not fncAppendChild(oNodeBase, "meta", objOwner, , "name", "ncc:timeInThisSmil", "content", fncConvertMS2SmilClockVal(lngTimeCount, 0, False)) Then GoTo ErrHandler
 Else
   objOwner.addlog "<error in='fncCreateTimeInThisSmil'>no head element found in smil</error>"
 End If
  
 'update mother seq dur with same value but other output smilclock syntax
 If Not fncUpdateMotherSeqDur(lngTimeCount, oSmilDom, objOwner) Then GoTo ErrHandler

 'update lTotalTimeMs
  objOwner.lTotalTimeMs = objOwner.lTotalTimeMs + lngTimeCount
  
  fncCreateTimeInThisSmil = True
  
ErrHandler:
  If Not fncCreateTimeInThisSmil Then objOwner.addlog "<errH in='fncCreateTimeInThisSmil'>fncCreateTimeInThisSmil ErrHandler</errH>"
End Function

Public Function fncUpdateMotherSeqDur( _
    ByRef lngTimeCount As Long, _
    ByRef oSmilDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oNode As IXMLDOMNode
Dim oNodeValue As IXMLDOMNode
    
    On Error GoTo ErrHandler
    fncUpdateMotherSeqDur = False
        
    Set oNode = oSmilDom.selectSingleNode("//body/seq/@dur")
    If Not oNode Is Nothing Then
      oNode.nodeValue = fncConvertMS2SmilClockVal(lngTimeCount, 5)
    Else
      objOwner.addlog "<error in='fncUpdateMotherSeqDur'>//body/seq/@dur not found in smil</error>"
    End If
    fncUpdateMotherSeqDur = True
ErrHandler:
  If Not fncUpdateMotherSeqDur Then objOwner.addlog "<errH in='fncUpdateMotherSeqDur'>fncUpdateMotherSeqDur ErrH</errH>"
End Function

Public Function fncCreateTotalElapsedTime( _
    ByRef oSmilDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
    'this function uses the global variable lTotalTimeMs
Dim oNodeBase As IXMLDOMNode
Dim oNode As IXMLDOMNode

  On Error GoTo ErrHandler
  fncCreateTotalElapsedTime = False
    
  Set oNodeBase = oSmilDom.selectSingleNode("//head")
  If Not oNodeBase Is Nothing Then
    'this elem should not exists but remove if it does
    Set oNode = oNodeBase.selectSingleNode("/meta[name='ncc:totalElapsedTime']")
    If Not oNode Is Nothing Then Set oNode = oNodeBase.removeChild(oNode)
    'then add the whole thing
    If Not fncAppendChild(oNodeBase, "meta", objOwner, , "name", "ncc:totalElapsedTime", "content", fncConvertMS2SmilClockVal(objOwner.lTotalTimeMs, 0, False)) Then GoTo ErrHandler
    
  Else
    objOwner.addlog "<error in='fncCreateTotalElapsedTime'>no head element found in smil</error>"
  End If

  fncCreateTotalElapsedTime = True
ErrHandler:
  If Not fncCreateTotalElapsedTime Then objOwner.addlog "<errH in='fncCreateTotalElapsedTime'>fncCreateTotalElapsedTime ErrH</errH>"
End Function

Public Function fncRemoveSchemeAttrs( _
    ByRef oSmilDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oMetaNodes As IXMLDOMNodeList
Dim oMetaNode As IXMLDOMNode

  On Error GoTo ErrHandler
  fncRemoveSchemeAttrs = False
  Set oMetaNodes = oSmilDom.selectNodes("//meta")
  If Not oMetaNodes Is Nothing Then
    For Each oMetaNode In oMetaNodes
      If Not fncRemoveAttribute(oMetaNode, "scheme", objOwner) Then GoTo ErrHandler
    Next
  End If

  fncRemoveSchemeAttrs = True
ErrHandler:
  If Not fncRemoveSchemeAttrs Then objOwner.addlog "<errH in='fncRemoveSchemeAttrs'>fncRemoveSchemeAttrs ErrH</errH>"
End Function
