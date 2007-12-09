Attribute VB_Name = "mInternalSmilFixPar"
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

'fncFixMultiTextInPar and fncFixFirstPar
'are called from mInternalSmil
'before smil meta time calc is done
'since playback length of smilfiles are altered if intersmil fixes are done

'fncFixMultiTextInPar logic:

'For each par in smilfile
'  If par has multiple text children
'    If par is topmost par in smilfile
'      If first text child is not refd by ncc heading
'         create new par, insert text and empty audio ref
'         move to end of previous smil (if existing)
'         rename id of text element
'         correct all references to this element
'      ElseIf first text child is refd by heading
'         create new par, insert text and empty audio ref
'         move above current par
'    ElseIf par is not topmost in smilfile
'         create new par, insert text and empty audio ref
'         move above current par
'    End If 'par is topmost par in smilfile
'  End If 'multiple text children
'Next
'  (recurse until no more multitextchildren in pars of this smil)

'above presumes that
'a) audio id rename is done at later stage
'b) pars created have sessionwide unique id-s or no ids at all

'now no par has multiple text children but first par of smil may be nonheading

'fncFixFirstPar logic:
'For each smilfile
' If first par is not refd by ncc heading
'   If later-sibling par in same smil is refd by heading
'     Move first par to end of previous smilfile (if existing)
'       rename id of text element
'       correct all references to this element
'       if par has id rename that as well
'   Else
'     abort operation
'   End If
' End If 'first par is not refd by ncc heading
'  (recurse until first par is refd by ncc heading)

Public Function fncFixFirstPar( _
    ByRef oSmilDom As MSXML2.DOMDocument40, _
    ByVal lCurrentArrayItem As Long, _
    ByVal lSmilNumber As Long, _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
  
Dim bOneParInThisSmilIsRefdByNccHeading As Boolean
Dim oFirstParNode As IXMLDOMNode
Dim oFirstParTextNode As IXMLDOMNode
Dim oTextNodeList As IXMLDOMNodeList
Dim oTextNode As IXMLDOMNode
Dim oFirstParIdNode As IXMLDOMNode
  
  On Error GoTo ErrHandler
  fncFixFirstPar = False

  Set oFirstParNode = oSmilDom.selectSingleNode("//par[1]")
  If oFirstParNode Is Nothing Then
    objOwner.addlog "<error in='fncFixFirstPar' smilNumber='" & lSmilNumber & "'>oFirstParNode Is Nothing</error>"
    'goto errhandler
    'mg 20030303
    fncFixFirstPar = True
    Exit Function
  End If
  Set oFirstParTextNode = oFirstParNode.selectSingleNode("text")
  If oFirstParTextNode Is Nothing Then
    objOwner.addlog "<error in='fncFixFirstPar' smilNumber='" & lSmilNumber & "'>oFirstParTextNode Is Nothing in fncFixFirstPar</error>"
    'goto errhandler
    'mg 20030303
    fncFixFirstPar = True
    Exit Function
  End If

' If first par is refd by ncc heading then exit
  If fncIsRefdByNccHeading _
    (oFirstParTextNode, lCurrentArrayItem, oNccDom, objOwner) Then
    fncFixFirstPar = True
    Exit Function
  Else
    'mg20030917: test for pars as well
    If fncIsRefdByNccHeading(oFirstParTextNode.parentNode, lCurrentArrayItem, oNccDom, objOwner) Then
      fncFixFirstPar = True
      Exit Function
    Else
      objOwner.addlog "<message smilNumber='" & lSmilNumber & "'>first par in " & objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sFileName & " not referenced by ncc hx</message>"
    End If
  End If

' else continue with mod
' check if any par in this smil is refd by heading
  bOneParInThisSmilIsRefdByNccHeading = False
  Set oTextNodeList = oSmilDom.selectNodes("//par/text[1]")
  If oTextNodeList Is Nothing Then
    objOwner.addlog "<error in='fncFixFirstPar' smilNumber='" & lSmilNumber & "'>oTextNodeList is nothing in fncFixFirstPar</error>"
    fncFixFirstPar = True
    Exit Function
  End If
  
  For Each oTextNode In oTextNodeList
    If fncIsRefdByNccHeading _
     (oTextNode, lCurrentArrayItem, oNccDom, objOwner) Then
     bOneParInThisSmilIsRefdByNccHeading = True
     Exit For
    End If
  Next

' If no later-sibling par in same smil is refd by heading then warn and exit
  If Not bOneParInThisSmilIsRefdByNccHeading Then
    'no par in this smil was refd by ncc heading
    objOwner.addlog "<warning smilNumber='" & lSmilNumber & "'>warning: no par in smilfile " & objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sFileName & " is referenced by ncc heading</warning>"
    fncFixFirstPar = True
    Exit Function
  End If
  
'if we are here, first par is not refd by heading,
'but there is a par below in this smil that is refd by heading
  
' prepare to move first par to end of previous smilfile (if existing)
  If lSmilNumber = 1 Then
    'very unlikely
    objOwner.addlog "<warning>warning: first par of first smil is not referenced by heading</warning>"
    fncFixFirstPar = True
    Exit Function
  End If
' if par has id rename that as well
 Set oFirstParIdNode = oFirstParNode.selectSingleNode("@id")
 If Not oFirstParIdNode Is Nothing Then oFirstParIdNode.Text = "rgn_mvd_" & oFirstParIdNode.Text

' now execute the move
' REVISIT probably need to clone here
  If Not fncMoveParToPreviousSmilEnd(oFirstParNode, lCurrentArrayItem, oNccDom, objOwner, lSmilNumber) Then GoTo ErrHandler
  
'  (recurse until first par is refd by ncc heading)
  If Not fncFixFirstPar(oSmilDom, lCurrentArrayItem, lSmilNumber, oNccDom, objOwner) Then GoTo ErrHandler

  fncFixFirstPar = True
  
ErrHandler:
  If Not fncFixFirstPar Then objOwner.addlog "<errH in='fncFixFirstPar' smilNumber='" & lSmilNumber & "' arrayItem='" & lCurrentArrayItem & "'>fncFixFirstPar ErrH</errH>"
End Function

Public Function fncFixMultiTextInPar( _
    ByRef oSmilDom As MSXML2.DOMDocument40, _
    ByVal lCurrentArrayItem As Long, _
    ByVal lSmilNumber As Long, _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim oParNodeList As IXMLDOMNodeList
Dim oParNode As IXMLDOMNode
Dim oTextNodeList As IXMLDOMNodeList
Dim oTextNode As IXMLDOMNode
Dim oNewParNode As IXMLDOMNode
Dim lParInThisSmil As Long
Dim bDupeFound As Boolean
Dim bCancelFileCheck As Boolean
Dim lDupeCount As Long
  On Error GoTo ErrHandler
  fncFixMultiTextInPar = False
  bDupeFound = False

' For each par in smilfile
  Set oParNodeList = oSmilDom.selectNodes("//par")
  lParInThisSmil = 0
  bCancelFileCheck = False
  
  For Each oParNode In oParNodeList
    'use lParInThisSmil to see if par is the first in the smil
    lParInThisSmil = lParInThisSmil + 1
'   If par has multiple text children
    Set oTextNodeList = oParNode.selectNodes("text")
    If (Not oTextNodeList Is Nothing) And (oTextNodeList.length > 1) Then
      bDupeFound = True
      lDupeCount = lDupeCount + 1
      'objOwner.addlog "<message>multiple text events in par...</message>"
'     create new par
      If Not fncCreateNewParNodeShell(oNewParNode, oSmilDom, objOwner, lDupeCount) Then GoTo ErrHandler
'     set a node to the text elem to be moved
      Set oTextNode = oTextNodeList.Item(0)
'     If par is topmost par in smilfile
      If lParInThisSmil = 1 Then
'       If first text child is not refd by ncc heading
        If Not fncIsRefdByNccHeading(oTextNode, lCurrentArrayItem, oNccDom, objOwner) Then
'          check if a previous smil exists
           If lSmilNumber > 1 Then
'            insert text
             Set oTextNode = oNewParNode.insertBefore(oTextNode, oNewParNode.firstChild)
'            move to end of previous smil, rename id of text element, correct all references to this element
             If Not fncMoveParToPreviousSmilEnd(oNewParNode, lCurrentArrayItem, oNccDom, objOwner, lSmilNumber) Then GoTo ErrHandler
           Else
             objOwner.addlog "<message>first par of first smil has duplicate text events. mod cancelled.</message>"
'            set the boolean so that this file is not recursed
'            otherwise it will loop forever
             bCancelFileCheck = True
           End If 'lSmilNumber > 1
        Else 'If first text child is refd by heading
          If lSmilNumber > 1 Then
'           insert text
            Set oTextNode = oNewParNode.insertBefore(oTextNode, oNewParNode.firstChild)
'           move above current par
            Set oNewParNode = oParNode.parentNode.insertBefore(oNewParNode, oParNode)
            objOwner.addlog "<warning smilNumber='" & lSmilNumber & "'>added silence par (at top) in " & objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sFileName & "</warning>"
          Else
             objOwner.addlog "<message>first par of first smil has duplicate text events. mod cancelled.</message>"
'            set the boolean so that this file is not recursed
'            otherwise it will loop forever
             bCancelFileCheck = True
          End If
        End If 'Not fncIsRefdByNccHeading(oTextNode)
'      ElseIf par is not topmost in smilfile
       ElseIf lParInThisSmil > 1 Then
'          insert text
           Set oTextNode = oNewParNode.insertBefore(oTextNode, oNewParNode.firstChild)
'          move above current par
           Set oNewParNode = oParNode.parentNode.insertBefore(oNewParNode, oParNode)
           objOwner.addlog "<warning smilNumber='" & lSmilNumber & "'>added silence par (not at top) in " & objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sFileName & "</warning>"
       End If 'lParInThisSmil = 1
       
    End If '(Not oTextNodeList Is Nothing) And (oTextNodeList.length > 1)
  Next 'For Each oParNode In oParNodeList
  
  ' (recurse until no more multitextchildren)
  If Not bCancelFileCheck Then
    If (bDupeFound) Then
      If Not fncFixMultiTextInPar(oSmilDom, lCurrentArrayItem, lSmilNumber, oNccDom, objOwner) Then GoTo ErrHandler
    End If
  End If
  
  fncFixMultiTextInPar = True

ErrHandler:
  If Not fncFixMultiTextInPar Then objOwner.addlog "<errH in='fncFixMultiTextInPar' arrayItem='" & lCurrentArrayItem & "'>fncFixMultiTextInPar ErrH</errH>"
End Function

Private Function fncMoveParToPreviousSmilEnd( _
    oNewParNode As IXMLDOMNode, _
    lCurrentArrayItem As Long, _
    oNccDom As MSXML2.DOMDocument40, _
    objOwner As oRegenerator, _
    lSmilNumber As Long _
    ) As Boolean
Dim oPrevSmilDom As New MSXML2.DOMDocument40
    oPrevSmilDom.async = False
    oPrevSmilDom.validateOnParse = False
    oPrevSmilDom.resolveExternals = False
    oPrevSmilDom.preserveWhiteSpace = False
    oPrevSmilDom.setProperty "SelectionLanguage", "XPath"
    oPrevSmilDom.setProperty "NewParser", True
Dim oMotherSeqNode As IXMLDOMNode
Dim oNewParTextNodeId As IXMLDOMNode
Dim lPreviousSmilArrayItem As Long
Dim sOldUri As String, sNewUri As String

'inserts parnode at end of previous smil,
'renames id of text element
'correct all references to this element

  On Error GoTo ErrHandler
  fncMoveParToPreviousSmilEnd = False
  
  If Not objOwner.objFileSetHandler.fncGetPreviousItemOfType(TYPE_SMIL_1, lPreviousSmilArrayItem, lCurrentArrayItem) Then GoTo ErrHandler
  
  If Not lPreviousSmilArrayItem = -1 Then
    'a previous smil was found
    'do some dblcheck before parse
    If lPreviousSmilArrayItem >= lCurrentArrayItem Then
      objOwner.addlog "<error in='fncMoveParToPreviousSmilEnd'>error: lPreviousSmilArrayItem >= lCurrentArrayItem in fncFixMultiTextInPar</error>"
      GoTo ErrHandler
    End If
    If objOwner.objFileSetHandler.aOutFileSet(lPreviousSmilArrayItem).eType <> TYPE_SMIL_1 Then
      objOwner.addlog "<error in='fncMoveParToPreviousSmilEnd'>error: aOutFileSet(lPreviousSmilArrayItem).eType <> smil_1</error>"
      GoTo ErrHandler
    End If
         
    'parse
    If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(lPreviousSmilArrayItem).sDomData, oPrevSmilDom, objOwner) Then GoTo ErrHandler
    Set oMotherSeqNode = oPrevSmilDom.selectSingleNode("//body/seq")
    If Not oMotherSeqNode Is Nothing Then
      'add the par as last child of motherseq
      Set oNewParNode = oMotherSeqNode.appendChild(oNewParNode)
      'prepare to mod id of and refs to new par text child
      Set oNewParTextNodeId = oNewParNode.selectSingleNode("text/@id")
      If oNewParTextNodeId Is Nothing Then GoTo ErrHandler
      sOldUri = LCase$(objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sFileName & "#" & oNewParTextNodeId.Text)
      'modify the id of the textnode in this par
      oNewParTextNodeId.Text = "rgn_mv_" & oNewParTextNodeId.Text
      sNewUri = objOwner.objFileSetHandler.aOutFileSet(lPreviousSmilArrayItem).sFileName & "#" & oNewParTextNodeId.Text
      'modify all references to this textnode
      If Not fncChangeSmilUriReference(sOldUri, sNewUri, oNccDom, objOwner) Then GoTo ErrHandler
      'save the modified prevsmil back to array
      objOwner.objFileSetHandler.aOutFileSet(lPreviousSmilArrayItem).sDomData = oPrevSmilDom.xml
      objOwner.addlog "<message smilNumber='" & lSmilNumber & "'>added par in " & objOwner.objFileSetHandler.aOutFileSet(lPreviousSmilArrayItem).sFileName & " moved from " & objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sFileName & "</message>"
    Else
      objOwner.addlog "<error in='fncFixMultiTextInPar' smilNumber='" & lSmilNumber & "'>oMotherSeqNode Is Nothing in in fncFixMultiTextInPar</error>"
      GoTo ErrHandler
    End If 'Not oMotherSeqNode Is Nothing
  Else
    'no previous smil was found
    objOwner.addlog "<message in='fncFixMultiTextInPar' smilNumber='" & lSmilNumber & "'>tried to move par from " & objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sFileName _
    & " but no previous smil found.</message>"
    GoTo ErrHandler
  End If 'Not lPreviousSmilArrayItem = -1

  fncMoveParToPreviousSmilEnd = True
  
ErrHandler:
  If Not fncMoveParToPreviousSmilEnd Then objOwner.addlog "<errH in='fncMoveParToPreviousSmilEnd' arrayItem='" & lCurrentArrayItem & "'>fncMoveParToPreviousSmilEnd ErrH</errH>"
End Function

Private Function fncIsRefdByNccHeading( _
    ByRef oTextNode As IXMLDOMNode, _
    ByVal lCurrentArrayItem As Long, _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim sTextNodeId As String
Dim sSmilFileName As String
Dim sTextNodeUri As String
Dim oNccAnchorNodes As IXMLDOMNodeList
Dim oNccAnchorNode As IXMLDOMNode
Dim oHrefNode As IXMLDOMNode
Dim oNccAnchorParent As IXMLDOMNode

    On Error GoTo ErrHandler
    'create the URI that references this par in ncc
    sTextNodeId = oTextNode.selectSingleNode("@id").Text
    If sTextNodeId = "" Then GoTo ErrHandler
    sSmilFileName = objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sFileName
    sTextNodeUri = LCase$(sSmilFileName & "#" & sTextNodeId)
    
    'find the a href in ncc that refs
    Set oNccAnchorNodes = oNccDom.selectNodes _
       ("//h1/a" & _
      "| //h2/a" & _
      "| //h3/a" & _
      "| //h4/a" & _
      "| //h5/a" & _
      "| //h6/a" & _
      "| //span/a" & _
      "| //div/a")
    For Each oNccAnchorNode In oNccAnchorNodes
      Set oHrefNode = oNccAnchorNode.selectSingleNode("@href")
      If Not oHrefNode Is Nothing Then
        If LCase$(oHrefNode.nodeTypedValue) = sTextNodeUri Then
          'we found the ncc href that references our smil text node
          'get the parent and check its name
          Set oNccAnchorParent = oNccAnchorNode.parentNode
          If oNccAnchorParent.nodeName = "h1" Or _
             oNccAnchorParent.nodeName = "h2" Or _
             oNccAnchorParent.nodeName = "h3" Or _
             oNccAnchorParent.nodeName = "h4" Or _
             oNccAnchorParent.nodeName = "h5" Or _
             oNccAnchorParent.nodeName = "h6" Then
            fncIsRefdByNccHeading = True
            Exit Function
          Else
            fncIsRefdByNccHeading = False
            Exit Function
          End If
        End If 'LCase$(oHrefNode.Text) =
      Else
        objOwner.addlog "<warning in='fncIsRefdByNccHeading'>warning: ncc anchor without href attribute</warning>"
      End If 'Not oHrefNode Is Nothing
    Next 'For Each oNccAnchorNode In oNccAnchorNodes
    
    'if we came here, the smil text node was not referenced by the ncc at all
    fncIsRefdByNccHeading = False
    Exit Function
ErrHandler:
 objOwner.addlog "<errH in='fncIsRefdByNccHeading' arrayItem='" & lCurrentArrayItem & "'>fncIsRefdByNccHeading ErrH</errH>"
End Function
    
Private Function fncCreateNewParNodeShell( _
    ByRef oNewParNode As IXMLDOMNode, _
    ByRef oSmilDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator, _
    ByRef lCount As Long _
    ) As Boolean
Dim oSeqNode As IXMLDOMNode

  On Error GoTo ErrHandler
  fncCreateNewParNodeShell = False
  
  Set oNewParNode = oSmilDom.createNode(NODE_ELEMENT, "par", "")
  If Not fncAppendAttribute(oNewParNode, "endsync", "last", objOwner) Then GoTo ErrHandler
  If Not fncAppendChild(oNewParNode, "seq", objOwner) Then GoTo ErrHandler
  Set oSeqNode = oNewParNode.selectSingleNode("seq")
  If Not fncAppendChild _
    (oSeqNode, "audio", objOwner, , "src", objOwner.sEmptyMp3Filename, "clip-begin", _
    "npt=0.000s", "clip-end", "npt=0.750s", "id", "rgn_ins_" & CStr(Format(lCount, "0000"))) Then GoTo ErrHandler
  'insert the audio file into OutArray
  If Not objOwner.objFileSetHandler.fncIsObjectInOutputArray(objOwner.sEmptyMp3Filename) Then
    If Not objOwner.objFileSetHandler.fncAddObjectToOutputArray(TYPE_SMIL_AUDIO_INSERTED, objOwner.sEmptyMp3Filename, "") Then GoTo ErrHandler
  End If
    
  fncCreateNewParNodeShell = True
ErrHandler:
  If Not fncCreateNewParNodeShell Then objOwner.addlog "<errH in='fncCreateNewParNodeShell'>fncCreateNewParNodeShell errH</errH>"
End Function

'Public Function fncFixMultiTextInParOld( _
'    ByRef oSmilDom As MSXML2.DOMDocument40, _
'    ByVal lCurrentArrayItem As Long _
'    ) As Boolean
'  '*********************************************************
'  '* this function fixes multiple text events inside par
'  '* by removing all but the last text in the series
'  '* by creating a preceeding par sibling for the moved text event
'  '* pointing to a 0.75 sec audio file of silence
'  '* extended logic described above in this module
'  '*********************************************************
'Dim oPrevSmilDom As New MSXML2.DOMDocument40
'    oPrevSmilDom.async = False
'    oPrevSmilDom.validateOnParse = False
'    oPrevSmilDom.resolveExternals = False
'    oPrevSmilDom.preserveWhiteSpace = False
'    oPrevSmilDom.setProperty "SelectionLanguage", "XPath"
'    oPrevSmilDom.setProperty "NewParser", True
'Dim oParNodeList As IXMLDOMNodeList
'Dim oParNode As IXMLDOMNode
'Dim oTextNodeList As IXMLDOMNodeList
'Dim oTextNode As IXMLDOMNode
'Dim oNewParNode As IXMLDOMNode
'Dim bDupeFound As Boolean
'Dim bNoPreviousSmil As Boolean
'Dim lParInThisSmil As Long
'Dim lPreviousSmilArrayItem As Long
'Dim oMotherSeqNode As IXMLDOMNode
'
'  On Error GoTo ErrHandler
'  fncFixMultiTextInPar = False
'
'  bDupeFound = False
'  bNoPreviousSmil = False
'  Set oParNodeList = oSmilDom.selectNodes("//par")
'  lParInThisSmil = 0
'  For Each oParNode In oParNodeList
'    lParInThisSmil = lParInThisSmil + 1
'    Set oTextNodeList = oParNode.selectNodes("text")
'    If (Not oTextNodeList Is Nothing) And (oTextNodeList.length > 1) Then
'      'there are multiple text elements in this par
'      bDupeFound = True
'      'create a new parnode
'      If Not fncCreateNewParNodeShell(oNewParNode, oSmilDom) Then GoTo ErrHandler
'      'set the first text elem
'      Set oTextNode = oTextNodeList.Item(0)
'      'insert the new par:
'      'if orig par is the first par in smilfile
'      'then insert shall be done in previous smilfile
'      'if a previous smilfile exists, if not, this is the title par
'      'then no move at all shall be done
'      If lParInThisSmil = 1 Then
'        'dupe text is in first par of this smilfile
'        If Not fncGetPreviousItemOfType( _
'          smil_1, lPreviousSmilArrayItem, lCurrentArrayItem) Then GoTo ErrHandler
'        If Not lPreviousSmilArrayItem = -1 Then
'         'a previous smil was found
'         'do some dblcheck before parse
'         If lPreviousSmilArrayItem >= lCurrentArrayItem Then
'           objowner.addlog "error: lPreviousSmilArrayItem >= lCurrentArrayItem in fncFixMultiTextInPar"
'           GoTo ErrHandler
'         End If
'         If aOutFileSet(lPreviousSmilArrayItem).eType <> smil_1 Then
'           objowner.addlog "error: aOutFileSet(lPreviousSmilArrayItem).eType <> smil_1"
'           GoTo ErrHandler
'         End If
'
'         'parse
'         If Not fncParseString(aOutFileSet(lPreviousSmilArrayItem).sDomData, oPrevSmilDom) Then GoTo ErrHandler
'         Set oMotherSeqNode = oPrevSmilDom.selectSingleNode("//body/seq")
'         If Not oMotherSeqNode Is Nothing Then
'           'put the textnode in the par shell
'           Set oTextNode = oNewParNode.insertBefore(oTextNode, oNewParNode.firstChild)
'           Set oNewParNode = oMotherSeqNode.appendChild(oNewParNode)
'           DoEvents
'           aOutFileSet(lPreviousSmilArrayItem).sDomData = oPrevSmilDom.xml
'           objowner.addlog "duplicate text elements found; added new silence par in " & aOutFileSet(lPreviousSmilArrayItem).sFileName & " moved from " & aOutFileSet(lCurrentArrayItem).sFileName
'         Else
'           objowner.addlog "oMotherSeqNode Is Nothing in in fncFixMultiTextInPar"
'           GoTo ErrHandler
'         End If
'        Else
'          'no previous smil was found
'          bNoPreviousSmil = True
'          objowner.addlog "tried to move par from " & aOutFileSet(lCurrentArrayItem).sFileName _
'          & " but no previous smil found. Dupe text child of par will remain."
'        End If 'lParinThisSmil = 1
'      Else
'        'dupe text is not in first par
'        'put the textnode in the par shell
'        Set oTextNode = oNewParNode.insertBefore(oTextNode, oNewParNode.firstChild)
'        Set oNewParNode = oParNode.parentNode.insertBefore(oNewParNode, oParNode)
'        objowner.addlog "duplicate text elements found; added new silence par in " & aOutFileSet(lCurrentArrayItem).sFileName
'      End If
'
'
'    Else
'      'this par contained only one text element child - do nothing
'    End If '(Not oTextNodeList Is Nothing) And (oTextNodeList.length > 1)
'  Next
'
'  'make recursive call (there may be more than two text elems in the same par)
'  If (bDupeFound) And (Not bNoPreviousSmil) Then
'    If Not fncFixMultiTextInPar(oSmilDom, lCurrentArrayItem) Then GoTo ErrHandler
'  End If
'
'  fncFixMultiTextInPar = True
'ErrHandler:
'  If Not fncFixMultiTextInPar Then objowner.addlog "fncFixMultiTextInPar ErrH"
'End Function
