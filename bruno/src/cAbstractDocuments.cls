VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAbstractDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private aAbstractDocuments() As cAbstractDocument
Public lAbstractDocumentCount As Long

Private oAbstractSmil As cAbstractDocument
Private oAbstractNavigation As cAbstractDocument

Private oSmilParIdGetter As cIdGetter
Private oSmilSeqIdGetter As cIdGetter
Private oContentIdGetter As cIdGetter
Private oSmilTextIdGetter As cIdGetter

Private oRemoveNodes As IXMLDOMNodeList
Private oSyncOmitNodes As IXMLDOMNodeList
Private oSyncForceNodes As IXMLDOMNodeList
Private oSmilRefOmitNodes As IXMLDOMNodeList
Private oSmilSequenceNodes As IXMLDOMNodeList
Private oSmilSequenceUserEscapeNodes As IXMLDOMNodeList
Private oSmilCustomTestNodes As IXMLDOMNodeList
Private oNavIncludeNodes As IXMLDOMNodeList
Private oNavIncludeNodesExplicitNavLabel As IXMLDOMNodeList
Private oMixedContentHandlerNodes As IXMLDOMNodeList

Private oSmilSplitNodes As IXMLDOMNodeList 'for optimization

Private lTotalSyncPoints As Long

Private lPhysicalSmilFileCount As Long

Private oLastAbstractSmilContainerAdded As IXMLDOMNode
Private oLastNodeAppended As IXMLDOMNode
Private oAbstractSmilSequences As cAbstractSmilSequences

Private Sub Class_Initialize()
  'debug.Print "cAbstractDocuments.Initialize"
  Set oContentIdGetter = New cIdGetter
  Set oSmilParIdGetter = New cIdGetter
  Set oSmilSeqIdGetter = New cIdGetter
  Set oSmilTextIdGetter = New cIdGetter
  lTotalSyncPoints = 0
  lPhysicalSmilFileCount = 1
End Sub

Public Function fncInstantiate() As Boolean
Dim oAbstractContentDoc As cAbstractDocument
'use oBruno.oInputDocuments
'create the three abstract documents to use
Dim i As Long
Dim oContentDocNode As IXMLDOMNode
Dim oContentDocNodeId As IXMLDOMNode
Dim oRemoveNode As IXMLDOMNode
Dim oBodyNodes As IXMLDOMNodeList
Dim oNavNode As IXMLDOMNode
Dim oAbstractSmil As cAbstractDocument
Dim oAbstractNavigation As cAbstractDocument

'for DTBuserEscape:
Dim oAbstractSmilUserEscapeContainers As IXMLDOMNodeList
Dim oAbstractSmilUserEscapeContainer As IXMLDOMNode
Dim oAbstractSmilUserEscapeContainerEndAttribute As IXMLDOMNode
Dim oTimeContainersInSeq As IXMLDOMNodeList
Dim oLastTimeContainerInSeq As IXMLDOMNode

 On Error GoTo errh
 
 fncInstantiate = False
 
 For i = 0 To oBruno.oInputDocuments.InputDocumentCount - 1
 'loop through input documents
   ReDim Preserve aAbstractDocuments(lAbstractDocumentCount)
   Set aAbstractDocuments(lAbstractDocumentCount) = New cAbstractDocument
   With aAbstractDocuments(lAbstractDocumentCount)
     
     .lAbstractType = TYPE_ABSTRACT_CONTENTDOC
     .fncSetFileName (oBruno.oInputDocuments.InputDocument(i).sFileName)
     .oDom.loadXML oBruno.oInputDocuments.InputDocument(i).oDom.xml
     .lInputContentDocSource = i
      
      'remove nodes as per Driver
     Set oRemoveNodes = .oDom.selectNodes(oBruno.oDriver.sXpathRemove)
     oBruno.oCmn.oDomCmn.fncRemoveNodesInNodeList oRemoveNodes
         
     'create persistent nodelists for matching while iterating the tree later
     
     'set nodelist relative to startElement
     Dim oDriverStartElement As IXMLDOMElement
     Set oDriverStartElement = .oDom.selectSingleNode(oBruno.oDriver.sXpathStartElem)
     If oDriverStartElement Is Nothing Then GoTo errh
     
     Set oSyncOmitNodes = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathSyncOmit)
     Set oSyncForceNodes = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathSyncForce)
     Set oSmilRefOmitNodes = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathSmilRefOmit)
     Set oSmilSequenceNodes = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathSmilSequence)
     Set oSmilSequenceUserEscapeNodes = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathSmilSequenceUserEscape)
     Set oSmilCustomTestNodes = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathSmilCustomTest)
     Set oNavIncludeNodes = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathNavInclude)
     Set oNavIncludeNodesExplicitNavLabel = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathNavIncludeExplicitNavLabel)
     Set oMixedContentHandlerNodes = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathMixedContentHandler)
     'for optim
     Set oSmilSplitNodes = oDriverStartElement.selectNodes(oBruno.oDriver.sXpathSmilSplit)
     
     'create abstract smilspine, add it to array later
     If oAbstractSmil Is Nothing Then
       'do the test above since this may be > first iterat through content doc list
       Set oAbstractSmil = New cAbstractDocument
       oAbstractSmil.lAbstractType = TYPE_ABSTRACT_SMIL
       oAbstractSmil.oDom.loadXML ("<smil></smil>")
     End If
     
     'create sequences object for handling smil nesting
     Set oAbstractSmilSequences = New cAbstractSmilSequences
     
     'create abstract nav, add it to array later
     If oAbstractNavigation Is Nothing Then
       'do the test above since this may be > first iterat through content doc list
       Set oAbstractNavigation = New cAbstractDocument
       oAbstractNavigation.lAbstractType = TYPE_ABSTRACT_NAVIGATION
       oAbstractNavigation.oDom.loadXML ("<navigation></navigation>")
     End If

     '*****************************************************************
     'first walk and normalize mixed content
     
     If (oBruno.oDriver.bMixedContentHandlerActive) Then
       fncTreeWalkerToNormalizer .oDom.selectSingleNode(oBruno.oDriver.sXpathStartElem)
     End If
     
     '*****************************************************************
     'then walk again append to abstraction
     
     fncTreeWalkerToAbstraction .oDom.selectSingleNode(oBruno.oDriver.sXpathStartElem), oAbstractSmil, .sFileName
     
     '*****************************************************************
     
     
     
     'now add all members of this contentdoc that are navinclude into abstract navigation
     For Each oNavNode In oNavIncludeNodes
       'if its not in syncomit, which would mean it will have no abstract smil item
       If Not oBruno.oCmn.oDomCmn.fncIsInNodeList(oNavNode, oSyncOmitNodes) Then
        If Not oNavNode.selectSingleNode("@smilref") Is Nothing Then
        'this node has not become a part of smil presentation for some reason
          Dim oCloneNode As IXMLDOMNode
          Set oCloneNode = oNavNode.cloneNode(True)
          
          'remove all childelems from the cloned node
          '20070119 here is the gita bug
          'try to comment out below to let later renderer deal with it
          'oBruno.oCmn.oDomCmn.fncRemoveAllChildElements oCloneNode
          
          If oBruno.oCmn.oDomCmn.fncIsInNodeList(oNavNode, oNavIncludeNodesExplicitNavLabel) Then
            'this node has an explicit navLabel set in driver
            'get the xpath statement to get the child text value from
            Dim sNavLabelXpathStatement As String
            sNavLabelXpathStatement = oBruno.oDriver.fncGetXpathForNavLabel(oNavNode)
            'get the childnode from oNavNode; (this might be a text() setter!)
            Dim oNavNodeChild As IXMLDOMNode
            Dim oBrunoNavLabelAttr As IXMLDOMAttribute
            Set oNavNodeChild = oNavNode.selectSingleNode(sNavLabelXpathStatement)
            If Not oNavNodeChild Is Nothing Then
              'append a special attribute to clonenode, to use later in ooutputdocuments
              If oNavNodeChild.Text <> "" Then
                Set oBrunoNavLabelAttr = oAbstractNavigation.oDom.createAttribute("bruno_navLabel")
                oBrunoNavLabelAttr.Value = oNavNodeChild.Text
                Set oBrunoNavLabelAttr = oCloneNode.Attributes.setNamedItem(oBrunoNavLabelAttr)
              Else
                'for this particular oNavNode, the xpath returned a node without text
              End If
            Else
             'for this particular oNavNode, the xpath didnt return anything
            End If
          Else
            'this node should be added as is
          End If
          oAbstractNavigation.oDom.documentElement.appendChild oCloneNode
        Else
          Debug.Print "navnode without smilref omitted from inclusion in navdoc"
        End If 'oNavNode.selectSingleNode("@smilref")
       End If
        'Debug.Print oNavNode.xml
     Next
   End With
  lAbstractDocumentCount = lAbstractDocumentCount + 1
  'kill smilseq since it contains contentdoc info
  Set oAbstractSmilSequences = Nothing
 Next
 
 'if user-escape events exists, add the id of the last descendant
 'attribute on seq before this is "end="DTBuserEscape;"
 Set oAbstractSmilUserEscapeContainers = oAbstractSmil.oDom.selectNodes("//seq[@end='DTBuserEscape;']")
 If Not oAbstractSmilUserEscapeContainers Is Nothing Then
   For Each oAbstractSmilUserEscapeContainer In oAbstractSmilUserEscapeContainers
     Set oAbstractSmilUserEscapeContainerEndAttribute = oAbstractSmilUserEscapeContainer.selectSingleNode("@end")
     'get the last time container id value
     'Set oTimeContainersInSeq = oAbstractSmilUserEscapeContainer.selectNodes(".//seq | .//par")
     Set oTimeContainersInSeq = oAbstractSmilUserEscapeContainer.selectNodes("./seq | ./par")
     Set oLastTimeContainerInSeq = oTimeContainersInSeq.Item(oTimeContainersInSeq.length - 1)
     oAbstractSmilUserEscapeContainerEndAttribute.Text = oAbstractSmilUserEscapeContainerEndAttribute.Text & oLastTimeContainerInSeq.selectSingleNode("@id").Text & ".end"
   Next
 End If
 
 'now all content docs have been added and mangled,
 'abs smil and nav are free floating objects, so:
 
 'add abstractsmilspine into array:
 ReDim Preserve aAbstractDocuments(lAbstractDocumentCount)
 Set aAbstractDocuments(lAbstractDocumentCount) = oAbstractSmil
 lAbstractDocumentCount = lAbstractDocumentCount + 1
     
 'add abstractnav into array:
 ReDim Preserve aAbstractDocuments(lAbstractDocumentCount)
 Set aAbstractDocuments(lAbstractDocumentCount) = oAbstractNavigation
 lAbstractDocumentCount = lAbstractDocumentCount + 1
 
'  now the abstract fileset is complete and all sits in aAbstractDocuments() array
'  below for debug:
'  fncSaveAbstractOutput
'  Stop
  
    
  fncInstantiate = True
  Exit Function
errh:
  frmMain.fncAddMessage "error in oAbstractDocuments"
End Function

Private Function fncTreeWalkerToNormalizer(ByRef oInElement As IXMLDOMElement)
'a treewalker that spits only elements
Dim oInElementChild As IXMLDOMNode
Dim oInElementChildNodes As IXMLDOMNodeList
 
 fncElementContentNormalizer oInElement
 If oInElement.hasChildNodes Then
   Set oInElementChildNodes = oInElement.selectNodes("*")
   For Each oInElementChild In oInElementChildNodes
     'DoEvents '''
     fncTreeWalkerToNormalizer oInElementChild
   Next
 End If
End Function

Private Function fncTreeWalkerToAbstraction(ByRef oInElement As IXMLDOMElement, oAbstractSmil As cAbstractDocument, sContentDocFileName As String)
'a treewalker that spits only elements
Dim oInElementChild As IXMLDOMNode
Dim oInElementChildNodes As IXMLDOMNodeList
 
 fncElementToAbstraction oInElement, oAbstractSmil, sContentDocFileName
 If oInElement.hasChildNodes Then
   Set oInElementChildNodes = oInElement.selectNodes("*")
   For Each oInElementChild In oInElementChildNodes
     'DoEvents '''
     fncTreeWalkerToAbstraction oInElementChild, oAbstractSmil, sContentDocFileName
   Next
 End If
End Function

Private Function fncElementContentNormalizer(ByRef oElement As IXMLDOMElement)
Dim oNode As IXMLDOMNode
Dim oNewElement As IXMLDOMNode
Dim oElementTextChildNodes As IXMLDOMNodeList

  If oBruno.oCmn.oDomCmn.fncHasMixedContentChildren(oElement) Then
    
    If fncAllElemSiblingsAreSyncOmit(oElement.childNodes) Then
       'all siblings are text or syncomit - do nothing
    ElseIf fncAllElemSiblingsAreSyncInclude(oElement.childNodes) Then
       'all textnodes should be splitted
       If oBruno.oCmn.oDomCmn.fncIsInNodeList(oElement, oMixedContentHandlerNodes) Then
         Set oElementTextChildNodes = oElement.selectNodes("text()")
         For Each oNode In oElementTextChildNodes
           Set oNewElement = fncMixedContentElemAppend(oNode)
         Next oNode
       End If
    Else
       'worst case: siblings are mixed text + syncinclude + syncomit elems
        If oBruno.oCmn.oDomCmn.fncIsInNodeList(oElement, oMixedContentHandlerNodes) Then
          Set oElementTextChildNodes = oElement.selectNodes("text()")
           For Each oNode In oElementTextChildNodes
             'check if previous or next sibling if intext is syncomit
             If oBruno.oCmn.oDomCmn.fncIsInNodeList(oNode.previousSibling, oSyncOmitNodes) _
               Or oBruno.oCmn.oDomCmn.fncIsInNodeList(oNode.nextSibling, oSyncOmitNodes) Then
                 'at least one of closest siblings is in syncomit
                 If Not fncAllElemSiblingsAreSyncOmit(oNode.parentNode.childNodes) Then
                   'add this check since a previous textnode iterat may have changed the surroundings
                   Set oNewElement = fncMixedContentElemAppendIncludeOmitSiblings(oNode)
                 End If
             Else
               'closest sibling is not in syncomit
               Set oNewElement = fncMixedContentElemAppend(oNode)
             End If
          Next oNode
        End If

    End If
  
  End If 'fncHasMixedContentChildren(oElement)

End Function

Private Function fncElementToAbstraction(ByRef oElement As IXMLDOMElement, ByRef oAbstractSmil As cAbstractDocument, ByRef sContentDocFileName As String)
 
 'we get here;
 'empty elems and others are in syncomit
 'we dont know if mixed content handling has been done
  
 'test whether the node should be added to abstraction
 With oBruno.oCmn.oDomCmn
   
   '20070910: if parent in syncforce, bail
   If (.fncIsInNodeList(oElement.parentNode, oSyncForceNodes)) Then
     Exit Function
   End If
      
   If (Not .fncIsInNodeList(oElement, oSyncForceNodes)) Then
   'if not in syncforce then do other tests
      If (Not .fncIsInNodeList(oElement, oNavIncludeNodes)) Then
      'if not in navinclude then do other tests
          If (.fncIsInNodeList(oElement, oSyncOmitNodes)) Then
            '20070910 workaround for a bug where for example
            '<p>
            '[6] one
            '<em>SYNC POINT</em>
            '<span class="page-normal" id="page-162">ONE SYNC POINT</span>
            '<em id="i001">ONE SYNC POINT</em>
            '</p>
            '... id=i001 above becomes unsynchronized.
            'This is fixable in the xpath driver, but very nasty to express and maintain
            'So: if current has a text sibling or a syncomit elem sibling or parent has been synced, then exit, else continue.
            If (Not fncOmitOmit(oElement)) Then
              Exit Function
            End If
          End If
          'if syncomit - exit (note: navinclude and syncforce wins over syncomit)
          If (.fncHasIgnorableWhiteSpaceOnly(oElement)) Then
          'it is whitespace only - should we exit?
            If ((Not .fncIsInNodeList(oElement, oSmilSequenceNodes)) _
              And (Not .fncIsInNodeList(oElement, oSmilCustomTestNodes))) Then
              'its not a sequence node, nor is it a customtestnode: exit
              Exit Function
            Else
              'its a sequence node or a customtestnode, continue
            End If
          Else
            'it didnt have ignorable whitespace only, continue
          End If
     Else
       'this node is in navinclude, continue
     End If 'Not .fncIsInNodeList(oElement, oNavIncludeNodes)
   Else
     'this node is in syncforce, continue
   End If 'Not .fncIsInNodeList(oElement, oSyncForceNodes)
 End With
  
  
  'if we are here, then at least one text node, and possibly one or more elem children
  'or elem children only (in which case it should become a seq, else something is wrong)
  
  ' if we get here, the text or element node is one that should be part of abstraction,
  ' else exit before
 
  'set id if not already present in contentdoc
  If oElement.selectSingleNode("@id") Is Nothing Then
    oBruno.oCmn.oDomCmn.fncAppendAttribute oElement, "id", "cn" & oContentIdGetter.fncGetId
  End If
    
  'append the node to abstraction, if not already done
  If oElement.selectSingleNode("@smilref") Is Nothing Then
    fncAppendNodeToAbstraction oElement, oAbstractSmil, sContentDocFileName
    lTotalSyncPoints = lTotalSyncPoints + 1
  Else
    Debug.Print "done before: " & oElement.xml
  End If
End Function

Private Function fncOmitOmit(oElement As IXMLDOMElement) As Boolean
  'test whether to override the assumed omit of inparam element
  'if the syncomit should be overridden to a syncinclude, return true
  'return false if current has a text sibling or a syncomit elem sibling, or if nodes parent has been synced
  
  If Not oElement.parentNode.selectSingleNode("@smilref") Is Nothing Then
    fncOmitOmit = False
    Exit Function
  End If
  
  Dim oPreviousNode As IXMLDOMNode
  Dim oNextNode As IXMLDOMNode
  
  Set oPreviousNode = oElement.previousSibling
  Set oNextNode = oElement.nextSibling
  
  If (Not oPreviousNode Is Nothing) Then
    If (oPreviousNode.nodeType = NODE_TEXT) Then
        fncOmitOmit = False
        Exit Function
    End If
    
    If (oPreviousNode.nodeType = NODE_ELEMENT) Then
        If (oBruno.oCmn.oDomCmn.fncIsInNodeList(oPreviousNode, oSyncOmitNodes)) Then
            fncOmitOmit = False
            Exit Function
        End If
    End If
  End If

  If (Not oNextNode Is Nothing) Then
    If (oNextNode.nodeType = NODE_TEXT) Then
        fncOmitOmit = False
        Exit Function
    End If
    
    If (oNextNode.nodeType = NODE_ELEMENT) Then
        If (oBruno.oCmn.oDomCmn.fncIsInNodeList(oNextNode, oSyncOmitNodes)) Then
            fncOmitOmit = False
            Exit Function
        End If
    End If
  End If

  fncOmitOmit = True

End Function

Public Function AbstractSmil() As cAbstractDocument
Dim i As Long
  
  For i = 0 To lAbstractDocumentCount
    If aAbstractDocuments(i).lAbstractType = TYPE_ABSTRACT_SMIL Then
      Set AbstractSmil = aAbstractDocuments(i)
      Exit For
    End If
  Next i
 
End Function

Private Function fncMixedContentElemAppend(ByRef oTextNode As IXMLDOMText) As IXMLDOMElement
'gets a textnode as input
'which has at least one element sibling
'returns a newly created parent of the text
Dim oNewWrapperElem As IXMLDOMElement
Dim oNewWrapperElemAtt As IXMLDOMAttribute
Dim oMovedTextNode As IXMLDOMNode

  Set oNewWrapperElem = oTextNode.ownerDocument.createElement(oBruno.oDriver.sItmtElemName)
  Set oNewWrapperElemAtt = oTextNode.ownerDocument.createAttribute(oBruno.oDriver.sItmtAttName)
  oNewWrapperElemAtt.Text = oBruno.oDriver.sItmtAttValue
  Set oNewWrapperElemAtt = oNewWrapperElem.Attributes.setNamedItem(oNewWrapperElemAtt)
  'add wrapper
  Set oNewWrapperElem = oTextNode.parentNode.insertBefore(oNewWrapperElem, oTextNode)
  'move textnode into wrapper
  Set oMovedTextNode = oNewWrapperElem.appendChild(oTextNode)
  Set fncMixedContentElemAppend = oNewWrapperElem
End Function


Private Function fncAppendNodeToAbstraction( _
    ByRef oNodeToAppend As IXMLDOMNode, _
    ByRef oAbstractSmil As cAbstractDocument, _
    ByRef sContentDocFileName As String) _
    As Boolean
Dim oAbstractTimeContainerToAdd As IXMLDOMNode
Dim oPointForTimeContainer As IXMLDOMNode
Dim bNewPhysicalSmilFile As Boolean
    'Debug.Print oNodeToAppend.nodeName
    
    If Not oAbstractSmilSequences.initialized Then
      oAbstractSmilSequences.initialize oNodeToAppend.ownerDocument.documentElement, oAbstractSmil
      'oAbstractSmilSequences array now has a rootsequence pointing to abstract smil documentelemement
      'and identifies all contentdoc elems
    End If
    
    'check if theoretic physical smilfilecount needs to be upped (optim)
    If oBruno.oCmn.oDomCmn.fncIsInNodeList(oNodeToAppend, oSmilSplitNodes) Then
      '20040727: only up count unless this is the very first synced node of the contentdoc
      '(since lPhysicalSmilFileCount starts at 1)
      If lTotalSyncPoints > 0 Then
        lPhysicalSmilFileCount = lPhysicalSmilFileCount + 1
        bNewPhysicalSmilFile = True
      End If
    End If
    
    'create the basic time container element
    Set oAbstractTimeContainerToAdd = fncCreateAbstractContainer(oNodeToAppend, sContentDocFileName, oAbstractSmil, bNewPhysicalSmilFile)
    
    'if its a sequence, create new sequence object in oAbstractSmilSequences
    If oAbstractTimeContainerToAdd.nodeName = "seq" Then
      oAbstractSmilSequences.AddSequence oNodeToAppend, oAbstractTimeContainerToAdd
    End If
        
        
    If oBruno.oCmn.oDomCmn.fncIsInNodeList(oNodeToAppend, oSmilCustomTestNodes) Then
      'a customTest should be added to this elem
      oBruno.oCmn.oDomCmn.fncAppendAttribute oAbstractTimeContainerToAdd, "customTest", oNodeToAppend.nodeName
    End If
        
        
    'add par/seq into abstract smil, appended to seq returned from oAbstractSmilSequences.GetPointForTimeContainer
    'orig:
    'Set oLastAbstractSmilContainerAdded = oAbstractSmilSequences.GetPointForTimeContainer(oNodeToAppend).appendChild(oAbstractTimeContainerToAdd)
                    
    Set oPointForTimeContainer = oAbstractSmilSequences.GetPointForTimeContainer(oNodeToAppend)
    If oPointForTimeContainer Is Nothing Then
      'if we come here no match in any sequence, update
      Dim i As Long
      For i = 0 To oAbstractSmilSequences.lSequencesCount - 1
        oAbstractSmilSequences.Sequence(i).update
      Next
      Set oPointForTimeContainer = oAbstractSmilSequences.GetPointForTimeContainer(oNodeToAppend)
      If oPointForTimeContainer Is Nothing Then
        'major break, due to mixedcontenthandleractive
        Stop
        'add it after last child of root
        Set oLastAbstractSmilContainerAdded = oAbstractSmil.oDom.documentElement.appendChild(oAbstractTimeContainerToAdd)
      Else
        Set oLastAbstractSmilContainerAdded = oAbstractSmilSequences.GetPointForTimeContainer(oNodeToAppend).appendChild(oAbstractTimeContainerToAdd)
      End If
    Else
      'there was a match before update
      Set oLastAbstractSmilContainerAdded = oPointForTimeContainer.appendChild(oAbstractTimeContainerToAdd)
    End If
          
    Set oPointForTimeContainer = Nothing
    Set oAbstractTimeContainerToAdd = Nothing
End Function

Private Function fncCreateAbstractContainer(ByRef oNodeToAppend As IXMLDOMNode, sContentDocFileName As String, oAbstractSmil As cAbstractDocument, bNewPhysicalSmilFile As Boolean) As IXMLDOMNode
Dim oTempContainer As IXMLDOMElement
Dim oTempContainerChild As IXMLDOMElement
Dim oTempContainerAttr As IXMLDOMAttribute
Dim oContentNodeAttr As IXMLDOMAttribute
Dim oNodeToAppendChildNodes As IXMLDOMNodeList
Dim bLetItBeASequence As Boolean

'   if marked as sequence type, create "seq"
'   else, create "par"
'     ... and add "text" child to par
'   add id to container
'   add smilref to container
'   add class to container
  
  'for 2.02 there should never be sequences
  If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
    bLetItBeASequence = False
  Else 'it is a zed book
    'test1: test whether a wanted sequence should be forced to par
    If (oBruno.oCmn.oDomCmn.fncIsInNodeList(oNodeToAppend, oSmilSequenceNodes)) Then
      'this is marked as a "sequence" node in Driver file
      'it should only really become a sequence
      'if at least one child will become a syncpoint
      'which will not happen if all are empty, syncomit or text
      bLetItBeASequence = fncHasAtLeastOneGoodSyncPointChild(oNodeToAppend)
    End If '(oBruno.oCmn.oDomCmn.fncIsInNodeList(oNodeToAppend, oSmilSequenceNodes))
    
    'test2: test whether a wanted par should be forced to sequence
    If (oBruno.oCmn.oDomCmn.fncHasIgnorableWhiteSpaceOnly(oNodeToAppend)) _
      And (oNodeToAppend.selectNodes("./*").length > 0) Then
      'this elem has no text on its own, but has element children
      'in other words, sounds like a seq
      'but it should only really become a sequence
      'if at least one child will become a syncpoint
      bLetItBeASequence = fncHasAtLeastOneGoodSyncPointChild(oNodeToAppend)
    End If
  End If 'oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202
  
  If bLetItBeASequence Then
    Set oTempContainer = oAbstractSmil.oDom.createElement("seq")
    If (oBruno.oCmn.oDomCmn.fncIsInNodeList(oNodeToAppend, oSmilSequenceUserEscapeNodes)) Then
      'this container should have the userescape attribute
      'only add the base content of attribute since children ids not set yet
      oBruno.oCmn.oDomCmn.fncAppendAttribute oTempContainer, "end", "DTBuserEscape;"
    End If
  Else
    'add a <par id> to abstract smil
    Set oTempContainer = oAbstractSmil.oDom.createElement("par")
    'add a text child to the par, pointing back to node id
    Set oTempContainerChild = oBruno.oCmn.oDomCmn.fncAppendElement(oTempContainer, "text", , "src", sContentDocFileName & "#" & oNodeToAppend.selectSingleNode("@id").Text, "id", "tx" & oSmilTextIdGetter.fncGetId)
  End If
      
  'add a "newFile" attribute to be used by cOutputDocuments later on
  If bNewPhysicalSmilFile Then
    Set oTempContainerAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oTempContainer, "newFile", "true")
  End If
  
  'add id to the abstract timecontainer
  If oTempContainer.nodeName = "par" Then
    Set oTempContainerAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oTempContainer, "id", "tcp" & oSmilParIdGetter.fncGetId)
  Else
    Set oTempContainerAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oTempContainer, "id", "tcs" & oSmilSeqIdGetter.fncGetId)
  End If
    
  'add to contentdocnode a smilref attribute reffing above time container
  If (oBruno.oCmn.oDomCmn.fncIsInNodeList(oNodeToAppend, oSmilRefOmitNodes)) And (Not oBruno.oCmn.oDomCmn.fncIsInNodeList(oNodeToAppend, oNavIncludeNodes)) Then
    'dont add the linkback reference to content doc
  Else
    ''Set oContentNodeAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oNodeToAppend, "smilref", "abstract.smil#" & oTempContainer.selectSingleNode("@id").Text)
    'the below is used instead for optimization purposes
    'Set oContentNodeAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oNodeToAppend, "smilref", "s" & CStr(Format(lPhysicalSmilFileCount, "0000") & ".smil#" & oTempContainer.selectSingleNode("@id").Text))
    'redid 20040611:
    'if 2.02, smilref points to text elem, else point to timecontainer
    'this because pro cleanup removes all timecontainer ids
    Dim sXpath As String
    If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
      If oTempContainer.nodeName = "seq" Then
        'this shouldnt happen
        sXpath = "@id"
      Else
        sXpath = "/text/@id"
      End If
    Else
      'it is zed
      sXpath = "@id"
    End If
    Set oContentNodeAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oNodeToAppend, "smilref", "s" & CStr(Format(lPhysicalSmilFileCount, "0000") & ".smil#" & oTempContainer.selectSingleNode(sXpath).Text))
  End If
  
  'add a class attribute to time container, with dest element name
  Set oTempContainerAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oTempContainer, "class", oNodeToAppend.nodeName)

  'return the result
  Set fncCreateAbstractContainer = oTempContainer

End Function

Private Function fncHasAtLeastOneGoodSyncPointChild(oElem As IXMLDOMElement) As Boolean
Dim bAtLeastOneGoodSyncPointChild As Boolean
Dim oElemChildNodes As IXMLDOMNodeList
Dim oElemChildNode As IXMLDOMNode
  
  bAtLeastOneGoodSyncPointChild = False
  Set oElemChildNodes = oElem.selectNodes(".//*")
  If oElemChildNodes.length < 1 Then
    'children are text only
    bAtLeastOneGoodSyncPointChild = False
  Else
    For Each oElemChildNode In oElemChildNodes
      If (oBruno.oCmn.oDomCmn.fncIsInNodeList(oElemChildNode, oSyncOmitNodes)) Or (Not oElemChildNode.hasChildNodes) Then
        'it is is syncomit or is empty: not a good syncpoint
      Else
        bAtLeastOneGoodSyncPointChild = True
        Exit For
      End If
    Next
  End If 'oElemChildNodes.length < 1
  fncHasAtLeastOneGoodSyncPointChild = bAtLeastOneGoodSyncPointChild
  
End Function

Private Function fncAllElemSiblingsAreSyncOmit( _
    ByRef oSiblings As IXMLDOMNodeList _
    ) As Boolean

Dim oSibling As IXMLDOMNode
Dim bHasElements As Boolean
  fncAllElemSiblingsAreSyncOmit = True
  
  For Each oSibling In oSiblings
    If oSibling.nodeType = NODE_ELEMENT Then
      bHasElements = True
      If Not oBruno.oCmn.oDomCmn.fncIsInNodeList(oSibling, oSyncOmitNodes) Then
        fncAllElemSiblingsAreSyncOmit = False
        Exit For
      End If
    End If
  Next
  If Not bHasElements Then fncAllElemSiblingsAreSyncOmit = False
End Function

Private Function fncAllElemSiblingsAreSyncInclude( _
    ByRef oSiblings As IXMLDOMNodeList _
    ) As Boolean
Dim oSibling As IXMLDOMNode
Dim bHasElements As Boolean
  fncAllElemSiblingsAreSyncInclude = True
  
  For Each oSibling In oSiblings
    If oSibling.nodeType = NODE_ELEMENT Then
      bHasElements = True
      If oBruno.oCmn.oDomCmn.fncIsInNodeList(oSibling, oSyncOmitNodes) Then
        fncAllElemSiblingsAreSyncInclude = False
        Exit For
      End If
    End If
  Next
  If Not bHasElements Then fncAllElemSiblingsAreSyncInclude = False
End Function

Private Function fncMixedContentElemAppendIncludeOmitSiblings( _
    ByRef oTextNode As IXMLDOMText _
    ) As IXMLDOMElement

'gets a textnode as input
'includes that, and all siblings that are syncomit, into a new wrapper elem
'returns a newly created parent of the text
Dim oNewWrapperElem As IXMLDOMElement
Dim oNewWrapperElemAtt As IXMLDOMAttribute
Dim oMovedTextNode As IXMLDOMNode
Dim oMovedSibling As IXMLDOMNode
Dim oTestNode As IXMLDOMNode

  'create the new element
  Set oNewWrapperElem = oTextNode.ownerDocument.createElement(oBruno.oDriver.sItmtElemName)
  Set oNewWrapperElemAtt = oTextNode.ownerDocument.createAttribute(oBruno.oDriver.sItmtAttName)
  oNewWrapperElemAtt.Text = oBruno.oDriver.sItmtAttValue
  Set oNewWrapperElemAtt = oNewWrapperElem.Attributes.setNamedItem(oNewWrapperElemAtt)
  'add wrapper
  Set oNewWrapperElem = oTextNode.parentNode.insertBefore(oNewWrapperElem, oTextNode)

  'check the previous siblings
  Set oTestNode = oNewWrapperElem.previousSibling

  Do
    If Not oTestNode Is Nothing Then
      If ((oTestNode.nodeType = NODE_ELEMENT) And (oBruno.oCmn.oDomCmn.fncIsInNodeList(oTestNode, oSyncOmitNodes))) Or (oTestNode.nodeType = NODE_TEXT) Then
        'it is an element in syncomit, or a text node
          If Not oNewWrapperElem.firstChild Is Nothing Then
            Set oMovedSibling = oNewWrapperElem.insertBefore(oTestNode, oNewWrapperElem.firstChild)
          Else
            Set oMovedSibling = oNewWrapperElem.appendChild(oTestNode)
          End If
      Else
        'we should not put into wrapper
          GoTo one
      End If
      Set oTestNode = oNewWrapperElem.previousSibling
    End If
  Loop Until oTestNode Is Nothing
one:
  'move textnode into wrapper
  Set oMovedTextNode = oNewWrapperElem.appendChild(oTextNode)

  'check the next siblings
  Set oTestNode = oNewWrapperElem.nextSibling
  
  Do
    If Not oTestNode Is Nothing Then
      If ((oTestNode.nodeType = NODE_ELEMENT) And (oBruno.oCmn.oDomCmn.fncIsInNodeList(oTestNode, oSyncOmitNodes))) Or (oTestNode.nodeType = NODE_TEXT) Then
        'it is an element in syncomit, or a text node
        Set oMovedSibling = oNewWrapperElem.appendChild(oTestNode)
      Else
        'we should not put into wrapper
          GoTo two
      End If
      Set oTestNode = oNewWrapperElem.nextSibling
    End If
  Loop Until oTestNode Is Nothing

two:
  Set fncMixedContentElemAppendIncludeOmitSiblings = oNewWrapperElem
End Function

Public Function AbstractDocument(lDocument As Long) As cAbstractDocument
  Set AbstractDocument = aAbstractDocuments(lDocument)
End Function


Private Sub Class_Terminate()
  'debug.Print "cAbstractDocuments.Terminate"
  lAbstractDocumentCount = 0
  ReDim Preserve aAbstractDocuments(lAbstractDocumentCount)
  Set oAbstractSmil = Nothing
  Set oAbstractNavigation = Nothing
  Set oSmilParIdGetter = Nothing
  Set oSmilSeqIdGetter = Nothing
  Set oContentIdGetter = Nothing
  Set oSmilTextIdGetter = Nothing
  Set oRemoveNodes = Nothing
  Set oSyncOmitNodes = Nothing
  Set oSyncForceNodes = Nothing
  Set oSmilRefOmitNodes = Nothing
  Set oSmilSequenceNodes = Nothing
  Set oNavIncludeNodes = Nothing
  Set oNavIncludeNodesExplicitNavLabel = Nothing
  Set oMixedContentHandlerNodes = Nothing
  Set oLastAbstractSmilContainerAdded = Nothing
  Set oLastNodeAppended = Nothing
  Set oAbstractSmilSequences = Nothing
End Sub

Private Function fncSaveAbstractOutput()
  oBruno.oCmn.oFsoCmn.fncDeleteAllFilesInFolder (oBruno.oPaths.AppPath & "externals\abstractoutput\")
  Dim k As Long
  For k = 0 To lAbstractDocumentCount - 1
    oBruno.oCmn.oDomCmn.fncXmlIndentSax aAbstractDocuments(k).oDom
    If aAbstractDocuments(k).lAbstractType = TYPE_ABSTRACT_CONTENTDOC Then
      aAbstractDocuments(k).oDom.save (oBruno.oPaths.AppPath & "externals\abstractoutput\content" & CStr(k) & ".xml")
    ElseIf aAbstractDocuments(k).lAbstractType = TYPE_ABSTRACT_SMIL Then
      aAbstractDocuments(k).oDom.save (oBruno.oPaths.AppPath & "externals\abstractoutput\abstract.smil")
    ElseIf aAbstractDocuments(k).lAbstractType = TYPE_ABSTRACT_NAVIGATION Then
      aAbstractDocuments(k).oDom.save (oBruno.oPaths.AppPath & "externals\abstractoutput\nav.xml")
    Else
      aAbstractDocuments(k).oDom.save (oBruno.oPaths.AppPath & "externals\abstractoutput\unknown" & CStr(k) & ".xml")
    End If
  Next
End Function

