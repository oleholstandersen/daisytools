VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutputDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private aOutputDocuments() As cOutputDocument
Private lOutputDocumentCount As Long

Private oSpine As cSpine
Private oFileNameIdGetter As cIdGetter

Public oSmilCustomTests As cOutputCustomTests

'for conversion of romans
Private Const M = 1000
Private Const D = 500
Private Const C = 100
Private Const L = 50
Private Const X = 10
Private Const V = 5
Private Const i = 1

Private Sub Class_Initialize()
 'debug.Print "oOutputDocuments.initialize"
 lOutputDocumentCount = 0
 Set oFileNameIdGetter = New cIdGetter
 Set oSmilCustomTests = New cOutputCustomTests
End Sub

Public Function OutputDocumentCount() As Long
    OutputDocumentCount = lOutputDocumentCount
End Function

Public Function Document(lArrayItem As Long) As cOutputDocument
    Set Document = aOutputDocuments(lArrayItem)
End Function

Private Function fncAddDocument( _
    isFileName As String, _
    ilType As Long, _
    Optional ioDom As MSXML2.DOMDocument40, _
    Optional ilInputContentDoc As Variant _
) As Boolean

  On Error GoTo errh
  fncAddDocument = False
  ReDim Preserve aOutputDocuments(lOutputDocumentCount)
  Set aOutputDocuments(lOutputDocumentCount) = New cOutputDocument
  With aOutputDocuments(lOutputDocumentCount)
    .lType = ilType
    .sFileName = isFileName
    If Not ioDom Is Nothing Then
      Set .oDom = ioDom
    End If
        
    'if this is a content doc, add reference to source in inputarray
    'value is -1 if this is not a contentdoc
    If Not IsMissing(ilInputContentDoc) Then
      .lInputDocumentSource = ilInputContentDoc
    Else
      .lInputDocumentSource = -1
    End If
    
      
  End With
  
  lOutputDocumentCount = lOutputDocumentCount + 1
  fncAddDocument = True
errh:

End Function

Public Function fncInstantiate() As Boolean
Dim aNewSmilUri() As String
Dim lNewSmilUriCount As Long

  'takes the abstract fileset and makes an array of documents for a certain dtb standard
'''  On Error GoTo errh
  fncInstantiate = False
  Set oSpine = New cSpine  'spine is a list of smilfile names (1-n); used for mastersmil and opf

  If Not fncCreateSmilFiles(oBruno.oAbstractDocuments, aNewSmilUri(), lNewSmilUriCount) Then GoTo errh

  If Not fncCreateContentDocs(oBruno.oAbstractDocuments, aNewSmilUri(), lNewSmilUriCount) Then GoTo errh

  If Not fncCreateNavigation(oBruno.oAbstractDocuments, aNewSmilUri(), lNewSmilUriCount) Then GoTo errh
  
  'now create documents that have no representation in the abstract set, use spine object
  If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
    'create mastersmil for 202
    fncAddDocument "master.smil", TYPE_ACTUAL_202MSMIL, fncCreateD202MasterSmil
  ElseIf oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_Z39 Then
    'create opf for z39
    fncAddDocument oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension(oBruno.oAbstractDocuments.AbstractDocument(0).sFileName) & ".opf", TYPE_ACTUAL_Z39OPF, fncCreateZedOpf
  Else
  End If
  
  'add 2.02 (and lpPro specific stuff): these are placed at a specific place
  'that oOutputSaver knows about
  If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
    'create the lpp
    fncAddDocument "project.lpp", TYPE_ACTUAL_LPP
    fncCreateLpp
    'add a reference to the bogus mdf
    fncAddDocument "default.mdf", TYPE_ACTUAL_MDF
  End If
           
  fncInstantiate = True

errh:

End Function

Private Function fncCreateSmilFiles( _
    ByRef oAbstractDocuments As cAbstractDocuments, _
    ByRef aNewSmilUri() As String, _
    ByRef lNewSmilUriCount As Long _
    ) As Boolean

Dim i As Long
Dim oSmilSplitNodes As IXMLDOMNodeList
Dim oRemoveNodes As IXMLDOMNodeList, oRemoveNode As IXMLDOMNode

  On Error GoTo errh
  fncCreateSmilFiles = False
  
  For i = 0 To oAbstractDocuments.lAbstractDocumentCount - 1
    DoEvents
    With oAbstractDocuments.AbstractDocument(i)
      If .lAbstractType = TYPE_ABSTRACT_SMIL Then
        Dim oClonedNode As IXMLDOMNode
        Dim oNode As IXMLDOMNode
        Dim oEndSync As IXMLDOMNode
        Dim oCurrentAbstractSmilContextElem As IXMLDOMNode
        Dim bIsLastChild As Boolean
        
        Set oCurrentAbstractSmilContextElem = .oDom.documentElement.firstChild
        Set oSmilSplitNodes = .oDom.selectNodes("//*[@newFile]")
                
        'remove 'newFile' attribute
        Set oRemoveNodes = .oDom.selectNodes("//*[@newFile]")
        For Each oRemoveNode In oRemoveNodes
          oBruno.oCmn.oDomCmn.fncRemoveAttribute oRemoveNode, "newFile"
        Next oRemoveNode
                
        'loop and create 1-n smilfile objects
        Do
          Dim oSmilDom As New MSXML2.DOMDocument40
              oSmilDom.validateOnParse = False
              oSmilDom.resolveExternals = False
              oSmilDom.preserveWhiteSpace = False
              oSmilDom.setProperty "SelectionLanguage", "XPath"
              oSmilDom.setProperty "NewParser", True
              oSmilDom.setProperty "SelectionNamespaces", "xmlns:smil2='http://www.w3.org/2001/SMIL20/'"
          
          Dim sCurrentSmilFileName As String
          Dim sCurrentSmilFileType As Long
          
          sCurrentSmilFileName = "s" & oFileNameIdGetter.fncGetId & ".smil"
          
          'add the smil file to the spine object
          oSpine.fncAppendSpineItem sCurrentSmilFileName
          
          If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
            sCurrentSmilFileType = TYPE_ACTUAL_202SMIL
            If Not oBruno.oCmn.oDomCmn.fncParseFile(oBruno.oPaths.ShellPath & "d202smil.shell", oSmilDom, "") Then GoTo errh
          ElseIf oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_Z39 Then
            sCurrentSmilFileType = TYPE_ACTUAL_Z39SMIL
            If Not oBruno.oCmn.oDomCmn.fncParseFile(oBruno.oPaths.ShellPath & "z39smil.shell", oSmilDom, "") Then GoTo errh
          Else
'            sCurrentSmilFileType = TYPE_ACTUAL_W3CSMIL
'            If Not oBruno.oCmn.oDomCmn.fncParseFile(oBruno.oPaths.ShellPath & "w3csmil.shell", oSmilDom, "") Then GoTo errH
          End If
          
          'loop through abstract smil and add children
          For Each oNode In .oDom.documentElement.childNodes
           DoEvents
           If oNode Is oCurrentAbstractSmilContextElem Then
             'do the above to skip nodes added in possible earlier iterations
             Set oClonedNode = oSmilDom.documentElement.lastChild.firstChild.appendChild(oNode.cloneNode(True))
             'add the uri to aNewSmilUri array: when content and nav is created, replace
''made redundant during optimization:
''               ReDim Preserve aNewSmilUri(lNewSmilUriCount)
''               aNewSmilUri(lNewSmilUriCount) = sCurrentSmilFileName & "#" & oClonedNode.selectSingleNode("@id").Text
''               lNewSmilUriCount = lNewSmilUriCount + 1
             
             'add endsync last to all pars
             'error, only occurs on body children
             If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
               If oClonedNode.nodeName = "par" Then Set oEndSync = oBruno.oCmn.oDomCmn.fncAppendAttribute(oClonedNode, "endsync", "last")
             End If
             
             If Not oNode Is oNode.parentNode.lastChild Then
               Set oCurrentAbstractSmilContextElem = oNode.nextSibling
               'if next elem is in smil-split, exit loop
               If oBruno.oCmn.oDomCmn.fncIsInNodeList(oNode.nextSibling, oSmilSplitNodes) Then
                 Exit For
               End If
             Else
               bIsLastChild = True
             End If
           End If
          Next
                                                           
          If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
            'add metadata to smilfile
            oBruno.oCmn.oDomCmn.fncAppendElement oSmilDom.selectSingleNode("//head"), "meta", , "name", "dc:identifier", "content", oBruno.oInputDocuments.oInputMetadata.sDcIdentifier, , , , , True
            Dim oNccGenNode As IXMLDOMNode
            Set oNccGenNode = oSmilDom.selectSingleNode("//meta[@name='ncc:generator']/@content")
            If Not oNccGenNode Is Nothing Then
              If oBruno.sAppVersion = "" Then
                oNccGenNode.Text = "bruno"
              Else
                oNccGenNode.Text = oBruno.sAppVersion
              End If
            End If
            
            'remove class attributes
            Dim oElemsWithClassAttrs As IXMLDOMNodeList, oElem As IXMLDOMNode
            Set oElemsWithClassAttrs = oSmilDom.selectNodes("//body//*[@class]")
            For Each oElem In oElemsWithClassAttrs
              oBruno.oCmn.oDomCmn.fncRemoveAttribute oElem, "class"
            Next
          ElseIf oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_Z39 Then
            'add metadata to smilfile
            oSmilDom.selectSingleNode("//smil2:meta[@name='dtb:uid']/@content").Text = oBruno.oInputDocuments.oInputMetadata.sDtbookUid
            oSmilDom.selectSingleNode("//smil2:meta[@name='dtb:generator']/@content").Text = oBruno.sAppVersion
            'check for customTest matches in the smilfile
            fncSmilCustomTest oSmilDom
          Else
            'raw
          End If
                          
          'append to output document array
          fncAddDocument sCurrentSmilFileName, sCurrentSmilFileType, oSmilDom.cloneNode(True)
          Set oSmilDom = Nothing
          
         'exit loop when last child in abstract smil has been added:
        Loop Until bIsLastChild
      End If
    End With
  Next i
  
  fncCreateSmilFiles = True
errh:
End Function

Private Function fncCreateContentDocs( _
    ByRef oAbstractDocuments As cAbstractDocuments, _
    ByRef aNewSmilUri() As String, _
    ByRef lNewSmilUriCount As Long _
    ) As Boolean
Dim oTmpDom As New MSXML2.DOMDocument40
  oTmpDom.validateOnParse = False
  oTmpDom.resolveExternals = False
  oTmpDom.preserveWhiteSpace = False
  oTmpDom.setProperty "SelectionLanguage", "XPath"
  oTmpDom.setProperty "NewParser", True
  'oTmpDom.setProperty "SelectionNamespaces", "xmlns:ncx='http://www.loc.gov/nls/z3986/2004/ncx/'"

Dim i As Long
Dim sCurrentContentDocFileName As String
Dim sCurrentContentDocFileType As Long

    fncCreateContentDocs = False
    
    For i = 0 To oAbstractDocuments.lAbstractDocumentCount - 1
      With oAbstractDocuments.AbstractDocument(i)
        If .lAbstractType = TYPE_ABSTRACT_CONTENTDOC Then
          sCurrentContentDocFileName = .sFileName
          
          'since the content docs are validated beforehand, just reference and continue:
          Set oTmpDom = .oDom
          'fix the uris filenames (all are now "abstract.smil")
          If Not fncReplaceSmilUris(oTmpDom, "//@smilref|//@href", aNewSmilUri(), lNewSmilUriCount) Then GoTo errh
        
          If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
            sCurrentContentDocFileType = TYPE_ACTUAL_202CONTENT
            'abstracttype uses @smilref, change that to <a href
            If Not fncSmilRefToAnchor(oTmpDom) Then GoTo errh
            'fix empty <title> since pro freaks out otherwise
            If Not fncSetContentDocTitle(oTmpDom) Then GoTo errh
          ElseIf oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_Z39 Then
            sCurrentContentDocFileType = TYPE_ACTUAL_Z39CONTENT
          Else
            'raw
          End If
          'append to array
          fncAddDocument sCurrentContentDocFileName, sCurrentContentDocFileType, oTmpDom.cloneNode(True), .lInputContentDocSource
          Set oTmpDom = Nothing
        End If
      End With
     Next i
     fncCreateContentDocs = True
errh:
End Function

Private Function fncCreateNavigation( _
    ByRef oAbstractDocuments As cAbstractDocuments, _
    ByRef aNewSmilUri() As String, _
    ByRef lNewSmilUriCount As Long _
    ) As Boolean
Dim oTmpDom As New MSXML2.DOMDocument40
  oTmpDom.validateOnParse = False
  oTmpDom.resolveExternals = False
  oTmpDom.preserveWhiteSpace = False
  oTmpDom.setProperty "SelectionLanguage", "XPath"
  oTmpDom.setProperty "NewParser", True

Dim i As Long
Dim lNavDocType As Long
Dim sNavDocFileName As String

  fncCreateNavigation = False
  For i = 0 To oAbstractDocuments.lAbstractDocumentCount - 1
    With oAbstractDocuments.AbstractDocument(i)
      If .lAbstractType = TYPE_ABSTRACT_NAVIGATION Then
        If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
          lNavDocType = TYPE_ACTUAL_202NCC
          sNavDocFileName = "ncc.html"
          Set oTmpDom = fncCreateD202Ncc(.oDom)
          'add metadata to ncc
          fncAddMetaDataToNcc oTmpDom
        ElseIf oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_Z39 Then
          lNavDocType = TYPE_ACTUAL_Z39NCX
          'Debug.Print oAbstractDocuments.AbstractDocument(i).oDom.xml
          'sNavDocFileName = oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension(oBruno.oInputDocuments.InputDocument(0).sFileName) & ".ncx"
          sNavDocFileName = oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension(oBruno.oAbstractDocuments.AbstractDocument(0).sFileName) & ".ncx"
          Set oTmpDom = fncCreateZedNcx(.oDom)
        Else
         'raw
        End If
        'fix the smil references from nav doc (all now point to abstract smil)
        If Not fncReplaceSmilUris(oTmpDom, "//@href | //@src", aNewSmilUri(), lNewSmilUriCount) Then GoTo errh
        
        fncAddDocument sNavDocFileName, lNavDocType, oTmpDom.cloneNode(True)
        Set oTmpDom = Nothing

      End If
    End With
  Next i

  fncCreateNavigation = True
errh:
End Function

Private Function fncCreateD202MasterSmil() As MSXML2.DOMDocument40
Dim oDom As New MSXML2.DOMDocument40
  oDom.async = False
  oDom.validateOnParse = False
  oDom.resolveExternals = False
  oDom.preserveWhiteSpace = False
  oDom.setProperty "SelectionLanguage", "XPath"
  oDom.setProperty "NewParser", True
Dim i As Long
Dim oHeadnode As IXMLDOMNode
  

  On Error GoTo errh
  If Not oBruno.oCmn.oDomCmn.fncParseFile(oBruno.oPaths.ShellPath & "d202msmil.shell", oDom, "") Then GoTo errh
  For i = 0 To oSpine.lSpineItemCount - 1
    DoEvents
    oBruno.oCmn.oDomCmn.fncAppendElement _
      oDom.documentElement.lastChild, "ref", , _
      "src", oSpine.fncGetSpineItem(i).selectSingleNode("@name").Text, _
      "id", oSpine.fncGetSpineItem(i).selectSingleNode("@id").Text
  Next
  'add some metas
  Set oHeadnode = oDom.selectSingleNode("//head")
  oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:identifier", "content", oBruno.oInputDocuments.oInputMetadata.sDcIdentifier, , , , , True
  oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:title", "content", oBruno.oInputDocuments.oInputMetadata.sDcTitle, , , , , True
  
  Dim oNccGenNode As IXMLDOMNode
  Set oNccGenNode = oDom.selectSingleNode("//meta[@name='ncc:generator']/@content")
  If Not oNccGenNode Is Nothing Then
    If oBruno.sAppVersion = "" Then
      oNccGenNode.Text = "bruno"
    Else
      oNccGenNode.Text = oBruno.sAppVersion
    End If
  End If

  Set fncCreateD202MasterSmil = oDom.cloneNode(True)
  Set oDom = Nothing
  'debug.Print oDom.xml
  Exit Function
errh:
  Set oDom = Nothing
End Function

Private Sub Class_Terminate()
  'debug.Print "oOutputDocuments.terminate"
  lOutputDocumentCount = 0
  ReDim Preserve aOutputDocuments(lOutputDocumentCount)
  Set oSpine = Nothing
  Set oFileNameIdGetter = Nothing
  Set oSmilCustomTests = Nothing
End Sub

Private Function fncReplaceSmilUrisSlow(ByRef oDomToReplaceIn As MSXML2.DOMDocument40, ByVal sXpathUriCarriers As String, ByRef aNewSmilUriArray() As String, ByRef lNewSmilUriArrayCount As Long) As Boolean
Dim oUriCarriers As IXMLDOMNodeList
Dim oUriCarrier As IXMLDOMNode
Dim i As Long
  fncReplaceSmilUrisSlow = False
  Set oUriCarriers = oDomToReplaceIn.selectNodes(sXpathUriCarriers)
  Stop
  'Debug.Print oDomToReplaceIn.xml
  If Not oUriCarriers Is Nothing Then
    For Each oUriCarrier In oUriCarriers
      For i = 0 To lNewSmilUriArrayCount - 1
        If oUriCarrier.Text = "abstract.smil#" & oBruno.oCmn.oUriParser.fncGetId(aNewSmilUriArray(i)) Then
          'this is possible since all par ids ("p0000") are unique in the whole dtb
          oUriCarrier.Text = aNewSmilUriArray(i)
          Exit For
        End If
      Next
    Next
  End If
  Stop
  'Debug.Print oDomToReplaceIn.xml
  fncReplaceSmilUrisSlow = True
End Function

Private Function fncReplaceSmilUris( _
    ByRef oDomToReplaceIn As MSXML2.DOMDocument40, _
    ByVal sXpathUriCarriers As String, _
    ByRef aNewSmilUriArray() As String, _
    ByRef lNewSmilUriArrayCount As Long _
    ) As Boolean
    'this was made redundant during optimization
    fncReplaceSmilUris = True
    Exit Function
    
    
    
Dim i As Long
Dim sDomToReplaceIn As String, sNewUri As String, sExistingUri As String
  fncReplaceSmilUris = False

  Debug.Print "fncReplaceSmilUris in: " & Now

  sDomToReplaceIn = oDomToReplaceIn.xml
  For i = 0 To lNewSmilUriArrayCount - 1
    sNewUri = aNewSmilUriArray(i)
    sExistingUri = "abstract.smil#" & Mid$(sNewUri, InStr(1, sNewUri, "#") + 1)
    sDomToReplaceIn = Replace$(sDomToReplaceIn, sExistingUri, sNewUri)
  Next

  Debug.Print "fncReplaceSmilUris out: " & Now
  
  oDomToReplaceIn.loadXML (sDomToReplaceIn)
    
  fncReplaceSmilUris = True
End Function

Private Function fncSmilRefToAnchor(ByRef oDom As MSXML2.DOMDocument40) As Boolean
Dim oSmilRefElems As IXMLDOMNodeList
Dim oSmilRefElem As IXMLDOMNode
Dim oSmilRefAttr As IXMLDOMNode
Dim oNewAnchorNode As IXMLDOMElement
Dim oNewAnchorElems As IXMLDOMNodeList
Dim oNewHrefNode As IXMLDOMAttribute
Dim oNode As IXMLDOMNode
Dim oMovedNode As IXMLDOMNode
Dim oErrorTest As IXMLDOMNodeList

'turns @smilref into a wrapping anchor child, for xhtml

  On Error GoTo errh
  fncSmilRefToAnchor = False

  Set oSmilRefElems = oDom.documentElement.selectNodes("//*[@smilref]")
    
  For Each oSmilRefElem In oSmilRefElems
    'if there isnt a to-be descendant anchor
    If Not oBruno.oCmn.oDomCmn.fncIsRelativeInNodeList(oSmilRefElem, oSmilRefElems, RELATION_DESCENDANT) Then
      'create the element to add
      Set oNewAnchorNode = oDom.createElement("a")
      Set oNewHrefNode = oDom.createAttribute("href")
      oNewHrefNode.Text = oSmilRefElem.selectSingleNode("@smilref").Text
      Set oNewHrefNode = oNewAnchorNode.Attributes.setNamedItem(oNewHrefNode)
      'add anchor
      Set oNewAnchorNode = oSmilRefElem.insertBefore(oNewAnchorNode, oSmilRefElem.firstChild)
      'move content as descendant of anchor
      For Each oNode In oSmilRefElem.childNodes
        If Not oNode Is oNewAnchorNode Then
          Set oMovedNode = oNewAnchorNode.appendChild(oNode)
        End If
      Next
     Else
       'frmMain.fncAddMessage "#" & oSmilRefElem.selectSingleNode("@id").Text & " has linkback removed because of anchor nesting"
     End If
     Set oSmilRefAttr = oSmilRefElem.Attributes.removeNamedItem("smilref")
  Next
    
  Set oErrorTest = oDom.selectNodes("//xhtml:a[.//xhtml:a]|//a[.//a]|//xhtml:a[.//a]|//a[.//xhtml:a]")
  If oErrorTest.length > 0 Then
    frmMain.fncAddMessage "Warning: " & oErrorTest.length & " nested anchors in output contentdoc"
    Dim oErrorNode As IXMLDOMNode
    For Each oErrorNode In oErrorTest
      Debug.Print "nested anchors: " & oErrorNode.parentNode.nodeName
    Next
  End If
   
  fncSmilRefToAnchor = True
errh:
End Function

Private Function fncHandleNccContent(ByRef inElement As IXMLDOMElement) As Boolean
'since ncc childrens content model is restricted to an anchor with text,
'we need to deal with multielement content.
'In some cases the inElement has an attribute bruno_navLabel to use,
'in other cases not.

Dim oPresetBrunoNavlabel As IXMLDOMAttribute

    Set oPresetBrunoNavlabel = inElement.Attributes.getNamedItem("bruno_navLabel")
    If Not oPresetBrunoNavlabel Is Nothing Then
      'a label was preset in driver
      'delete all content and insert the label
      oBruno.oCmn.oDomCmn.fncRemoveAllChildElements inElement
      inElement.Text = oPresetBrunoNavlabel.Text
      inElement.Attributes.removeNamedItem ("bruno_navLabel")
    Else
      'no label was preset in driver
      'do something with the existing content
      'but first find out if it is needed
      Dim oInElementChildElems As IXMLDOMNodeList
      Set oInElementChildElems = inElement.selectNodes("./*")
      If Not oInElementChildElems Is Nothing Then
        'it is needed, there was at least one element child
        Dim sCompleteText As String
        sCompleteText = inElement.nodeTypedValue
        oBruno.oCmn.oDomCmn.fncRemoveAllChildElements inElement
        inElement.Text = sCompleteText
      Else
        'we assume inElement only had a textnode as child
      End If
    End If ' Not oPresetBrunoNavlabel Is Nothing

   fncHandleNccContent = True

End Function

Private Function fncCreateD202Ncc(oAbstractNavDom As MSXML2.DOMDocument40) As MSXML2.DOMDocument40
Dim oNccDom As New MSXML2.DOMDocument40
  oNccDom.validateOnParse = False
  oNccDom.resolveExternals = False
  oNccDom.preserveWhiteSpace = False
  oNccDom.setProperty "SelectionLanguage", "XPath"
  oNccDom.setProperty "NewParser", True
  oNccDom.setProperty "SelectionNamespaces", "xmlns:xhtml='http://www.w3.org/1999/xhtml'"
Dim oNccDomBodyNode As IXMLDOMNode
Dim oAbstractNode As IXMLDOMNode
Dim oFirstNode As IXMLDOMElement
Dim oCloneNode As IXMLDOMElement
Dim oAttrNode As IXMLDOMAttribute
'takes the abstract navigation dom as input
'that is a document <navigation> with arbitrary sequentially ordered elements
'turns it into a proper ncc document using the available shells
  
  'Debug.Print oAbstractNavDom.xml
  'oAbstractNavDom.save "D:\save.xml"
  
  If Not oBruno.oCmn.oDomCmn.fncParseFile(oBruno.oPaths.ShellPath & "d202ncc.shell", oNccDom, "") Then GoTo errh
  Set oNccDomBodyNode = oNccDom.selectSingleNode("//xhtml:body")
  For Each oAbstractNode In oAbstractNavDom.documentElement.childNodes
    Set oCloneNode = oAbstractNode.cloneNode(True)
    fncHandleNccContent oCloneNode
    Set oCloneNode = oNccDomBodyNode.appendChild(oCloneNode)
  Next
  
  'change version meta
  Set oAbstractNode = oNccDom.selectSingleNode("//xhtml:meta[@name='ncc:generator']/@content")
  If Not oAbstractNode Is Nothing Then oAbstractNode.Text = oBruno.sAppVersion
    
  'change the smilref attrs to anchors
  If Not fncSmilRefToAnchor(oNccDom) Then GoTo errh
  
  Set oFirstNode = oNccDomBodyNode.firstChild
  If Not oFirstNode Is Nothing Then
    'add class=title
    Set oAttrNode = oFirstNode.ownerDocument.createAttribute("class")
    oAttrNode.nodeValue = "title"
    Set oAttrNode = oFirstNode.Attributes.setNamedItem(oAttrNode)
  End If
  
 ' 20070201: this done in fncHandleNccContent instead
 ' 'now delete any elems nested inside ncc anchor to comply with d202ncc.dtd
 ' Dim oNewAnchorElems As IXMLDOMNodeList
 ' Dim oNewAnchorNode As IXMLDOMNode
 ' Dim sCompleteText As String
 ' 'Debug.Print oNccDom.xml
 '
 ' Set oNewAnchorElems = oNccDom.selectNodes("//a[./*]")
 ' For Each oNewAnchorNode In oNewAnchorElems
 '   'Stop
 '   'Debug.Print oNewAnchorNode.xml
 '   sCompleteText = oNewAnchorNode.nodeTypedValue
 '   oBruno.oCmn.oDomCmn.fncRemoveAllChildElements oNewAnchorNode
 '   oNewAnchorNode.Text = sCompleteText
 '   'Debug.Print oNewAnchorNode.xml
 ' Next oNewAnchorNode
    
  'bugfixes added 20070119:
  If Not fncDivToSpan(oNccDom) Then GoTo errh
  If Not fncRemoveBodyref(oNccDom) Then GoTo errh
        
  Set fncCreateD202Ncc = oNccDom
  Exit Function

errh:
  'return the input
  fncCreateD202Ncc = oAbstractNavDom
End Function

Private Function fncDivToSpan(ByRef oNccDom) As Boolean
  '20070119: any div.sidebar and div.prodnote from the content doc
  'changed to span.sidebar and span.prodnote
  'special tweak for TPBs content docs
  
  fncDivToSpan = False
  
  Dim oDivElems As IXMLDOMNodeList
  Dim oDivElem As IXMLDOMNode
  
  Set oDivElems = oNccDom.selectNodes("//xhtml:div[@class='sidebar']|//xhtml:div[@class='optional-prodnote']")
  For Each oDivElem In oDivElems
    Dim oSpanElem As IXMLDOMNode
    Set oSpanElem = oBruno.oCmn.oDomCmn.fncRenameElement(oDivElem, "span", "http://www.w3.org/1999/xhtml")
    oDivElem.parentNode.replaceChild oSpanElem, oDivElem
  Next oDivElem

  fncDivToSpan = True
End Function

Private Function fncRemoveBodyref(ByRef oNccDom) As Boolean
  '20070119: remove any lingering bodyref attributes
  
  fncRemoveBodyref = False
  
  Dim oBodyrefElems As IXMLDOMNodeList
  Dim oBodyrefElem As IXMLDOMNode
  
  Set oBodyrefElems = oNccDom.selectNodes("//xhtml:span[@bodyref]")
  For Each oBodyrefElem In oBodyrefElems
    oBruno.oCmn.oDomCmn.fncRemoveAttribute oBodyrefElem, "bodyref"
  Next oBodyrefElem

  fncRemoveBodyref = True
End Function

Private Function fncCreateZedNcx(oAbstractNavDom As MSXML2.DOMDocument40) As MSXML2.DOMDocument40
Dim oNcxDom As New MSXML2.DOMDocument40
  oNcxDom.validateOnParse = False
  oNcxDom.resolveExternals = False
  oNcxDom.preserveWhiteSpace = False
  oNcxDom.setProperty "SelectionLanguage", "XPath"
  oNcxDom.setProperty "NewParser", True
  oNcxDom.setProperty "SelectionNamespaces", "xmlns:ncx='http://www.daisy.org/z3986/2005/ncx/'"
Dim oTmpDom As New MSXML2.DOMDocument40
  oTmpDom.validateOnParse = False
  oTmpDom.resolveExternals = False
  oTmpDom.preserveWhiteSpace = False
  oTmpDom.setProperty "SelectionLanguage", "XPath"
  oTmpDom.setProperty "NewParser", True
  oTmpDom.setProperty "SelectionNamespaces", "xmlns:ncx='http://www.daisy.org/z3986/2005/ncx/'"

Dim oNcxDomBodyNode As IXMLDOMNode
Dim oNcxDomNavMapNode As IXMLDOMNode
Dim oAbstractNavNode As IXMLDOMNode
Dim oAbstractNavNodes As IXMLDOMNodeList
Dim oAbstractNavNodesToMap As IXMLDOMNodeList

Dim oFirstNode As IXMLDOMElement
Dim oCloneNode As IXMLDOMElement
Dim oAttrNode As IXMLDOMAttribute
Dim oNode As IXMLDOMNode
Dim k As Long
Dim oNcxIdGetter As cIdGetter
Dim sNavPointShell As String

'takes the abstract navigation dom as input
'that is a document <navigation> with arbitrary sequentially ordered elements
'turns it into a proper ncx document using the available shells

  Set oNcxIdGetter = New cIdGetter
  If Not oBruno.oCmn.oDomCmn.fncParseFile(oBruno.oPaths.ShellPath & "z39ncx.shell", oNcxDom, "") Then GoTo errh
  
  Set oNcxDomBodyNode = oNcxDom.selectSingleNode("//ncx:ncx")
  Set oNcxDomNavMapNode = oNcxDom.selectSingleNode("//ncx:ncx/ncx:navMap")
  Set oAbstractNavNodes = oAbstractNavDom.documentElement.selectNodes(".//*")
  Set oAbstractNavNodesToMap = oAbstractNavDom.selectNodes("//dtbook:h1|//dtbook:h2|//dtbook:h3|//dtbook:h4|//dtbook:h5|//dtbook:h6")
              
  'add playOrder attribute
  k = 0
  For Each oAbstractNavNode In oAbstractNavDom.documentElement.selectNodes("./*")
    k = k + 1
    oBruno.oCmn.oDomCmn.fncAppendAttribute oAbstractNavNode, "playOrder", CStr(k)
  Next oAbstractNavNode
    
  'do some head stuff
  Set oNode = oNcxDom.selectSingleNode("//ncx:ncx/ncx:docTitle/ncx:text")
  If Not oNode Is Nothing Then oNode.Text = oBruno.oInputDocuments.oInputMetadata.sDtbookDocTitle
  
  Set oNode = oNcxDom.selectSingleNode("//ncx:ncx/ncx:docAuthor/ncx:text")
  If Not oNode Is Nothing Then oNode.Text = oBruno.oInputDocuments.oInputMetadata.sDtbookDocAuthor
  
  Set oNode = oNcxDom.selectSingleNode("//ncx:ncx/ncx:head/ncx:meta[@name='dtb:uid']/@content")
  If Not oNode Is Nothing Then oNode.Text = oBruno.oInputDocuments.oInputMetadata.sDtbookUid
    
  Set oNode = oNcxDom.selectSingleNode("//ncx:ncx/ncx:head/ncx:meta[@name='dtb:generator']/@content")
  If Not oNode Is Nothing Then oNode.Text = oBruno.sAppVersion
        
' <meta name="dtb:depth" content="0"/>
  Set oNode = oNcxDom.selectSingleNode("//ncx:ncx/ncx:head/ncx:meta[@name='dtb:depth']/@content")
  If Not oNode Is Nothing Then
    oNode.Text = fncGetHeadingDepth(oAbstractNavNodesToMap)
  End If

' <meta name="dtb:totalPageCount" content="0"/>
  Dim oPageNodes As IXMLDOMNodeList
  Set oPageNodes = oAbstractNavDom.selectNodes("//dtbook:pagenum")
  
  Set oNode = oNcxDom.selectSingleNode("//ncx:ncx/ncx:head/ncx:meta[@name='dtb:totalPageCount']/@content")
  If Not oPageNodes Is Nothing Then
    If Not oNode Is Nothing Then
      oNode.Text = oPageNodes.length
    End If
  End If
  
' <meta name="dtb:maxPageNumber" content="0"/>
  Set oNode = oNcxDom.selectSingleNode("//ncx:ncx/ncx:head/ncx:meta[@name='dtb:maxPageNumber']/@content")
  Dim oLastPageNode As IXMLDOMNode
  Set oLastPageNode = oPageNodes.Item(oPageNodes.length - 1)
  If Not oLastPageNode Is Nothing Then
    oNode.Text = oLastPageNode.Text
  End If
    
  'add customTests from oSmilCustomTests
  Dim s As Long, oNcxCustomTest As IXMLDOMElement
  For Each oNode In oSmilCustomTests.oTests.documentElement.childNodes
    Set oNcxCustomTest = oNcxDom.documentElement.firstChild.appendChild(oSmilCustomTests.oTests.documentElement.childNodes.Item(s).cloneNode(False))
    s = s + 1
  Next oNode
  
  'create navMap
  sNavPointShell = _
  "<navPoint class='class' id='id' playOrder='integer'>" & _
  "  <navLabel><text>navLabel.text</text></navLabel>" & _
  "  <content src='smilref' />" & _
  "</navPoint>"
  
  Dim lCurrentNestingLevel As Long
  Dim lPreviousNestingLevel As Long
  Dim oLastInsertedNavPoint As IXMLDOMNode
  
  lCurrentNestingLevel = 0
  lPreviousNestingLevel = 0
  
  For Each oAbstractNavNode In oAbstractNavNodesToMap
      'set lNestingLevel for for navpoint to be inserted
      lCurrentNestingLevel = CLng(fncMakeNumeric(oAbstractNavNode.baseName))
      
      'create a freefloating navpoint
      oTmpDom.loadXML (sNavPointShell)
      oTmpDom.selectSingleNode("//navPoint/@class").Text = oAbstractNavNode.baseName
      oTmpDom.selectSingleNode("//navPoint/@id").Text = "nx" & oNcxIdGetter.fncGetId
      oTmpDom.selectSingleNode("//navPoint/@playOrder").Text = oAbstractNavNode.selectSingleNode("@playOrder").Text
      oTmpDom.selectSingleNode("//navPoint/navLabel/text").Text = oAbstractNavNode.Text
      oTmpDom.selectSingleNode("//navPoint/content/@src").Text = oAbstractNavNode.selectSingleNode("@smilref").Text
      
      'insert it
      If lPreviousNestingLevel = 0 Then
        'first append, always to root of navMap
        Set oLastInsertedNavPoint = oNcxDomNavMapNode.appendChild(oTmpDom.documentElement.cloneNode(True))
      Else
        'not first append, check what nesting level should be used
        If lPreviousNestingLevel = lCurrentNestingLevel Then
          Set oLastInsertedNavPoint = oLastInsertedNavPoint.parentNode.appendChild(oTmpDom.documentElement.cloneNode(True))
        ElseIf lPreviousNestingLevel < lCurrentNestingLevel Then
         'one step down, inside previous navPoint
         Set oLastInsertedNavPoint = oLastInsertedNavPoint.appendChild(oTmpDom.documentElement.cloneNode(True))
        Else '(lPreviousNestingLevel > lCurrentNestingLevel)
          'back up
          Dim nestChange As Long
          nestChange = lPreviousNestingLevel - lCurrentNestingLevel
          Dim oNewParent As IXMLDOMNode
          Set oNewParent = oLastInsertedNavPoint.parentNode 'start from here
          
          Dim r As Long
          For r = 0 To nestChange - 1
            Set oNewParent = oNewParent.parentNode
          Next r
          
          Set oLastInsertedNavPoint = oNewParent.appendChild(oTmpDom.documentElement.cloneNode(True))
        
        End If
      End If 'lPreviousNestingLevel = 0
      lPreviousNestingLevel = lCurrentNestingLevel
  Next
    
  'For Each oAbstractNavNode In oAbstractNavNodes
  'mg20070308: should only iter through direct descendants of root
  For Each oAbstractNavNode In oAbstractNavDom.documentElement.selectNodes("./*")
    If Not oBruno.oCmn.oDomCmn.fncIsInNodeList(oAbstractNavNode, oAbstractNavNodesToMap) Then
      'this node is not included in navMap
      'Debug.Print "navlisting a " & oAbstractNavNode.baseName
      If Not fncNavListExistsForThisType(oAbstractNavNode.baseName, oNcxDom) Then
        'create a new navList
        fncCreateNavList oAbstractNavNode, oNcxDom, oNcxIdGetter
      End If
      'append navtarget to navlist
      fncAppendNavTargetToNavList oAbstractNavNode, oNcxDom, oNcxIdGetter
    End If
  Next oAbstractNavNode
      
  Set fncCreateZedNcx = oNcxDom
  Exit Function

errh:
  'on error return the input
  fncCreateZedNcx = oAbstractNavDom
End Function

Private Function fncGetHeadingDepth(oAbstractNavNodesToMap As IXMLDOMNodeList) As String
 'oAbstractNavNodesToMap contains all h1-h6 nodes
 If oAbstractNavNodesToMap Is Nothing Then fncGetHeadingDepth = "0": Exit Function
 If oAbstractNavNodesToMap.length = 0 Then fncGetHeadingDepth = "0": Exit Function
 
 Dim oNode As IXMLDOMNode
 Dim highest As Long: highest = 0
 For Each oNode In oAbstractNavNodesToMap
   If oNode.nodeName = "h1" Then
     If highest < 1 Then highest = 1
   ElseIf oNode.nodeName = "h2" Then
     If highest < 2 Then highest = 2
   ElseIf oNode.nodeName = "h3" Then
     If highest < 3 Then highest = 3
   ElseIf oNode.nodeName = "h4" Then
     If highest < 4 Then highest = 4
   ElseIf oNode.nodeName = "h5" Then
     If highest < 5 Then highest = 5
   ElseIf oNode.nodeName = "h6" Then
     If highest < 6 Then highest = 6
   End If
 Next oNode
 
 fncGetHeadingDepth = highest
 
End Function



Private Function fncAppendNavTargetToNavList( _
  oAbstractNavNode As IXMLDOMNode, _
  oNcxDom As MSXML2.DOMDocument40, _
  oNcxIdGetter As cIdGetter) As Boolean
'locate the navList to append to via basename/class
'create navtarget, append it as lastchild to navList
Dim oMyNavList As IXMLDOMNode
Dim oNewNavTarget As IXMLDOMNode
Dim oNewNavLabel As IXMLDOMNode
Dim oNewNavLabelText As IXMLDOMNode
Dim oNewContent As IXMLDOMNode
Dim oNewTypeAttr As IXMLDOMAttribute
Dim oNewValueAttr As IXMLDOMAttribute

'    <navTarget class="abstractnode.basename" id="ncxid" playorder="integer">
'      <navLabel>
'        <text>abstractnode.text</text>
'      </navLabel>
'      <content src="smilref" />
'    </navTarget>
' and for pages:
'    <pageTarget value="integer" type="abstractnavnode@page" class="abstractnode.basename" id="ncxid" playorder="integer">
'      <pageLabel>
'        <text>abstractnode.text</text>
'      </pageLabel>
'      <content src="smilref" />
'    </pageTarget>


  If oAbstractNavNode.baseName = "pagenum" Then
    Set oMyNavList = oNcxDom.selectSingleNode("//pageList")
  Else
    Set oMyNavList = oNcxDom.selectSingleNode("//navList[@class='" & oAbstractNavNode.baseName & "']")
  End If
    
  If Not oMyNavList Is Nothing Then
    If oAbstractNavNode.baseName = "pagenum" Then
      Set oNewNavTarget = oBruno.oCmn.oDomCmn.fncAppendElement(oMyNavList, "pageTarget", "", "id", "nx" & oNcxIdGetter.fncGetId, "playOrder", oAbstractNavNode.selectSingleNode("@playOrder").Text)
      Set oNewNavLabel = oBruno.oCmn.oDomCmn.fncAppendElement(oNewNavTarget, "navLabel")
      Set oNewNavLabelText = oBruno.oCmn.oDomCmn.fncAppendElement(oNewNavLabel, "text")
      'create type attribute which is only for pages
      Set oNewTypeAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oNewNavTarget, "type", oAbstractNavNode.selectSingleNode("@page").Text)
      'try to create a numeric value of the text, if so, put into "value" attribute
      Dim lPageValue As Long
      lPageValue = fncMakePageTextNumeric(oAbstractNavNode.Text)
      If lPageValue > -1 Then
        Set oNewValueAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oNewNavTarget, "value", CStr(lPageValue))
      End If
    Else
      Set oNewNavTarget = oBruno.oCmn.oDomCmn.fncAppendElement(oMyNavList, "navTarget", "", "id", "nx" & oNcxIdGetter.fncGetId, "class", oAbstractNavNode.baseName, "playOrder", oAbstractNavNode.selectSingleNode("@playOrder").Text)
      Set oNewNavLabel = oBruno.oCmn.oDomCmn.fncAppendElement(oNewNavTarget, "navLabel")
      Set oNewNavLabelText = oBruno.oCmn.oDomCmn.fncAppendElement(oNewNavLabel, "text")
    End If
    
    'determine which navLabel to set (oNewNavLabelText.Text)
    'first check if an explicit wish has been set in driver
    'if so, the attribute 'bruno_navLabel' sits on oAbstractNavNode, with the text as value
    Dim oBrunoNavLabel As IXMLDOMAttribute
    Set oBrunoNavLabel = oAbstractNavNode.selectSingleNode("@bruno_navLabel")
    If Not (oBrunoNavLabel Is Nothing) Then
      'Stop
      oNewNavLabelText.Text = oBrunoNavLabel.Text
      ' ...then remove the temporary attribute
      oBruno.oCmn.oDomCmn.fncRemoveAttribute oAbstractNavNode, "bruno_navLabel"
    Else
      'there was no @bruno_navLabel at this oAbstractNavNode, fall back
      If oAbstractNavNode.Text = "" Then
        oNewNavLabelText.Text = oAbstractNavNode.baseName
      Else
        oNewNavLabelText.Text = oAbstractNavNode.Text
        If Len(oNewNavLabelText.Text) > 64 Then
          'truncate, split by space and build until > len 64
           Dim sTruncedLabel As String
           Dim sArray
           sArray = Split(oNewNavLabelText.Text, " ")
           Dim i As Long
           For i = 0 To UBound(sArray)
             sTruncedLabel = sTruncedLabel & " " & sArray(i)
             If Len(sTruncedLabel) > 64 Then Exit For
           Next i
           oNewNavLabelText.Text = Trim$(sTruncedLabel)
        End If
      End If
    End If
                
    Set oNewContent = oBruno.oCmn.oDomCmn.fncAppendElement(oNewNavTarget, "content", , "src", oAbstractNavNode.selectSingleNode("@smilref").Text)
  Else
    'navList not found
  End If
End Function

Private Function fncMakePageTextNumeric(sPageText As String) As Long
  On Error GoTo errh
  If IsNumeric(sPageText) Then
    fncMakePageTextNumeric = CLng(sPageText)
    Exit Function
  Else
    If fncIsRoman(sPageText) Then
      fncMakePageTextNumeric = CLng(GetRoman(sPageText))
      Exit Function
    Else
      fncMakePageTextNumeric = -1
      Exit Function
    End If
  End If
errh:
  fncMakePageTextNumeric = -1
End Function

Private Function fncCreateNavList(ByRef oAbstractNavNode As IXMLDOMNode, ByRef oNcxDom As MSXML2.DOMDocument40, ByRef oNcxIdGetter As cIdGetter) As Boolean
'creates a navlist element with a navlabel child
' appends it to ncx
Dim oNewNavList As IXMLDOMElement
Dim oNewNavLabel As IXMLDOMElement
Dim oNewNavLabelText As IXMLDOMElement
Dim sListName As String
'  <navList id="notes" class="note">
'    <navLabel>
'        <text>Notes</text>
'    </navLabel>
'  <navList>
'
'and for pages:
' <pageList>
'
' </pageList>

  If oAbstractNavNode.baseName = "pagenum" Then
    sListName = "pageList"
  Else
    sListName = "navList"
  End If
  Set oNewNavList = oBruno.oCmn.oDomCmn.fncAppendElement(oNcxDom.documentElement, sListName, , "id", "nx" & oNcxIdGetter.fncGetId, "class", oAbstractNavNode.baseName)
  Set oNewNavLabel = oBruno.oCmn.oDomCmn.fncAppendElement(oNewNavList, "navLabel")
  Set oNewNavLabelText = oBruno.oCmn.oDomCmn.fncAppendElement(oNewNavLabel, "text")
  oNewNavLabelText.Text = oAbstractNavNode.baseName
 ' Debug.Print oNcxDom.xml
    
End Function

Private Function fncNavListExistsForThisType(ByVal sType As String, ByRef oNcxDom As MSXML2.DOMDocument40) As Boolean
'the navlist is identified by its class attribute
'which is the same as the basename of the abstract node (=sType)
Dim oTestNode As IXMLDOMNode

  fncNavListExistsForThisType = False
  If sType = "pagenum" Then
    Set oTestNode = oNcxDom.selectSingleNode("//pageList")
  Else
    Set oTestNode = oNcxDom.selectSingleNode("//navList[@class='" & sType & "']")
  End If
  If Not oTestNode Is Nothing Then fncNavListExistsForThisType = True

End Function

Private Function fncAddMetaDataToNcc(oNccDom As MSXML2.DOMDocument40)
Dim oHeadnode As IXMLDOMNode
          
  Set oHeadnode = oNccDom.selectSingleNode("//xhtml:head")

    'these two are always set (may be "unkown")
    oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:identifier", "content", oBruno.oInputDocuments.oInputMetadata.sDcIdentifier
    oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:title", "content", oBruno.oInputDocuments.oInputMetadata.sDcTitle
    'below may be set (are else "")
    If oBruno.oInputDocuments.oInputMetadata.sDcCreator <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:creator", "content", oBruno.oInputDocuments.oInputMetadata.sDcCreator
    If oBruno.oInputDocuments.oInputMetadata.sDcPublisher <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:publisher", "content", oBruno.oInputDocuments.oInputMetadata.sDcPublisher
    If oBruno.oInputDocuments.oInputMetadata.sDcDate <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:date", "content", oBruno.oInputDocuments.oInputMetadata.sDcDate
    If oBruno.oInputDocuments.oInputMetadata.sDcSubject <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:subject", "content", oBruno.oInputDocuments.oInputMetadata.sDcSubject
    'added for niels 20050310
    If oBruno.oInputDocuments.oInputMetadata.sDcSource <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:source", "content", oBruno.oInputDocuments.oInputMetadata.sDcSource
    If oBruno.oInputDocuments.oInputMetadata.sNccMultimediaType <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "ncc:multimediaType", "content", oBruno.oInputDocuments.oInputMetadata.sNccMultimediaType
    If oBruno.oInputDocuments.oInputMetadata.sNccNarrator <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "ncc:narrator", "content", oBruno.oInputDocuments.oInputMetadata.sNccNarrator
    If oBruno.oInputDocuments.oInputMetadata.sNccSourceDate <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "ncc:sourceDate", "content", oBruno.oInputDocuments.oInputMetadata.sNccSourceDate
    If oBruno.oInputDocuments.oInputMetadata.sNccSourceEdition <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "ncc:sourceEdition", "content", oBruno.oInputDocuments.oInputMetadata.sNccSourceEdition
    If oBruno.oInputDocuments.oInputMetadata.sNccSourcePublisher <> "" Then oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "ncc:sourcePublisher", "content", oBruno.oInputDocuments.oInputMetadata.sNccSourcePublisher
    'end add neils
    If oBruno.oInputDocuments.oInputMetadata.sDcLanguage <> "" Then
      oBruno.oCmn.oDomCmn.fncAppendElement oHeadnode, "meta", , "name", "dc:language", "content", oBruno.oInputDocuments.oInputMetadata.sDcLanguage
      Dim oRootNode As IXMLDOMNode
      Set oRootNode = oNccDom.selectSingleNode("xhtml:html")
      oBruno.oCmn.oDomCmn.fncAppendAttribute oRootNode, "lang", oBruno.oInputDocuments.oInputMetadata.sDcLanguage
      oBruno.oCmn.oDomCmn.fncAppendAttribute oRootNode, "xml:lang", oBruno.oInputDocuments.oInputMetadata.sDcLanguage
    End If
    'Debug.Print oNccDom.xml
    'oNccDom.save ("D:\save.xml")
    'dtb-related metas:
    With oBruno.oCmn.oDomCmn
        'tocitems
        .fncAppendElement oHeadnode, "meta", , "name", "ncc:tocItems", "content", CStr(oNccDom.selectNodes("//xhtml:body/*").length)
        'span-pageNormal
        .fncAppendElement oHeadnode, "meta", , "name", "ncc:pageNormal", "content", CStr(oNccDom.selectNodes("//xhtml:body/xhtml:span[@class='page-normal']").length)
        'span-pageFront
        .fncAppendElement oHeadnode, "meta", , "name", "ncc:pageFront", "content", CStr(oNccDom.selectNodes("//xhtml:body/xhtml:span[@class='page-front']").length)
        'span-pageSpecial
        .fncAppendElement oHeadnode, "meta", , "name", "ncc:pageSpecial", "content", CStr(oNccDom.selectNodes("//xhtml:body/xhtml:span[@class='page-special']").length)
                
        'sidebar
        .fncAppendElement oHeadnode, "meta", , "name", "ncc:sidebars", "content", CStr(oNccDom.selectNodes("//xhtml:body/xhtml:span[@class='sidebar']").length)
        'prodnote
        .fncAppendElement oHeadnode, "meta", , "name", "ncc:prodNotes", "content", CStr(oNccDom.selectNodes("//xhtml:body/xhtml:span[@class='optional-prodnote']").length)
        'noteref
        .fncAppendElement oHeadnode, "meta", , "name", "ncc:footnotes", "content", CStr(oNccDom.selectNodes("//xhtml:body/xhtml:span[@class='noteref']").length)
                
        'charset
        .fncAppendElement oHeadnode, "meta", , "name", "ncc:charset", "content", oBruno.oInputDocuments.InputDocument(0).sEncoding
        'http-equiv
        .fncAppendElement oHeadnode, "meta", , "http-equiv", "Content-type", "content", "text/html; charset=" & oBruno.oInputDocuments.InputDocument(0).sEncoding, , , , , True
    End With

    oHeadnode.selectSingleNode("xhtml:title").Text = oBruno.oInputDocuments.oInputMetadata.sXhtmlTitle
  
End Function

Private Function fncSmilCustomTest(ByRef oSmilDom As MSXML2.DOMDocument40) As Boolean
  On Error GoTo errh
  fncSmilCustomTest = False
'  <smil>
'    <head>
'       <meta name="dtb:uid" content="uid" />
'       <meta name="dtb:generator" content="bruno" />
'       <meta name="dtb:totalElapsedTime" content="00:00:00.000" />
'       <customAttributes>
'          <customTest id="pagenum" defaultState="false" override="visible"/>
'          <customTest id="note" defaultState="true" override="visible"/>
'          <customTest id="prodnote" defaultState="true" override="visible"/>
'          <customTest id="linenum" defaultState="true" override="visible"/>
'          <customTest id="noteref" defaultState="true" override="visible"/>
'          <customTest id="annotation" defaultState="true" override="visible"/>
'          <customTest id="sidebar" defaultState="true" override="visible"/>
'       </customAttributes>

' check if there are seqs/pars in smil body who
' has a customTest attribute (added during abstract creation)
'
' if yes, add a customTest to customAttributes
' with an id equal to timecontainer customTest attr value

'Dim oCustomTestsInHead As IXMLDOMNodeList
Dim oCustomTestInHead As IXMLDOMNode
Dim oTimeContainersWithCustomTestAttrs As IXMLDOMNodeList
Dim oTimeContainer As IXMLDOMNode
Dim bCustomTestMatch As Boolean
Dim oCustomAttributesElem As IXMLDOMElement
Dim oTimeContainerCustomTestAttribute As IXMLDOMAttribute

  Set oCustomAttributesElem = oSmilDom.selectSingleNode("//smil2:customAttributes")
  'Set oCustomTestsInHead = oSmilDom.selectNodes("//smil2:customAttributes/smil2:customTest")
  Set oTimeContainersWithCustomTestAttrs = oSmilDom.selectNodes("//smil2:body//seq[@customTest]|//smil2:body//par[@customTest]")
  If (Not oTimeContainersWithCustomTestAttrs Is Nothing) Then
    'there were some customTest timecontainers
    For Each oTimeContainer In oTimeContainersWithCustomTestAttrs
      Set oTimeContainerCustomTestAttribute = oTimeContainer.selectSingleNode("@customTest")
      'check if this idref value has a match in head
      Set oCustomTestInHead = oCustomAttributesElem.selectSingleNode("customTest[@id='" & oTimeContainerCustomTestAttribute.Text & "']")
      If oCustomTestInHead Is Nothing Then
        'if not, add a new customTest to head
        Set oCustomTestInHead = oBruno.oCmn.oDomCmn.fncAppendElement(oCustomAttributesElem, "customTest", , "id", oTimeContainerCustomTestAttribute.Text, "defaultState", "true", "override", "visible")
        'add the customTest to the pan-smilfile oSmilCustomTests collection
        oSmilCustomTests.fncAddCustomTest oCustomTestInHead.cloneNode(True)
      Else
        'Stop
        'the head custom test already existed
      End If
    Next 'oTimeContainer In oTimeContainersWithCustomTestAttrs
    
'    For Each oCustomTest In oCustomTestsInHead
'      bCustomTestMatch = False
'      For Each oTimeContainer In oTimeContainersInBody
'        If oTimeContainer.selectSingleNode("@class").Text = oCustomTest.selectSingleNode("@id").Text Then
'          'there is a match in body, this customtest should remain in head
'          bCustomTestMatch = True
'          'append to customtest collection
'          oSmilCustomTests.fncAddCustomTest oCustomTest.cloneNode(True)
'          'add attr to time container
'          oBruno.oCmn.oDomCmn.fncAppendAttribute oTimeContainer, "customTest", oCustomTest.selectSingleNode("@id").Text
'        End If
'      Next oTimeContainer
'      If Not bCustomTestMatch Then Set oCustomTest = oCustomTest.parentNode.removeChild(oCustomTest)
'    Next oCustomTest
  End If '(Not oTimeContainersWithCustomTestAttrs Is Nothing)
  
  'if no customTests in this smilfile, remove empty customAttributes element
  If oCustomAttributesElem.childNodes.length = 0 Then
    Set oCustomAttributesElem = oCustomAttributesElem.parentNode.removeChild(oCustomAttributesElem)
  End If
  
  fncSmilCustomTest = True
errh:
End Function

Private Function fncCreateZedOpf() As MSXML2.DOMDocument40
Dim oOpfDom As New MSXML2.DOMDocument40
  oOpfDom.async = False
  oOpfDom.validateOnParse = False
  oOpfDom.resolveExternals = False
  oOpfDom.preserveWhiteSpace = False
  oOpfDom.setProperty "SelectionLanguage", "XPath"
  oOpfDom.setProperty "SelectionNamespaces", "xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:oeb='http://openebook.org/namespaces/oeb-package/1.0/'"
  oOpfDom.setProperty "NewParser", True
Dim i As Long
Dim oManifestNode As IXMLDOMNode
Dim oSpineNode As IXMLDOMNode
Dim oNode As IXMLDOMNode
Dim oContentIdGetter As cIdGetter
Dim oAuxIdGetter As cIdGetter

  On Error GoTo errh
   Set oContentIdGetter = New cIdGetter
   Set oAuxIdGetter = New cIdGetter
  
  If Not oBruno.oCmn.oDomCmn.fncParseFile(oBruno.oPaths.ShellPath & "z39opf.shell", oOpfDom, "") Then GoTo errh
  Set oManifestNode = oOpfDom.selectSingleNode("//oeb:manifest")
  Set oSpineNode = oOpfDom.selectSingleNode("//oeb:spine")
  
  'do the spine
  For i = 0 To oSpine.lSpineItemCount - 1
    oBruno.oCmn.oDomCmn.fncAppendElement _
      oSpineNode, "itemref", , _
      "idref", oSpine.fncGetSpineItem(i).selectSingleNode("@id").Text
  Next
  
  'do the manifest
  
  'add the opf itself
    oBruno.oCmn.oDomCmn.fncAppendElement _
      oManifestNode, "item", , _
      "href", oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension(oBruno.oAbstractDocuments.AbstractDocument(0).sFileName) & ".opf", _
      "id", "opf", _
      "media-type", "text/xml"
  
  'add the smilfiles from oSpine
  For i = 0 To oSpine.lSpineItemCount - 1
    oBruno.oCmn.oDomCmn.fncAppendElement _
      oManifestNode, "item", , _
      "href", oSpine.fncGetSpineItem(i).selectSingleNode("@name").Text, _
      "id", oSpine.fncGetSpineItem(i).selectSingleNode("@id").Text, _
      "media-type", "application/smil"
  Next
  
  'check if smilcustomtests, if so, add resourcefile to output array
  
  If oSmilCustomTests.bHasTests Then
   Dim oResDom As New MSXML2.DOMDocument40
   oResDom.async = False
   oResDom.validateOnParse = False
   oResDom.resolveExternals = False
   oResDom.preserveWhiteSpace = False
   oResDom.setProperty "NewParser", True
   oBruno.oCmn.oDomCmn.fncParseFile oBruno.oPaths.ShellPath & "z39res.shell", oResDom, ""
   'fncAddDocument oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension(oBruno.oInputDocuments.InputDocument(0).sFileName) & ".res", TYPE_ACTUAL_Z39RESOURCE, oResDom
   fncAddDocument oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension(oBruno.oAbstractDocuments.AbstractDocument(0).sFileName) & ".res", TYPE_ACTUAL_Z39RESOURCE, oResDom
  End If

 'add auxfiles
  Dim sMimeType As String, sFileExt As String
  Dim r As Long
  Dim bFilesetContainsImages As Boolean
  bFilesetContainsImages = False
  For r = 0 To oBruno.oInputDocuments.oInputAuxFiles.InputAuxFileCount - 1
    With oBruno.oInputDocuments.oInputAuxFiles.InputAuxFile(r)
      sFileExt = oBruno.oCmn.oFsoCmn.fncGetExtensionFromString(.sFullPath)
      Select Case sFileExt
         Case "mp4"
           sMimeType = "audio/mpeg4-generic"
         Case "mp3"
           sMimeType = "audio/mpeg"
         Case "wav"
           sMimeType = "audio/x-wav"
         Case "jpg"
           sMimeType = "image/jpeg"
           bFilesetContainsImages = True
         Case "png"
           sMimeType = "image/png"
           bFilesetContainsImages = True
         Case "svg"
           sMimeType = "image/svg+xml"
           bFilesetContainsImages = True
         Case "css"
           sMimeType = "text/css"
         Case Else
           sMimeType = "mime/unkown"
        End Select
        
        oBruno.oCmn.oDomCmn.fncAppendElement _
        oManifestNode, "item", , _
        "href", .sFileName, _
        "id", "aux" & oAuxIdGetter.fncGetId, _
        "media-type", sMimeType
     End With
  Next
  
  'now we know if the input dtbook contained images.
  'if it did not, mod the multimediaContent value that was preset in the shell as "text,image"
  If Not bFilesetContainsImages Then
    Set oNode = oOpfDom.selectSingleNode("//oeb:x-metadata/oeb:meta[@name='dtb:multimediaContent']/@content")
    If Not oNode Is Nothing Then
      oNode.Text = "text"
    End If
  End If
  'add the rest from aoutputdocuments
  
  For r = 0 To lOutputDocumentCount - 1
    If (aOutputDocuments(r).lType <> TYPE_ACTUAL_Z39SMIL) Then
      If (aOutputDocuments(r).lType = TYPE_ACTUAL_Z39NCX) Then
        oBruno.oCmn.oDomCmn.fncAppendElement _
        oManifestNode, "item", , _
        "href", aOutputDocuments(r).sFileName, _
        "id", "ncx", _
        "media-type", "application/x-dtbncx+xml"
      ElseIf (aOutputDocuments(r).lType = TYPE_ACTUAL_Z39OPF) Then
        oBruno.oCmn.oDomCmn.fncAppendElement _
        oManifestNode, "item", , _
        "href", aOutputDocuments(r).sFileName, _
        "id", "opf", _
        "media-type", "text/xml"
      ElseIf (aOutputDocuments(r).lType = TYPE_ACTUAL_Z39CONTENT) Then
        oBruno.oCmn.oDomCmn.fncAppendElement _
        oManifestNode, "item", , _
        "href", aOutputDocuments(r).sFileName, _
        "id", "xml" & oContentIdGetter.fncGetId, _
        "media-type", "application/x-dtbook+xml"
      ElseIf (aOutputDocuments(r).lType = TYPE_ACTUAL_Z39RESOURCE) Then
        oBruno.oCmn.oDomCmn.fncAppendElement _
        oManifestNode, "item", , _
        "href", aOutputDocuments(r).sFileName, _
        "id", "resource", _
        "media-type", "application/x-dtbresource+xml"
      ElseIf (aOutputDocuments(r).lType = TYPE_ACTUAL_AUXILLIARY) Then
        MsgBox "shouldnt be here: outputdocuments[] has auxfiles"
      Else
       'unknown
      End If
    
    End If
  Next
  
  'add some metas
'  Set oNode = oOpfDom.selectSingleNode("//oeb:package/@unique-identifier")
'  If Not oNode Is Nothing Then
'    oNode.Text = oBruno.oInputDocuments.oInputMetadata.sDtbookUid
'  End If
    
  Set oNode = oOpfDom.selectSingleNode("//oeb:dc-metadata/dc:Title")
  If Not oNode Is Nothing Then
    oNode.Text = oBruno.oInputDocuments.oInputMetadata.sDcTitle
  End If
    
  Set oNode = oOpfDom.selectSingleNode("//oeb:dc-metadata/dc:Creator")
  If Not oNode Is Nothing Then
    oNode.Text = oBruno.oInputDocuments.oInputMetadata.sDcCreator
  End If
    
  Set oNode = oOpfDom.selectSingleNode("//oeb:dc-metadata/dc:Language")
  If Not oNode Is Nothing Then

  Else
    'go for xml:lang
  End If
    
  Set oNode = oOpfDom.selectSingleNode("//oeb:dc-metadata/dc:Publisher")
  If Not oNode Is Nothing Then
    oNode.Text = oBruno.oInputDocuments.oInputMetadata.sDcPublisher
  End If
    
  Set oNode = oOpfDom.selectSingleNode("//oeb:dc-metadata/dc:Date")
  If Not oNode Is Nothing Then
    oNode.Text = oBruno.oInputDocuments.oInputMetadata.sDcDate
  End If
  
  Set oNode = oOpfDom.selectSingleNode("//oeb:dc-metadata/dc:Identifier")
  If Not oNode Is Nothing Then
    oNode.Text = oBruno.oInputDocuments.oInputMetadata.sDtbookUid
  End If
        
  Set oNode = oOpfDom.selectSingleNode("//oeb:x-metadata/oeb:meta[@name='prod:generator']/@content")
  If Not oNode Is Nothing Then oNode.Text = oBruno.sAppVersion
      
  Set fncCreateZedOpf = oOpfDom

  Exit Function
errh:
  Set oOpfDom = Nothing
  Set oContentIdGetter = Nothing
  Set oAuxIdGetter = Nothing
  Set oManifestNode = Nothing
  Set oSpineNode = Nothing
  Set oNode = Nothing
End Function

Private Function fncCreateLpp() As Boolean
Dim db As New ADODB.Connection
'Dim db As ADODB.Connection
Dim RsSel As ADODB.Recordset
Dim RsUpd As ADODB.Recordset
Dim sQuery As String, sUpdate As String
Const cProjectDbPwd = "20KHhl00"
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim i As Long, k As Long
Dim bEmptyWavAdded As Boolean
    
    On Error GoTo errh
    
    bEmptyWavAdded = False
    
    'make a temp copy of the shell file
    fso.CopyFile oBruno.oPaths.ShellPath & "pro.lpp", oBruno.oPaths.ShellPath & "temp.lpp", True

    With db
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        '.Provider = "Microsoft.Jet.OLEDB.3.51."
        .Properties("Jet OLEDB:Database Password") = cProjectDbPwd
        .Mode = adModeReadWrite
        .Properties("Data Source") = oBruno.oPaths.ShellPath & "temp.lpp"
        
        '.Open oBruno.oPaths.ShellPath & "temp.lpp"
        .open
    End With
    
    If Err.Number <> 0 Then
        frmMain.fncAddMessage "Error opening Project Database!" & vbCrLf & Err.Number & " - " & Err.Description
        Err.Clear
    Else
        'set ncc items
        Dim lNccElemCount As Long
        For i = 0 To lOutputDocumentCount - 1
          If aOutputDocuments(i).sFileName = "ncc.html" Then
            lNccElemCount = aOutputDocuments(i).oDom.selectNodes("//xhtml:body/*").length
            Exit For
          End If
        Next i
        sUpdate = "UPDATE ProjectSettings SET FirstNccItem = '1'"
        fncDbExecute db, sUpdate, RsUpd
        sUpdate = "UPDATE ProjectSettings SET LastNccItem='" & lNccElemCount & "'"
        fncDbExecute db, sUpdate, RsUpd
        
        
        '20040616: rewrite of fileusage, same order of files as Pro
        'set fileusage
        'delete all thats in there
        sQuery = "Delete * From FileUsage"
        Set RsSel = db.Execute(sQuery)
        
        Dim sLppFileName As String
        sLppFileName = oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension(oBruno.oInputDocuments.InputDocument(0).sFileName)
                
        For i = 0 To lOutputDocumentCount - 1
          'first add smilfiles
          If aOutputDocuments(i).lType = TYPE_ACTUAL_202SMIL Then
            sUpdate = "INSERT INTO FileUsage VALUES("
            sUpdate = sUpdate & Chr(39) & aOutputDocuments(i).sFileName & Chr(39) & ",'" & sLppFileName & ".lpp',"
            sUpdate = sUpdate & Chr(39) & "smil',"
            ''sUpdate = sUpdate & Chr(39) & Date & " " & Time & Chr(39) & ",0)"
            sUpdate = sUpdate & fncGetDateAndTime & ",0)"
            fncDbExecute db, sUpdate, RsUpd
            'if this was the first smilfile, insert empty wav reference after it
            If Not bEmptyWavAdded Then
              sUpdate = "INSERT INTO FileUsage VALUES("
              sUpdate = sUpdate & Chr(39) & "brun0001.wav" & Chr(39) & ",'none',"
              'sUpdate = sUpdate & Chr(39) & oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension(aOutputDocuments(i).sFileName) & ".wav" & Chr(39) & ",'none',"
              sUpdate = sUpdate & Chr(39) & "audio',"
              sUpdate = sUpdate & fncGetDateAndTime & ",0)"
              fncDbExecute db, sUpdate, RsUpd
              bEmptyWavAdded = True
            End If
          End If
        Next i
        'now add default mdf
        sUpdate = "INSERT INTO FileUsage VALUES('default.mdf','" & sLppFileName & ".lpp' ,'mdf'," & fncGetDateAndTime & ",0)"
        fncDbExecute db, sUpdate, RsUpd
                
        'now add mastersmil
        sUpdate = "INSERT INTO FileUsage VALUES('master.smil','" & sLppFileName & ".lpp' ,'master'," & fncGetDateAndTime & ",0)"
        fncDbExecute db, sUpdate, RsUpd

        'now add ncc
        sUpdate = "INSERT INTO FileUsage VALUES('ncc.html','" & sLppFileName & ".lpp' ,'ncc'," & fncGetDateAndTime & ",0)"
        fncDbExecute db, sUpdate, RsUpd

        'now add text docs
        For i = 0 To lOutputDocumentCount - 1
          If aOutputDocuments(i).lType = TYPE_ACTUAL_202CONTENT Then
            k = k + 1
            sUpdate = "INSERT INTO FileUsage VALUES("
            sUpdate = sUpdate & Chr(39) & aOutputDocuments(i).sFileName & Chr(39) & ",'" & sLppFileName & ".lpp' ,"
            sUpdate = sUpdate & Chr(39) & "doc0" & CStr(k) & "',"
            sUpdate = sUpdate & fncGetDateAndTime & ",0)"
            fncDbExecute db, sUpdate, RsUpd
          End If
        Next i

        'add auxilliary files if existing
        'debug:
        'For i = 0 To oBruno.oInputDocuments.oInputAuxFiles.InputAuxFileCount - 1
        '  Debug.Print oBruno.oInputDocuments.oInputAuxFiles.InputAuxFile(i).sFileName
        'Next i
        'Stop
                
        For i = 0 To oBruno.oInputDocuments.oInputAuxFiles.InputAuxFileCount - 1
          sUpdate = "INSERT INTO FileUsage VALUES("
          sUpdate = sUpdate & Chr(39) & oBruno.oInputDocuments.oInputAuxFiles.InputAuxFile(i).sFileName & Chr(39) & ",'" & sLppFileName & ".lpp' ,"
          sUpdate = sUpdate & Chr(39) & "ref',"
          sUpdate = sUpdate & fncGetDateAndTime & ",0)"
          fncDbExecute db, sUpdate, RsUpd
        Next
                        
        'add meta

        sQuery = "Delete * From meta"
        Set RsSel = db.Execute(sQuery)

        'Debug.Print "insert identifier"
        sUpdate = "INSERT INTO meta (name, content) VALUES ('dc:identifier','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sDcIdentifier) & "')"
        fncDbExecute db, sUpdate, RsUpd

        sUpdate = "INSERT INTO meta (name, content) VALUES ('dc:title','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sDcTitle) & "')"
        fncDbExecute db, sUpdate, RsUpd
        
        'below are not always set
        If oBruno.oInputDocuments.oInputMetadata.sDcCreator <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('dc:creator','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sDcCreator) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        If oBruno.oInputDocuments.oInputMetadata.sDcDate <> "" Then
          sUpdate = "INSERT INTO meta (name, content, scheme) VALUES ('dc:date','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sDcDate) & "','yyyy-mm-dd')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        If oBruno.oInputDocuments.oInputMetadata.sDcLanguage <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('dc:language','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sDcLanguage) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        If oBruno.oInputDocuments.oInputMetadata.sDcPublisher <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('dc:publisher','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sDcPublisher) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        If oBruno.oInputDocuments.oInputMetadata.sDcSubject <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('dc:subject','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sDcSubject) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        'added for niels 20050510
        
        If oBruno.oInputDocuments.oInputMetadata.sDcSource <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('dc:source','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sDcSource) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        If oBruno.oInputDocuments.oInputMetadata.sNccMultimediaType <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('ncc:multimediaType','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sNccMultimediaType) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        If oBruno.oInputDocuments.oInputMetadata.sNccNarrator <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('ncc:narrator','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sNccNarrator) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        If oBruno.oInputDocuments.oInputMetadata.sNccSourceDate <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('ncc:sourceDate','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sNccSourceDate) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        If oBruno.oInputDocuments.oInputMetadata.sNccSourceEdition <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('ncc:sourceEdition','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sNccSourceEdition) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        If oBruno.oInputDocuments.oInputMetadata.sNccSourcePublisher <> "" Then
          sUpdate = "INSERT INTO meta (name, content) VALUES ('ncc:sourcePublisher','" & fncEscapeSql(oBruno.oInputDocuments.oInputMetadata.sNccSourcePublisher) & "')"
          fncDbExecute db, sUpdate, RsUpd
        End If
        
        'end niels add
        
        sUpdate = "INSERT INTO meta (name, content) VALUES ('ncc:charset','" & oBruno.oInputDocuments.InputDocument(0).sEncoding & "')"
        fncDbExecute db, sUpdate, RsUpd
                
'        'code to look at result
'test:
'        sQuery = "Select * From FileUsage"
'        Set RsSel = Db.Execute(sQuery)
'
'        If Not RsSel.EOF Then
'          Do Until RsSel.EOF
'            Debug.Print RsSel("SourceFile") & " -> " & RsSel("RefereceFile") & " [" & RsSel("LastWritten") & "]" & vbCrLf
'            RsSel.MoveNext
'          Loop
'        End If
'        Stop
    End If

    
    db.Close
    Set RsSel = Nothing
    Set RsUpd = Nothing
    Set db = Nothing
    Exit Function
errh:
    MsgBox "errhandler caught an exception:" & Err.Description & Err.LastDllError & Err.Number & Err.Source & " Note: lpp file is unusable" & vbCrLf & "String: " & sUpdate
    
End Function

Private Function fncEscapeSql(ByVal str As String) As String
 'fncEscapeSql = Replace(str, "'", "&#39;")
 
 'str = Replace(str, "\", "\\")
 str = Replace(str, "'", "''")
 'str = Replace(str, "_", "\_")
 'str = Replace(str, "%", "\%")
 'str = Replace(str, Chr(34), Chr(92) & Chr(34))
    
 fncEscapeSql = str
End Function

Private Function fncDbExecute(ByRef db As ADODB.Connection, ByRef sUpdate As String, ByRef RsUpd As ADODB.Recordset)
  'Debug.Print sUpdate
  Do Until (db.State <> adStateExecuting) And (db.State <> adStateFetching)
    Debug.Print "Change command executing...."
  Loop
  Set RsUpd = db.Execute(sUpdate)
  sUpdate = ""
End Function

Sub Wait(ByVal s As Integer)
Dim t As Double
t = Now() + s / 86400#
While t > Now()
DoEvents
Wend
End Sub

Private Function fncGetDateAndTime() As String
 fncGetDateAndTime = Chr(39) & Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & Chr(39)
End Function

Private Function fncMakeNumeric(ByRef sString As String) As Long
Dim i As Long, sCh As String
  On Error GoTo errhandler
    
  Do
    For i = 0 To Len(sString)
      sCh = Mid(sString, i + 1, 1)
      If Not IsNumeric(sCh) Then
        sString = Replace$(sString, sCh, "")
      End If
    Next
  Loop Until IsNumeric(sString) Or Len(sString) = 0

  fncMakeNumeric = CLng(sString)
  
errhandler:
  
End Function

Private Function fncIsRoman(ByVal sString As String) As Boolean
'This function is given a string and
'returns true if it only contains
'valid roman numerals, false otherwise.
Dim i As Long
Dim sChar As String
  
  For i = 1 To Len(sString)
    sChar = Mid(sString, i, 1)
    'Convert digit to UpperCase
    sChar = UCase(sChar)
    'Test the digit
    Select Case sChar
      Case "M", "D", "C", "L", "X", "V", "I"
      fncIsRoman = True
    Case Else
      fncIsRoman = False
    End Select
  Next i

End Function
Private Function fncConvertRoman(ByVal numr As String) As String
'input: a roman numeral
'returns its value.
'NULL is returned if the character is not valid

Dim digit As Integer
 'Convert digit to UpperCase
 numr = UCase(numr)
 'Convert the digit
 Select Case numr
   Case "M"
    digit = M
   Case "D"
     digit = D
   Case "C"
     digit = C
   Case "L"
     digit = L
   Case "X"
     digit = X
   Case "V"
     digit = V
   Case "I"
     digit = i
   Case Else
     digit = vbNull
End Select
  'And return its value
  fncConvertRoman = digit
End Function

Public Function GetRoman(ByVal numr As String) As String
'reads the next number in
'roman numerals from the input
'and returns it as an integer
Dim rdigit As String
Dim num As Long
Dim DigValue As Long
Dim LastDigValue As String
Dim j As Long
  j = 1
  num = 0
  LastDigValue = M
  'Get the first digit
  rdigit = Mid(numr, j, 1)
  'While it is a roman digit
  Do While fncIsRoman(rdigit)
    'Convert roman digit to its value
    DigValue = fncConvertRoman(rdigit)
    'If previous digit was a prefix digit
    If DigValue > LastDigValue Then
      'Adjust total
      num = num - 2 * LastDigValue + DigValue
    Else
      'Otherwise accumulate the total
       num = num + DigValue
      'Save this digit as previous
       LastDigValue = DigValue
    End If
    'Get next digit
    j = j + 1
    rdigit = Mid(numr, j, 1)
    'End of the string detected, exit
    If Len(rdigit) = 0 Then
      Exit Do
    End If
  Loop
  'Return the number
  GetRoman = num
End Function

Private Function fncSetContentDocTitle(oTmpDom As MSXML2.DOMDocument40) As Boolean
Dim oTitleElement As IXMLDOMNode
 On Error GoTo errh
 
 fncSetContentDocTitle = False
 
 Set oTitleElement = oTmpDom.selectSingleNode("//xhtml:head/xhtml:title")
 If Not oTitleElement Is Nothing Then
   If Trim$(oTitleElement.Text) = "" Then
     'this has been set to 'something' when cInputMetadata is instantiated
     oTitleElement.Text = oBruno.oInputDocuments.oInputMetadata.sXhtmlTitle
   End If
 End If
  
 fncSetContentDocTitle = True
errh:

End Function
