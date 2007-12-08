VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDomCmn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
  'debug.Print "cDomCmn.initialize"
End Sub

Private Sub Class_Terminate()
  'debug.Print "cDomCmn.terminate"
End Sub

Public Function fncParseFile( _
    ByVal isAbsPath As String, _
    ByRef ioDom As MSXML2.DOMDocument40, _
    ByRef sReturnMessage As String _
    ) As Boolean
    
    fncParseFile = False
    
    If Not ioDom.Load(isAbsPath) Then
        sReturnMessage = "Parse error in " & oBruno.oCmn.oFsoCmn.fncGetFileName(isAbsPath) & ": " & ioDom.parseError.reason & vbCrLf & _
           "filepos: " & ioDom.parseError.filepos & vbCrLf & _
           "line: " & ioDom.parseError.Line & vbCrLf & _
           "linepos: " & ioDom.parseError.linepos & vbCrLf & _
           "srctext: " & ioDom.parseError.srcText & vbCrLf & _
           "Url: " & ioDom.parseError.url
    Else
        fncParseFile = True
    End If
    
End Function

Public Function fncParseString( _
    ByVal isContent As String, _
    ByRef ioDom As MSXML2.DOMDocument40, _
    ByRef sReturnMessage As String _
    ) As Boolean
    
    fncParseString = False

    If Not ioDom.loadXML(isContent) Then
        sReturnMessage = "Parse error: " & ioDom.parseError.reason & vbCrLf & _
           "filepos: " & ioDom.parseError.filepos & vbCrLf & _
           "line: " & ioDom.parseError.Line & vbCrLf & _
           "linepos: " & ioDom.parseError.linepos & vbCrLf & _
           "srctext: " & ioDom.parseError.srcText & vbCrLf & _
           "Url: " & ioDom.parseError.url
    Else
        fncParseString = True
    End If
End Function

Public Function fncAppendAttribute( _
    ByRef oParentElem As IXMLDOMElement, _
    ByVal sAttrName As String, _
    ByVal sAttrValue As String _
    ) As IXMLDOMAttribute
Dim oNewAttr As IXMLDOMNode
Dim oNamedNodeMap As IXMLDOMNamedNodeMap

  On Error GoTo errhandler
  If sAttrName <> "" Then
    Set oNewAttr = oParentElem.ownerDocument.createNode(NODE_ATTRIBUTE, sAttrName, "")
    oNewAttr.nodeTypedValue = sAttrValue
    Set oNamedNodeMap = oParentElem.Attributes
    oNamedNodeMap.setNamedItem oNewAttr
  End If

  Set fncAppendAttribute = oNewAttr
errhandler:
End Function

'Public Function fncCheckValidityOld(oInputDom As MSXML2.DOMDocument40, lDoctype As Long) As Boolean
'Dim oValDom As New MSXML2.DOMDocument40
'    oValDom.async = False
'    oValDom.validateOnParse = True
'    oValDom.resolveExternals = True
'    oValDom.preserveWhiteSpace = False
'    oValDom.setProperty "NewParser", False
'
'Dim sTempDom As String
'Dim sTempUri As String, sTempProlog As String, sRootElemName
'
'  sTempDom = oInputDom.xml
'  'create the specific dtd path
'
'  sTempUri = "file:///" & oBruno.oPaths.DtdPath
'
'  sRootElemName = "html"
'  If lDoctype = DOCTYPE_XHTML1_TRANSITIONAL Then
'    sTempUri = sTempUri & "xhtml1-transitional.dtd"
'  ElseIf lDoctype = DOCTYPE_XHTML1_STRICT Then
'    sTempUri = sTempUri & "xhtml1-strict.dtd"
'  ElseIf lDoctype = DOCTYPE_DTBOOK Then
'    sTempUri = sTempUri & "dtbook110.dtd"
'    sRootElemName = "dtbook"
'  Else
'    sTempUri = sTempUri & "confused.dtd"
'  End If
'
'  'get rid of backslashes
'  sTempUri = Replace$(sTempUri, "\", "/")
'
'  'create a new prolog
'
'   sTempProlog = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>" & vbCrLf & _
'       "<!DOCTYPE " & sRootElemName & " SYSTEM " & Chr(34) & sTempUri & Chr(34) & ">" & vbCrLf
'
'  'merge
'  sTempDom = sTempProlog & Mid(sTempDom, InStr(1, sTempDom, "<" & sRootElemName))
'
'  'validate
'  fncCheckValidityOld = fncParseString(sTempDom, oValDom)
'  Set oValDom = Nothing
'
'End Function

'Public Function fncCheckValidity(ByRef oInputDom As MSXML2.DOMDocument40, ByVal sLocalDtdPath As String, ByRef sXmlParseMessage As String) As Boolean

Public Function fncCheckValidity(ByRef oInputDom As MSXML2.DOMDocument40, ByRef oDriver As cDriver, ByRef sXmlParseMessage As String) As Boolean
Dim oValDom As New MSXML2.DOMDocument40
    oValDom.async = False
    oValDom.validateOnParse = True
    oValDom.resolveExternals = True
    oValDom.preserveWhiteSpace = False
    oValDom.setProperty "NewParser", False
Dim sTempDom As String
Dim sTempUri As String, sTempProlog As String, sRootElemName As String, sDtdToUse As String
  
  Dim bLocalDTDSet As Boolean
  
  'check for 202 skippable dtd extenstion
  If oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
    If bDTDIsExtendedInternallyForSkippableStructures(oInputDom.xml) Then
      oDriver.fncSetLocalDtdPath "xhtml1-transitional-bodyref.dtd"
      bLocalDTDSet = True
    Else
      'continue to use the default dtd given in driver
    End If
  End If
    
  If bLocalDTDSet = False Then
    'create the specific dtd path
    Dim sLocalDtdPath As String
    sLocalDtdPath = oDriver.sLocalDtdPath
    'there may be several pipeseparated DTDs mentioned in the driver
    If InStr(1, sLocalDtdPath, "|") > 0 Then
      Dim aDTDs() As String
      aDTDs = Split(sLocalDtdPath, "|")
      Dim k As Long
      For k = 0 To UBound(aDTDs)
        If InStr(1, oInputDom.doctype.xml, aDTDs(k)) > 0 Then
          oDriver.fncSetLocalDtdPath aDTDs(k)
         Exit For
       End If
      Next k
    Else
      oDriver.fncSetLocalDtdPath sLocalDtdPath
    End If
  End If
    
  sRootElemName = oInputDom.documentElement.nodeName
  sTempDom = oInputDom.xml
    
  sTempUri = Replace$("file:///" & oDriver.sLocalDtdPath, "\", "/")
  'Debug.Print sTempUri
  'create a new prolog
   sTempProlog = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>" & vbCrLf & _
       "<!DOCTYPE " & sRootElemName & " SYSTEM " & Chr(34) & sTempUri & Chr(34) & ">" & vbCrLf
  'merge
  sTempDom = sTempProlog & Mid(sTempDom, InStr(1, sTempDom, "<" & sRootElemName))
  
  'validate
  If fncParseString(sTempDom, oValDom, sXmlParseMessage) Then
    fncCheckValidity = True
  Else
    fncCheckValidity = False
  End If
  
  'check for 202 h1.title
  If oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
    Dim oFirstBodyChild As IXMLDOMNode
    Dim oFirstBodyChildAtt As IXMLDOMAttribute
    'TODO
    Set oFirstBodyChild = oValDom.documentElement.lastChild.firstChild
    If oFirstBodyChild.nodeName = "h1" Then
      Set oFirstBodyChildAtt = oFirstBodyChild.Attributes.getNamedItem("class")
      If Not oFirstBodyChildAtt Is Nothing Then
        If oFirstBodyChildAtt.nodeValue <> "title" Then fncCheckValidity = False
      Else
        fncCheckValidity = False
      End If
    Else
      fncCheckValidity = False
    End If
    
    If Not fncCheckValidity Then frmMain.fncAddMessage "first body element is not a h1 class title"
    
  ElseIf oDriver.lOutFileSet = OUTPUT_TYPE_Z39 Then
    Dim oMetaNodes As IXMLDOMNodeList
    Dim oMeta As IXMLDOMNode
    Dim bHadMetaDtbUid As Boolean: bHadMetaDtbUid = False
    Dim bHadMetaDcTitle As Boolean: bHadMetaDcTitle = False
        
    Set oMetaNodes = oValDom.documentElement.firstChild.childNodes
    For Each oMeta In oMetaNodes
      If oMeta.Attributes.getNamedItem("name").Text = "dtb:uid" Then bHadMetaDtbUid = True
      If oMeta.Attributes.getNamedItem("name").Text = "dc:Title" Then bHadMetaDcTitle = True
    Next oMeta
    
    If Not bHadMetaDtbUid Then
      fncCheckValidity = False
      frmMain.fncAddMessage "input dtbook document does not have the required dtb:uid meta element"
    End If
    
    If Not bHadMetaDcTitle Then
      fncCheckValidity = False
      frmMain.fncAddMessage "input dtbook document does not have the required dc:Title meta element"
    End If
    
    
  End If

  
  Set oValDom = Nothing
  
End Function

Public Function fncAppendElement( _
    ByRef oParentNode As IXMLDOMElement, _
    ByVal sName As String, _
    Optional ByVal sTextNode As String, _
    Optional ByVal sAttr1Name As String, _
    Optional ByVal sAttr1Value As String, _
    Optional ByVal sAttr2Name As String, _
    Optional ByVal sAttr2Value As String, _
    Optional ByVal sAttr3Name As String, _
    Optional ByVal sAttr3Value As String, _
    Optional ByVal sAttr4Name As String, _
    Optional ByVal sAttr4Value As String, _
    Optional bInsertBefore As Boolean, _
    Optional sNamespaceUri As String _
    ) As IXMLDOMNode

Dim oNewNode As IXMLDOMNode
Dim oNewAttr As IXMLDOMNode
Dim oNamedNodeMap As IXMLDOMNamedNodeMap
    
 '   On Error GoTo ErrHandler

    Set oNewNode = oParentNode.ownerDocument.createNode(NODE_ELEMENT, sName, sNamespaceUri)
    
    If sTextNode <> "" Then oNewNode.nodeTypedValue = sTextNode
    
    If sAttr1Name <> "" Then
        Set oNewAttr = oParentNode.ownerDocument.createNode(NODE_ATTRIBUTE, sAttr1Name, "")
        oNewAttr.nodeTypedValue = sAttr1Value
        Set oNamedNodeMap = oNewNode.Attributes
        oNamedNodeMap.setNamedItem oNewAttr
    End If
    
    If sAttr2Name <> "" Then
        Set oNewAttr = oParentNode.ownerDocument.createNode(NODE_ATTRIBUTE, sAttr2Name, "")
        oNewAttr.nodeTypedValue = sAttr2Value
        Set oNamedNodeMap = oNewNode.Attributes
        oNamedNodeMap.setNamedItem oNewAttr
    End If
    
    If sAttr3Name <> "" Then
        Set oNewAttr = oParentNode.ownerDocument.createNode(NODE_ATTRIBUTE, sAttr3Name, "")
        oNewAttr.nodeTypedValue = sAttr3Value
        Set oNamedNodeMap = oNewNode.Attributes
        oNamedNodeMap.setNamedItem oNewAttr
    End If
    
    If sAttr4Name <> "" Then
        Set oNewAttr = oParentNode.ownerDocument.createNode(NODE_ATTRIBUTE, sAttr4Name, "")
        oNewAttr.nodeTypedValue = sAttr4Value
        Set oNamedNodeMap = oNewNode.Attributes
        oNamedNodeMap.setNamedItem oNewAttr
    End If
    
    If bInsertBefore And oParentNode.hasChildNodes Then
      Set oNewNode = oParentNode.insertBefore(oNewNode, oParentNode.firstChild)
    Else
      Set oNewNode = oParentNode.appendChild(oNewNode)
    End If
    Set fncAppendElement = oNewNode
    
errhandler:
End Function

Public Function fncRemoveAllChildElements(ByRef oNode As IXMLDOMNode) As Boolean
Dim oChildElems As IXMLDOMNodeList
Dim oRChild As IXMLDOMNode
Dim i As Long
  
  If oNode Is Nothing Then
    fncRemoveAllChildElements = True
    Exit Function
  End If
  
  fncRemoveAllChildElements = False
  
  'remove all child elements of input node
  Set oChildElems = oNode.selectNodes(".//*")
  For i = 0 To oChildElems.length - 1
    'i=1 because first node in list is context node
    If oChildElems.Item(i).nodeType = NODE_ELEMENT Then
      'debug.Print oChildElems.Item(i).nodeName
      Set oRChild = oChildElems.Item(i).parentNode.removeChild(oChildElems.Item(i))
    End If
  Next
  fncRemoveAllChildElements = True

End Function

'Public Function fncGetAllTextFromContextAndDescendants(ByRef oInNode As IXMLDOMNode) As String
'Dim oChildElements As IXMLDOMNodeList
'Dim oInNode As IXMLDOMNode
'Dim sTextToReturn'
'
' 'gets a node as input
' 'returns its text, and all childelement text as well
' Set oChildElements = oInNode.selectNodes(".//*")
' For Each oInNode In oChildElements'
''
'
'End Function

Public Function fncRemoveNodesInNodeList(ByRef oNodeList As IXMLDOMNodeList) As Boolean
Dim oNode As IXMLDOMNode
Dim oRemoveNode As IXMLDOMNode
     For Each oNode In oNodeList
       If oNode.nodeType = NODE_ELEMENT Then
         Set oRemoveNode = oNode.parentNode.removeChild(oNode)
       ElseIf oNode.nodeType = NODE_ATTRIBUTE Then
         Set oRemoveNode = oNode.selectSingleNode("..")
         Set oRemoveNode = oRemoveNode.Attributes.removeNamedItem(oNode.nodeName)
       End If
     Next
End Function

Public Function fncXmlIndentXsl( _
  ByRef oDom As MSXML2.DOMDocument40 _
  ) As Boolean
Dim oXsl As New MSXML2.FreeThreadedDOMDocument40
    oXsl.async = False
Dim oXslTemplate As New MSXML2.XSLTemplate40
Dim oXslProc As IXSLProcessor
    'parse the xsl
    If oBruno.oCmn.oDomCmn.fncParseFile(oBruno.oPaths.ShellPath & "ws.xsl", oXsl, "") Then
      Set oXslTemplate.stylesheet = oXsl
      Set oXslProc = oXslTemplate.createProcessor()
      oXslProc.input = oDom
      oXslProc.Transform
      Stop
      oDom.preserveWhiteSpace = True
      If oBruno.oCmn.oDomCmn.fncParseString(oXslProc.output, oDom, "") Then
        'all ok
      Else
        'couldnt parse xsl output
      End If
    Else
      'couldnt parse xsl
    End If
End Function

Public Function fncXmlIndentDom(ByRef oElem As IXMLDOMNode, sIndent As String) As Boolean
'pretty prints a dom tree, handles anchors differently
'shall have documentelement as intitial input
Dim oChild As IXMLDOMNode
Dim oNew As IXMLDOMNode
Dim oChildElems As IXMLDOMNodeList
Static indentOrigLen As Integer
    
    fncXmlIndentDom = False
    
    If indentOrigLen = 0 Then indentOrigLen = Len(sIndent)
    
    Set oChildElems = oElem.selectNodes("*")
    
    If Not oChildElems Is Nothing Then
      If oChildElems.length > 0 Then
        For Each oChild In oChildElems
            fncXmlIndentDom oChild, sIndent & Left$(sIndent, indentOrigLen)
            If (oElem.nodeName <> "a") And (Not fncHasChildNodeWithLocalName(oElem, "a")) Then
                Set oNew = oElem.ownerDocument.createNode(NODE_TEXT, vbNullString, vbNullString)
                oNew.nodeValue = vbCrLf & sIndent
                Set oNew = oElem.insertBefore(oNew, oChild)
                Set oNew = Nothing
            End If
        Next
        If (oElem.nodeName <> "a") And (Not fncHasChildNodeWithLocalName(oElem, "a")) Then
            Set oNew = oElem.ownerDocument.createNode(NODE_TEXT, vbNullString, vbNullString)
            oNew.nodeValue = vbCrLf & Left(sIndent, Len(sIndent) - 1)
            Set oNew = oElem.appendChild(oNew)
            Set oNew = Nothing
        End If
      End If 'oChildElems.length > 0
    End If 'Not oChildElems Is Nothing
    
    fncXmlIndentDom = True

End Function

Public Function fncHasChildNodeWithLocalName(oElem As IXMLDOMElement, sLocalName As String) As Boolean
Dim oChild As IXMLDOMNode
'do this namespace unaware so dont use xpath

  fncHasChildNodeWithLocalName = False
  For Each oChild In oElem.childNodes
    If oChild.nodeType = NODE_ELEMENT And oChild.nodeName = sLocalName Then
      fncHasChildNodeWithLocalName = True
      Exit Function
    End If
  Next

End Function

Public Function fncXmlIndentSax( _
  ByRef oDom As MSXML2.DOMDocument40 _
  ) As Boolean


Dim rdr As New SAXXMLReader40
Dim wrt As New MXXMLWriter40

  'On Error GoTo errhandler
  fncXmlIndentSax = False

  Set rdr.contentHandler = wrt
  Set rdr.errorHandler = wrt
  'If the application does not register a DTDHandler, all DTD events reported by the reader are ignored
  'Set rdr.dtdHandler = wrt
    
  rdr.putFeature "preserve-system-identifiers", True
  rdr.putFeature "schema-validation", False
  
  rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt
  rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
  
  'added 20050309:
  'rdr.putFeature "http://xml.org/sax/features/external-general-entities", True
  'rdr.putFeature "http://xml.org/sax/features/external-parameter-entities", True
  'rdr.putFeature "http://xml.org/sax/features/lexical-handler/parameter-entities", True

    
  wrt.output = ""
  wrt.byteOrderMark = False
  wrt.standalone = True
  wrt.indent = True
  wrt.omitXMLDeclaration = False
  'added 20050309:
  'wrt.disableOutputEscaping = True
  
  
  rdr.parse oDom
    
  'parse
  If Not fncParseString( _
    Replace(wrt.output, "standalone=" & Chr(34) & "yes" & Chr(34), ""), oDom, "") Then GoTo errhandler
  
  'get linefeeds out of textnode
  'mg20041007 removed bug reported by JK
'  Dim oTextNodes As IXMLDOMNodeList
'  Dim oTextNode As IXMLDOMNode
'
'  Set oTextNodes = oDom.selectNodes("//text()")
'  For Each oTextNode In oTextNodes
'    oTextNode.Text = Trim$(Replace(oTextNode.Text, vbLf, " "))
'    oTextNode.Text = Replace(oTextNode.Text, Chr(9), " ")
'    Do
'      oTextNode.Text = Replace(oTextNode.Text, "  ", " ")
'    Loop Until InStr(1, oTextNode.Text, "  ") = 0
'  Next
  
  fncXmlIndentSax = True
errhandler:
  
End Function

Public Function fncIsInNodeList(oTestNode As IXMLDOMNode, oNodeList As IXMLDOMNodeList) As Boolean
Dim oNode As IXMLDOMNode
  fncIsInNodeList = False
  For Each oNode In oNodeList
    If oNode Is oTestNode Then
      fncIsInNodeList = True
      Exit For
    End If
  Next
End Function

Public Function fncAllRelativesInNodeList(ByRef oTestNode As IXMLDOMNode, lRelativeType As Long) As Boolean
'are all relatives of input type in input nodelist?

'input relativetype is one of:
'Public Const RELATION_CHILD = 100
'Public Const RELATION_DESCENDANT = 101
'Public Const RELATION_SIBLING = 102
'Public Const RELATION_PARENT = 103
'Public Const RELATION_ANCESTOR = 104
'Public Const RELATION_SELF = 105
'Public Const RELATION_UNKNOWN = 106


End Function

Public Function fncIsRelativeInNodeList(ByRef oTestNode As IXMLDOMNode, ByRef oNodeList As IXMLDOMNodeList, lRelativeType As Long) As Boolean
Dim oNode As IXMLDOMNode
Dim oNodes As IXMLDOMNodeList
'does relative of oTestNode exist in input nodelist?

'input relativetype is one of:
'Public Const RELATION_CHILD = 100
'Public Const RELATION_DESCENDANT = 101
'Public Const RELATION_SIBLING = 102
'Public Const RELATION_PARENT = 103
'Public Const RELATION_ANCESTOR = 104
'Public Const RELATION_SELF = 105
'Public Const RELATION_UNKNOWN = 106
  
  fncIsRelativeInNodeList = False
  
  Select Case lRelativeType
    Case RELATION_CHILD
      Set oNodes = oTestNode.childNodes
      For Each oNode In oNodes
        If fncIsInNodeList(oNode, oNodeList) Then
          fncIsRelativeInNodeList = True
          Exit For
        End If
      Next
    Case RELATION_DESCENDANT
      Set oNodes = oTestNode.selectNodes(".//*")
      For Each oNode In oNodes
        If fncIsInNodeList(oNode, oNodeList) Then
          fncIsRelativeInNodeList = True
          Exit For
        End If
      Next
    Case RELATION_SIBLING
      Set oNodes = oTestNode.parentNode.childNodes
      For Each oNode In oNodes
        If fncIsInNodeList(oNode, oNodeList) Then
          fncIsRelativeInNodeList = True
          Exit For
        End If
      Next
    Case RELATION_PARENT
      Set oNode = oTestNode.parentNode
      If fncIsInNodeList(oNode, oNodeList) Then
          fncIsRelativeInNodeList = True
      End If
    Case RELATION_ANCESTOR
      Set oNode = oTestNode
      Do
        Set oNode = oNode.parentNode
        If fncIsInNodeList(oNode, oNodeList) Then fncIsRelativeInNodeList = True
      Loop Until (fncIsRelativeInNodeList) Or (oNode.parentNode.nodeType = NODE_DOCUMENT)
    Case RELATION_UNKNOWN
     '?
     
  End Select
  
End Function

Public Function fncRemoveAttribute( _
    ByRef oParentElem As IXMLDOMElement, _
    ByVal sAttrName As String _
    ) As Boolean
Dim oRemoveAttr As IXMLDOMNode
Dim oNamedNodeMap As IXMLDOMNamedNodeMap

  On Error GoTo errhandler
  fncRemoveAttribute = False
  
  Set oNamedNodeMap = oParentElem.Attributes
  If (Not oNamedNodeMap Is Nothing) And (oNamedNodeMap.length > 0) Then
      Set oRemoveAttr = oNamedNodeMap.removeNamedItem(sAttrName)
  End If
      
  fncRemoveAttribute = True
  
errhandler:
End Function

Public Function fncSetEncoding(ByRef oDom As MSXML2.DOMDocument40, ByVal sEncoding As String) As Boolean
Dim oProcessInstr As IXMLDOMProcessingInstruction
Dim oProcNodeMap As IXMLDOMNamedNodeMap
Dim oEncodingAttr As IXMLDOMAttribute
  On Error GoTo errh
  fncSetEncoding = False
  Set oEncodingAttr = oDom.createAttribute("encoding")
  oEncodingAttr.Text = sEncoding
  Set oProcessInstr = oDom.firstChild
  If Not oProcessInstr Is Nothing Then
    If oProcessInstr.nodeType = NODE_PROCESSING_INSTRUCTION Then
      Set oProcNodeMap = oDom.firstChild.Attributes
      Set oEncodingAttr = oProcNodeMap.setNamedItem(oEncodingAttr)
    End If
  End If
  fncSetEncoding = True
errh:
End Function

Public Function fncGetEncoding(oDom As MSXML2.DOMDocument40) As String
Dim oProcessInstr As IXMLDOMProcessingInstruction
Dim oProcNodeMap As IXMLDOMNamedNodeMap
Dim oEncodingAttr As IXMLDOMAttribute

  Set oProcessInstr = oDom.firstChild
  If Not oProcessInstr Is Nothing Then
    If oProcessInstr.nodeType = NODE_PROCESSING_INSTRUCTION Then
      Set oProcNodeMap = oDom.firstChild.Attributes
      Set oEncodingAttr = oProcNodeMap.getNamedItem("encoding")
      If Not oEncodingAttr Is Nothing Then
        fncGetEncoding = oEncodingAttr.Text
        Exit Function
      Else
        fncGetEncoding = "utf-8"
        Exit Function
      End If
    Else
      fncGetEncoding = "utf-8"
      Exit Function
    End If
    fncGetEncoding = "utf-8"
    Exit Function
  End If
End Function

Public Function fncRedundantNsRemove(ByRef oDom As MSXML2.DOMDocument40, ByVal bClearNullifiers As Boolean) As Boolean
Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim bNullNsFound As Boolean

  On Error GoTo errh
  fncRedundantNsRemove = False
  'takes a dom as input, iterates through it,
  'and removes duplicate in-scope xmlns attributes
  'if bClearNullifiers is true, it removes xmlns="" as well
  
  'select all descendants of root element
  Set oNodes = oDom.documentElement.selectNodes(".//*")
  
  If bClearNullifiers Then
    For Each oNode In oNodes
      If oNode.namespaceURI = "" Then
        bNullNsFound = True
        Me.fncAppendAttribute oNode, "xmlns", fncGetAncestorNs(oNode)
      End If
    Next oNode
    
    If bNullNsFound Then
      'reset the dom to get .namespaceURI prop right
      oDom.loadXML (oDom.xml)
    End If
  End If
  
  Set oNodes = oDom.documentElement.selectNodes(".//*")
  For Each oNode In oNodes
    If oNode.namespaceURI = oNode.parentNode.namespaceURI Then
      'inscope dupe, remove
      If Not Me.fncRemoveAttribute(oNode, "xmlns") Then GoTo errh
    End If
  Next oNode
    
  fncRedundantNsRemove = True
errh:

  Set oNodes = Nothing
  Set oNode = Nothing
End Function

Public Function fncGetAncestorNs(ByRef oNode As IXMLDOMNode) As String
Dim oParentNode As IXMLDOMNode
Dim oContextNode As IXMLDOMNode
'iterates upwards until an ancestor with a nonnull namespaceuri is found
  On Error GoTo errh
  Set oContextNode = oNode
  Do
    Set oContextNode = oContextNode.parentNode
  Loop Until (oContextNode.namespaceURI <> "") Or (oContextNode.parentNode.nodeType = NODE_DOCUMENT)
  fncGetAncestorNs = oContextNode.namespaceURI
errh:

End Function

Public Function fncHasIgnorableWhiteSpaceOnly(oElem As IXMLDOMElement) As Boolean
'returns true if children that are text are only iws
Dim oNode As IXMLDOMNode

  fncHasIgnorableWhiteSpaceOnly = True
  If oElem.hasChildNodes Then
    For Each oNode In oElem.childNodes
      If oNode.nodeType = NODE_TEXT Then
        If Len(Trim$(oNode.Text)) > 0 Then
          fncHasIgnorableWhiteSpaceOnly = False
          Exit Function
        End If
      End If
    Next
  End If

End Function

Public Function fncGetRelation( _
    oContextNode As IXMLDOMNode, _
    oRelativeNode As IXMLDOMNode) _
    As Long
Dim oNode As IXMLDOMNode
Dim oNodes As IXMLDOMNodeList

'returns a long describing the relation of oRelativeNode to oContextNode
'regards child as a subset of descendant
'and parent as a subset of ancestor
'subset returns are prioritized

''uses:
''Public Const RELATION_CHILD = 14
''Public Const RELATION_DESCENDANT = 15
''Public Const RELATION_SIBLING = 16
''Public Const RELATION_PARENT = 17
''Public Const RELATION_ANCESTOR = 18
''Public Const RELATION_UNKNOWN = 19


  'on error got errH
  
  'check if relative is parent
  If oContextNode.parentNode Is oRelativeNode Then
    fncGetRelation = RELATION_PARENT
    Exit Function
  End If
  
  'check if relative is child
  For Each oNode In oContextNode.childNodes
    If oNode Is oRelativeNode Then
      fncGetRelation = RELATION_CHILD
      Exit Function
    End If
  Next oNode
  
  'check if relative is sibling
  For Each oNode In oContextNode.parentNode.childNodes
    If oNode Is oRelativeNode Then
      fncGetRelation = RELATION_SIBLING
      Exit Function
    End If
  Next oNode

  'check if relative is descendant
  Set oNodes = oContextNode.selectNodes(".//*")
  For Each oNode In oNodes
    If oNode Is oRelativeNode Then
      fncGetRelation = RELATION_DESCENDANT
      Exit Function
    End If
  Next oNode
  
  'check if relative is ancestor
  Dim oPNode As IXMLDOMNode
  Set oPNode = oContextNode
  Do
    Set oPNode = oPNode.parentNode
    Set oNodes = oPNode.childNodes
    For Each oNode In oNodes
      If oNode Is oRelativeNode Then
        fncGetRelation = RELATION_ANCESTOR
        Exit Function
      End If
    Next
  Loop Until oPNode.parentNode.nodeType = NODE_DOCUMENT

errh:
  fncGetRelation = RELATION_UNKNOWN
End Function

Public Function fncRenameElement(oInElem As IXMLDOMElement, sNewName As String, sNsUri As String) As IXMLDOMElement
Dim oChild As IXMLDOMNode
Dim oInAttr As IXMLDOMAttribute
Dim oNewAttr As IXMLDOMAttribute
Dim oNewElem As IXMLDOMElement
Dim oOwner As New MSXML2.DOMDocument40
    oOwner.preserveWhiteSpace = True
    oOwner.setProperty "NewParser", True
    
 'set nsuri inparam to "" have a namespaceless rename
 'renames top level incoming element
  If sNsUri = "" Then
    Set oNewElem = oOwner.createElement(sNewName)
  Else
    Set oNewElem = oOwner.createNode(MSXML2.NODE_ELEMENT, sNewName, sNsUri)
  End If
  
  For Each oInAttr In oInElem.Attributes
    Set oNewAttr = oNewElem.Attributes.setNamedItem(oInAttr.cloneNode(True))
  Next
  
  For Each oChild In oInElem.childNodes
    oNewElem.appendChild oChild.cloneNode(True)
  Next
  Set fncRenameElement = oNewElem.cloneNode(True)
  Set oOwner = Nothing
End Function

Public Function fncHasMixedContentChildren(oElem As IXMLDOMElement) As Boolean
 'returns true if oElem has text and element children
Dim oChildNode As IXMLDOMNode
Dim bHasText As Boolean
Dim bHasElem As Boolean
 fncHasMixedContentChildren = False
 For Each oChildNode In oElem.childNodes
   If oChildNode.nodeType = NODE_ELEMENT Then
     bHasElem = True
   ElseIf oChildNode.nodeType = NODE_TEXT Or oChildNode.nodeType = NODE_ENTITY_REFERENCE Then
     bHasText = True
   End If
 Next
 If bHasText And bHasElem Then fncHasMixedContentChildren = True
End Function
       
Private Function bDTDIsExtendedInternallyForSkippableStructures( _
  ByRef sXml As String) As Boolean
  Dim lTest1 As Long, lTest2 As Long, lTest3 As Long, lTest4 As Long, lTest5 As Long
  
  On Error GoTo errh
  
  bDTDIsExtendedInternallyForSkippableStructures = False
  'check that all strings of internal declaration are there
  lTest1 = InStr(1, sXml, "ATTLIST", vbBinaryCompare): If lTest1 < 1 Then Exit Function
  lTest2 = InStr(1, sXml, "bodyref", vbBinaryCompare): If lTest2 < 1 Then Exit Function
  lTest3 = InStr(1, sXml, "CDATA", vbBinaryCompare): If lTest3 < 1 Then Exit Function
  lTest4 = InStr(1, sXml, "#IMPLIED", vbBinaryCompare): If lTest4 < 1 Then Exit Function
  'also check that all these string occur before root
  lTest5 = InStr(1, sXml, "<html", vbBinaryCompare)
  If (lTest1 > lTest5) Or (lTest2 > lTest5) Or (lTest3 > lTest5) Or (lTest4 > lTest5) Then Exit Function
  'appears to be an extended DTD
  bDTDIsExtendedInternallyForSkippableStructures = True
 
errh:
End Function
