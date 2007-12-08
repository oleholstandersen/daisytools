VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oDriverDom As New MSXML2.DOMDocument40

Public sName As String
Public sFullPath As String
Public sDefaultNameSpace As String
Public sDefaultNameSpaceUri As String
Public sLocalDtdPath As String

'Private sOutFileSet As String
Public lOutFileSet As Long

Public sXpathStartElem As String
Public sXpathRemove As String
Public sXpathSyncOmit As String
Public sXpathSyncForce As String
Public sXpathSmilRefOmit As String
Public sXpathSmilSequence As String
Public sXpathSmilSequenceUserEscape As String
Public sXpathSmilCustomTest As String
Public sXpathSmilSplit As String
Public sXpathNavInclude As String
Public sXpathNavIncludeExplicitNavLabel As String
Public sXpathMixedContentHandler As String

Public bMixedContentHandlerActive As Boolean
Public sItmtAttValue As String
Public sItmtAttName As String
Public sItmtElemName As String

Private Sub Class_Initialize()
  'debug.Print "cDriver.initialize"
  oDriverDom.async = False
  oDriverDom.validateOnParse = True
  oDriverDom.resolveExternals = True
  oDriverDom.preserveWhiteSpace = False
  oDriverDom.setProperty "SelectionLanguage", "XPath"
  oDriverDom.setProperty "SelectionNamespaces", "xmlns:bruno='http://www.daisy.org/2004/bruno'"
  oDriverDom.setProperty "NewParser", False
End Sub

Private Sub Class_Terminate()
  'debug.Print "cDriver.terminate"
  Set oDriverDom = Nothing
End Sub

Public Function fncInstantiate(sDriverFullPath As String) As Boolean
Dim oItemNode As IXMLDOMNode
Dim oItemNodes As IXMLDOMNodeList
    
  fncInstantiate = False
   
  If Not oBruno.oCmn.oDomCmn.fncParseFile(sDriverFullPath, oDriverDom, "") Then GoTo errh
  
  Me.sFullPath = sDriverFullPath
  Me.sName = oDriverDom.selectSingleNode("//bruno:name").Text
    
  Dim stempOutFileSet As String
  stempOutFileSet = oDriverDom.selectSingleNode("//bruno:bruno-driver/@outfileset").Text
  If stempOutFileSet = "d202" Then lOutFileSet = OUTPUT_TYPE_D202 Else lOutFileSet = OUTPUT_TYPE_Z39
  
  sDefaultNameSpace = oDriverDom.selectSingleNode("//bruno:bruno-driver/@target-namespace").Text
  sDefaultNameSpaceUri = oBruno.oCmn.oUriParser.fncGetAttrValueFromString(sDefaultNameSpace, Mid(sDefaultNameSpace, 1, InStr(1, sDefaultNameSpace, "=")))
  
  Me.sLocalDtdPath = oDriverDom.selectSingleNode("//bruno:bruno-driver/@local-dtd").Text
' mg 20070308: removed this since dtdPath in driver now can contain pipes
'  If Not oBruno.oCmn.oFsoCmn.fncFileExists(sLocalDtdPath) Then
'    'the user didnt provide a full path to the dtd in the driver,
'    'fallback to Bruno default DTD dir:
'    If Not fncSetLocalDtdPath(Me.sLocalDtdPath) Then
'      frmMain.fncAddMessage ("Local DTD not found")
'      Exit Function
'    End If
'  End If
  
  Set oItemNodes = oDriverDom.selectNodes("//bruno:start/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathStartElem) Then GoTo errh
  
  Set oItemNodes = oDriverDom.selectNodes("//bruno:remove/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathRemove) Then GoTo errh
    
  Set oItemNodes = oDriverDom.selectNodes("//bruno:sync-omit/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSyncOmit) Then GoTo errh
    
  Set oItemNodes = oDriverDom.selectNodes("//bruno:sync-force/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSyncForce) Then GoTo errh
    
  Set oItemNodes = oDriverDom.selectNodes("//bruno:smilref-omit/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSmilRefOmit) Then GoTo errh
          
  Set oItemNodes = oDriverDom.selectNodes("//bruno:smil-split/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSmilSplit) Then GoTo errh
    
  Set oItemNodes = oDriverDom.selectNodes("//bruno:nav-include/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathNavInclude) Then GoTo errh
  
  Set oItemNodes = oDriverDom.selectNodes("//bruno:nav-include/bruno:item[@navLabel]")
  If Not fncBuildXpath(oItemNodes, sXpathNavIncludeExplicitNavLabel) Then GoTo errh
      
  Set oItemNodes = oDriverDom.selectNodes("//bruno:mixed-content-handler/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathMixedContentHandler) Then GoTo errh
            
  Set oItemNode = oDriverDom.selectSingleNode("//bruno:mixed-content-handler[@active='true']")
  If Not oItemNode Is Nothing Then
    bMixedContentHandlerActive = True
  End If
  
  Set oItemNode = oDriverDom.selectSingleNode("//bruno:mixed-content-handler/@elemName")
  If Not oItemNode Is Nothing Then
    sItmtElemName = oItemNode.Text
  End If
        
  Set oItemNode = oDriverDom.selectSingleNode("//bruno:mixed-content-handler/@attName")
  If Not oItemNode Is Nothing Then
    sItmtAttName = oItemNode.Text
  End If
        
  Set oItemNode = oDriverDom.selectSingleNode("//bruno:mixed-content-handler/@attValue")
  If Not oItemNode Is Nothing Then
    sItmtAttValue = oItemNode.Text
  End If
        
        
  'these should not give any return for 2.02 output
  'since 2.02 does not allow constructs like these
  Set oItemNodes = oDriverDom.selectNodes("//bruno:smil-sequence/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSmilSequence) Then GoTo errh
    
  Set oItemNodes = oDriverDom.selectNodes("//bruno:smil-sequence/bruno:item[@user-escape='true']")
  If Not fncBuildXpath(oItemNodes, sXpathSmilSequenceUserEscape) Then GoTo errh
  
  Set oItemNodes = oDriverDom.selectNodes("//bruno:smil-customTest/bruno:item")
  If Not fncBuildXpath(oItemNodes, sXpathSmilCustomTest) Then GoTo errh
                
  fncInstantiate = True
errh:

 Set oItemNode = Nothing
 Set oItemNodes = Nothing
End Function

Private Function fncBuildXpath(ByRef oItemNodes As IXMLDOMNodeList, ByRef sXpathString As String) As Boolean
Dim i As Long
  fncBuildXpath = False
  
  sXpathString = ""
  If Not oItemNodes Is Nothing Then
    For i = 0 To oItemNodes.length - 2
      sXpathString = sXpathString & oItemNodes(i).selectSingleNode("@xpath").Text & " | "
    Next
    'add last iterat without ending pipe
    If Not oItemNodes.length = 0 Then sXpathString = sXpathString & oItemNodes(i).selectSingleNode("@xpath").Text
  End If
  If sXpathString = "" Then sXpathString = "/nothing"
  
  fncBuildXpath = True
End Function

Public Function fncSetLocalDtdPath(sLocalDtdName As String) As Boolean
  fncSetLocalDtdPath = False
  Me.sLocalDtdPath = oBruno.oPaths.DtdPath & sLocalDtdName
  If oBruno.oCmn.oFsoCmn.fncFileExists(Me.sLocalDtdPath) Then fncSetLocalDtdPath = True
End Function

Public Function fncGetXpathForNavLabel(oInElement As IXMLDOMElement) As String
Dim oDriverNavIncludeElement As IXMLDOMElement
Dim oItems As IXMLDOMNodeList
Dim oDocElems As IXMLDOMNodeList
Dim oDocElem As IXMLDOMNode
Dim oItem As IXMLDOMElement
Dim sXpathAttrValue As String
Dim sNavLabelAttrValue As String
 'gets a node as input which is a member of
 'oAbstractDocuments.oNavIncludeNodesExplicitNavLabel
 'take the inElements elemname
 'and find the nav-include/item whose xpath statement matches
 'return the navLabel attribute string of that item

 Set oDriverNavIncludeElement = oDriverDom.selectSingleNode("//bruno:nav-include")
 If Not oDriverNavIncludeElement Is Nothing Then
   Set oItems = oDriverNavIncludeElement.selectNodes("./bruno:item")
   For Each oItem In oItems
     sXpathAttrValue = oItem.selectSingleNode("@xpath").Text
     'execute the xpath on oInElements parentdoc, and see if oInElem is part of return set
     'if so, we have the item we want
     Set oDocElems = oInElement.ownerDocument.selectNodes(sXpathAttrValue)
     For Each oDocElem In oDocElems
       If oInElement Is oDocElem Then
           'this is the item we are looking for
            fncGetXpathForNavLabel = oItem.selectSingleNode("@navLabel").Text
            Exit Function
       End If
     Next oDocElem
   Next oItem
 Else 'oDriverNavInclude is nothing
   fncGetXpathForNavLabel = ""
 End If 'oDriverNavInclude is nothing
End Function

'Public Function fncGetXpathForNavLabelOld(oInElement As IXMLDOMElement) As String
'Dim oDriverNavIncludeElement As IXMLDOMElement
'Dim oItems As IXMLDOMNodeList
'Dim oItem As IXMLDOMElement
'Dim sXpathAttrValue As String
'Dim sNavLabelAttrValue As String
' 'gets a node as input which is a member of
' 'oAbstractDocuments.oNavIncludeNodesExplicitNavLabel
' 'take the inElements elemname
' 'and find the nav-include/item whose xpath statement matches
' 'return the navLabel attribute string of that item
'
' Set oDriverNavIncludeElement = oDriverDom.selectSingleNode("//bruno:nav-include")
' If Not oDriverNavIncludeElement Is Nothing Then
'   Set oItems = oDriverNavIncludeElement.selectNodes("./bruno:item")
'   For Each oItem In oItems
'     sXpathAttrValue = oItem.selectSingleNode("@xpath").Text
'     sXpathAttrValue = Replace(sXpathAttrValue, "//dtbook:", "")
'     '20070129: add support for xhtml as well
'     'sXpathAttrValue = Replace(sXpathAttrValue, "//xhtml:", "")
'     If sXpathAttrValue = oInElement.nodeName Then
'       'this is the item we are looking for
'       fncGetXpathForNavLabel = oItem.selectSingleNode("@navLabel").Text
'       Exit Function
'     End If
'   Next oItem
' Else 'oDriverNavInclude is nothing
'   fncGetXpathForNavLabel = ""
' End If 'oDriverNavInclude is nothing
'
'End Function
