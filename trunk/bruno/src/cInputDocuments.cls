VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInputDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Private aInputDocuments() As cInputDocument
Private lInputDocumentCount As Long

Public oInputMetadata As cInputMetadata
Public oInputAuxFiles As cInputAuxFiles

Private Sub Class_Initialize()
  'debug.print"cInputDocuments.initialize"
  lInputDocumentCount = 0
End Sub

Public Function InputDocumentCount() As Long
  InputDocumentCount = lInputDocumentCount
End Function

Public Function InputDocument(lDocument As Long) As cInputDocument
  Set InputDocument = aInputDocuments(lDocument)
End Function

Public Function fncAddDocument(sInputDocFullPath As String) As Boolean
Dim oInputDom As New MSXML2.DOMDocument40
    oInputDom.async = False
    oInputDom.validateOnParse = False
    oInputDom.resolveExternals = False
    oInputDom.preserveWhiteSpace = False
    oInputDom.setProperty "SelectionLanguage", "XPath"
    oInputDom.setProperty "SelectionNamespaces", oBruno.oDriver.sDefaultNameSpace
    oInputDom.setProperty "NewParser", True
  On Error GoTo errh
Dim sXmlParseMessage As String
Dim sNsLocus As String
       
  If oBruno.oCmn.oFsoCmn.fncFileExists(sInputDocFullPath) Then
  
    ReDim Preserve aInputDocuments(lInputDocumentCount)
    Set aInputDocuments(lInputDocumentCount) = New cInputDocument
    With aInputDocuments(lInputDocumentCount)
      .sFullPath = sInputDocFullPath
      .sFileName = oBruno.oCmn.oFsoCmn.fncGetFileName(sInputDocFullPath)
      .sDirectory = oBruno.oCmn.oUriParser.fncGetPathName(sInputDocFullPath)
      'for safety:
      .sDirectory = oBruno.oCmn.oFsoCmn.fncGetFSOPathFromString(.sDirectory)
      .bWellformed = oBruno.oCmn.oDomCmn.fncParseFile(sInputDocFullPath, oInputDom, sXmlParseMessage)
      If .bWellformed = False Then
        .bValid = False
        frmMain.fncAddMessage sXmlParseMessage
      Else
        fncSetInputDocProperties oInputDom, aInputDocuments(lInputDocumentCount)
        If .bValid Then Set .oDom = oInputDom.cloneNode(True)
      End If '.bWellformed = False Then
      
    End With
    lInputDocumentCount = lInputDocumentCount + 1
                         
    fncAddDocument = True
    Set oInputDom = Nothing
  End If 'fncFileExists
errh:
End Function

Public Function fncReValidateInputDocuments() As Boolean
Dim oInputDom As New MSXML2.DOMDocument40
    oInputDom.async = False
    oInputDom.validateOnParse = False
    oInputDom.resolveExternals = False
    oInputDom.preserveWhiteSpace = False
    oInputDom.setProperty "SelectionLanguage", "XPath"
    oInputDom.setProperty "SelectionNamespaces", oBruno.oDriver.sDefaultNameSpace
    oInputDom.setProperty "NewParser", True
Dim i As Long, sXmlParseMessage As String

  If Me.InputDocumentCount < 1 Then Exit Function
  For i = 0 To Me.InputDocumentCount - 1
   With aInputDocuments(i)
     .bWellformed = oBruno.oCmn.oDomCmn.fncParseFile(.sFullPath, oInputDom, sXmlParseMessage)
      If .bWellformed = False Then
        .bValid = False
        frmMain.fncAddMessage sXmlParseMessage
      Else
        fncSetInputDocProperties oInputDom, aInputDocuments(i)
        If .bValid Then Set .oDom = oInputDom.cloneNode(True)
      End If '.bWellformed = False Then
   End With
  Next i
End Function
Private Function fncSetInputDocProperties(ByRef oInputDom As MSXML2.DOMDocument40, ByRef oInputDocument As cInputDocument) As Boolean
Dim sXmlParseMessage As String
 'return true if still valid after properties set
 With oInputDocument
  .sEncoding = oBruno.oCmn.oDomCmn.fncGetEncoding(oInputDom)
  .bValid = oBruno.oCmn.oDomCmn.fncCheckValidity(oInputDom, oBruno.oDriver, sXmlParseMessage)
  If sXmlParseMessage <> "" Then frmMain.fncAddMessage sXmlParseMessage
  If .bValid Then
    .bValid = fncGetNsLocus(oInputDom) 'check that a namespace decl is present
    .sDocTypeNiceName = fncGetDocTypeNiceName(oInputDom)
    If .bValid Then
      'check for aux files
      If oInputAuxFiles Is Nothing Then
        Set oInputAuxFiles = New cInputAuxFiles
        oInputAuxFiles.fncInstantiate
      End If
      'oInputAuxFiles.fncCheckInputDocumentForAuxFiles (lInputDocumentCount - 1), oInputDom
      oInputAuxFiles.fncCheckInputDocumentForAuxFiles lInputDocumentCount, oInputDocument.sFullPath, oInputDom
    End If
   End If
   
  fncSetInputDocProperties = .bValid
  
  End With
End Function
Private Function fncGetNsLocus(ByRef oInputDom As MSXML2.DOMDocument40) As Boolean
Dim sNsLocus As String

  If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
    sNsLocus = "//xhtml:html"
  Else
    sNsLocus = "//dtbook:dtbook"
  End If
        
  If oInputDom.selectSingleNode(sNsLocus) Is Nothing Then
    frmMain.fncAddMessage "no root namespace declaration in input document"
    fncGetNsLocus = False
  Else
    fncGetNsLocus = True
  End If
End Function

Private Function fncGetDocTypeNiceName(ByRef oInputDom As MSXML2.DOMDocument40) As String
Dim sTemp As String
  sTemp = oInputDom.documentElement.nodeName
  If sTemp = "html" Then sTemp = "xhtml"
  fncGetDocTypeNiceName = sTemp
End Function

Public Function fncResetArrays() As Boolean
Dim i As Long
  'clears the array, but not the inputmetadata object
  For i = 0 To lInputDocumentCount - 1
    Set aInputDocuments(i).oDom = Nothing
  Next i
  
  lInputDocumentCount = 0
  ReDim Preserve aInputDocuments(lInputDocumentCount)
End Function

Public Function fncResetDomObjectsInArrays() As Boolean
Dim i As Long
  'clears the doms in the array, nothing else (see fncResetArrays)
  For i = 0 To lInputDocumentCount - 1
    Set aInputDocuments(i).oDom = Nothing
  Next i
  
End Function

Public Function fncRemoveDocument(lArrayIndex As Long) As Boolean
 'removes a document from the array
Dim i As Long
 
  If lInputDocumentCount = 0 Then Exit Function

  For i = lArrayIndex To lInputDocumentCount - 2
   aInputDocuments(i).bValid = aInputDocuments(i + 1).bValid
    aInputDocuments(i).bWellformed = aInputDocuments(i + 1).bWellformed
    aInputDocuments(i).oDom.loadXML (aInputDocuments(i + 1).oDom.xml)
    aInputDocuments(i).sEncoding = aInputDocuments(i + 1).sEncoding
    aInputDocuments(i).sFileName = aInputDocuments(i + 1).sFileName
    aInputDocuments(i).sFullPath = aInputDocuments(i + 1).sFullPath
    aInputDocuments(i).sDocTypeNiceName = aInputDocuments(i + 1).sDocTypeNiceName
  Next

  lInputDocumentCount = lInputDocumentCount - 1
  If lInputDocumentCount = 0 Then
    ReDim Preserve aInputDocuments(lInputDocumentCount)
  Else
    ReDim Preserve aInputDocuments(lInputDocumentCount - 1)
  End If

End Function

Public Function fncMoveDocument(lArrayIndex As Long, lDirection As Long) As Long
 'moves a document in the array
Dim dest As Long
Dim tempStrEncoding As String
Dim tempFileName As String
Dim tempFullPath As String
Dim tempBolValid As Boolean
Dim tempBolWellformed As Boolean
Dim tempStrDom As String
Dim tempStrDocTypeNiceName As String
  If lDirection = DIRECTION_UP Then
    dest = lArrayIndex - 1
  Else
    dest = lArrayIndex + 1
  End If

  If dest > lInputDocumentCount - 1 Then
    fncMoveDocument = dest - 1
    Exit Function
  End If

  If dest < 0 Then
    fncMoveDocument = dest + 1
    Exit Function
  End If

  tempBolValid = aInputDocuments(dest).bValid
  tempBolWellformed = aInputDocuments(dest).bWellformed
  tempStrEncoding = aInputDocuments(dest).sEncoding
  tempStrDom = aInputDocuments(dest).oDom.xml
  tempFileName = aInputDocuments(dest).sFileName
  tempFullPath = aInputDocuments(dest).sFullPath
  tempStrDocTypeNiceName = aInputDocuments(dest).sDocTypeNiceName

  aInputDocuments(dest).bValid = aInputDocuments(lArrayIndex).bValid
  aInputDocuments(dest).bWellformed = aInputDocuments(lArrayIndex).bWellformed
  aInputDocuments(dest).sEncoding = aInputDocuments(lArrayIndex).sEncoding
  aInputDocuments(dest).oDom.loadXML (aInputDocuments(lArrayIndex).oDom.xml)
  aInputDocuments(dest).sFileName = aInputDocuments(lArrayIndex).sFileName
  aInputDocuments(dest).sFullPath = aInputDocuments(lArrayIndex).sFullPath
  aInputDocuments(dest).sDocTypeNiceName = aInputDocuments(lArrayIndex).sDocTypeNiceName

  aInputDocuments(lArrayIndex).bValid = tempBolValid
  aInputDocuments(lArrayIndex).bWellformed = tempBolWellformed
  aInputDocuments(lArrayIndex).sEncoding = tempStrEncoding
  aInputDocuments(lArrayIndex).oDom.loadXML (tempStrDom)
  aInputDocuments(lArrayIndex).sFileName = tempFileName
  aInputDocuments(lArrayIndex).sFullPath = tempFullPath
  aInputDocuments(lArrayIndex).sDocTypeNiceName = tempStrDocTypeNiceName

  fncMoveDocument = dest

End Function

Public Function fncCheckDocuments() As Boolean
Dim i As Long
Dim bAllValid As Boolean
  'returns true if all docs in array are fully valid
  
  If Me.InputDocumentCount < 1 Then
    fncCheckDocuments = False
    Exit Function
  End If
  bAllValid = True

  For i = 0 To Me.InputDocumentCount - 1
    If Me.InputDocument(i).bValid = False Then
      bAllValid = False
      Exit For
    End If
  Next
 
 fncCheckDocuments = bAllValid
 
End Function

Private Sub Class_Terminate()
  'debug.print "cInputDocuments.terminate"
  Set Me.oInputMetadata = Nothing
  Set Me.oInputAuxFiles = Nothing
End Sub

'Private Function fncGetDoctype(oInputDom As MSXML2.DOMDocument40) As Long
'
'Dim k As Long
'Dim oDcNode As IXMLDOMNode
'Dim oNode As IXMLDOMNode
'Dim oRootElemNode As IXMLDOMNode
'
'  'find the doctype node
'  For k = 0 To oInputDom.childNodes.length - 1
'    If oInputDom.childNodes.Item(k).nodeType = NODE_DOCUMENT_TYPE Then
'      Set oDcNode = oInputDom.childNodes.Item(k)
'      Exit For
'    End If
'  Next
'
'  'if doctypedecl node was not found
'   If oDcNode Is Nothing Then
'      fncGetDoctype = DOCTYPE_UNKNOWN
'      Exit Function
'   Else 'node was found, so check that root element name equals expectation
'     Set oRootElemNode = oInputDom.documentElement
'     Dim sPID As String
'     For Each oNode In oDcNode.Attributes
'       If oNode.nodeName = "PUBLIC" Then
'         sPID = oDcNode.Attributes.Item(0).nodeValue
'         Exit For
'       End If
'     Next oNode
'
'     If sPID = "-//W3C//DTD XHTML 1.0 Strict//EN" And oRootElemNode.nodeName = "html" Then
'       fncGetDoctype = DOCTYPE_XHTML1_STRICT
'     ElseIf sPID = "-//W3C//DTD XHTML 1.0 Transitional//EN" And oRootElemNode.nodeName = "html" Then
'       fncGetDoctype = DOCTYPE_XHTML1_TRANSITIONAL
'     ElseIf sPID = "-//W3C//DTD XHTML 1.1//EN" And oRootElemNode.nodeName = "html" Then
'       fncGetDoctype = DOCTYPE_XHTML11_TRANSITIONAL
'     ElseIf sPID = "-//NISO//DTD dtbook v1.1.0//EN" Or oRootElemNode.nodeName = "dtbook" Then
'       fncGetDoctype = DOCTYPE_DTBOOK
'     Else
'       fncGetDoctype = DOCTYPE_UNKNOWN
'     End If
'   End If
'
'
'     Set oDcNode = Nothing
'     Set oRootElemNode = Nothing
'
'End Function

