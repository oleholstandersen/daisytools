VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oValidateNcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

' *** Validate Content ***
' This object is for validation of ncc files
' objReport = report object that collects all files reported by this object

Public objReport As oReport

' The main function. Arguments work as follows:
' isAbsPath = Full path + filename of file to validate
' ibolNccIsMultivolume = should be set to true if this is a multi volume ncc
' Return value is "false" if an internal error occurs, otherwise it is "true"
'
Public Function fncValidate( _
  isAbsPath As String, Optional ibolNccIsMultivolume As Variant _
  ) As Boolean
  
  fncValidate = False
  'fncInsertTime "oValidateNcc.fncValidate"
  
  Dim bolResult As Boolean, enFileType As enuFileType, sCustomDtd As String
  
  Dim objFileIntegrity As New oFileIntegrity, objXmlIntegrity As New oXmlIntegrity
  Dim objXMLIntegrityMSXML As New oXmlIntegrityMSXML
  Dim objAttributeDTD As New oAttributeDTD, objDom As Object
  
  Set objReport = New oReport
  
  enFileType = ncc
  sCustomDtd = "d202ncc.dtd"
  
  If Not IsMissing(ibolNccIsMultivolume) Then
    If ibolNccIsMultivolume Then
      enFileType = nccmultivolume
      sCustomDtd = "d202nccmulti.dtd"
    End If
  End If
  
  fncSetProgress Me, 0, 10
  
  Dim objNode As Object, sContext As String
  
  If Not objFileIntegrity.fncFileExists( _
    objReport, isAbsPath _
    ) Then
    objEvent.subLog ("failure in oValidateNcc.fncValidate.objFileIntegrity.fncFileExists")
    objEvent.subLog ("validation process prematurely aborted")
    fncValidate = True 'if the ncc does not exist, no sense in continuing
    GoTo ErrorH
  End If
  fncSetProgress Me, 1, 10
    
  If Not objFileIntegrity.fncFileHasValidName( _
    objReport, isAbsPath, ncc _
    ) Then
    objEvent.subLog ("failure in ovalidateNcc.fncValidate.objFileIntegrity.fncFileHasValidName")
    GoTo ErrorH
  End If
  fncSetProgress Me, 2, 10
    
  If Not objFileIntegrity.fncFileIsReadable( _
    objReport, isAbsPath _
    ) Then
    objEvent.subLog ("failure in ovalidateNcc.fncValidate.objFileIntegrity.fncFileIsReadable")
    objEvent.subLog ("validation process prematurely aborted")
    fncValidate = True 'if the ncc cannot be read, no sense in continuing
    GoTo ErrorH
  End If
  fncSetProgress Me, 3, 10
  
  If Not objXmlIntegrity.fncIsWellformedXML( _
    objReport, isAbsPath, objDom _
    ) Then
    objEvent.subLog ("failure in ovalidateNcc.fncValidate.objXMLIntegrity.fncIsWellformedXML")
    GoTo ErrorH
  End If
  fncSetProgress Me, 4, 10
 
  If Not objDom Is Nothing Then
    If Not objXmlIntegrity.fncHasAndIsDocumenttype( _
      objReport, objDom, xhtml10, isAbsPath _
      ) Then
      objEvent.subLog ("failure in ovalidateNcc.fncValidate.objXMLIntegrity.fncHasAndIsDocumenttype")
      GoTo ErrorH
    End If
    fncSetProgress Me, 5, 10
    
    'set the dc:identifier property in oUserControl
    'fncSetDcIdentifierProp (objDom)
  
    If Not objXMLIntegrityMSXML.documentValidatesToGivenDTD( _
      objReport, isAbsPath, "xhtml1-transitional.dtd", ncc _
      ) Then
      objEvent.subLog ("failure in ovalidateNcc.fncValidate.objXMLIntegrityElcel.documentValidatesToGivenDTD")
      GoTo ErrorH
    End If
    fncSetProgress Me, 6, 10
  
    If Not objXMLIntegrityMSXML.documentValidatesToCustomDTD( _
      objReport, isAbsPath, sCustomDtd, ncc _
      ) Then
      objEvent.subLog ("failure in ovalidateNcc.fncValidate.objXMLIntegrityElcel.documentValidatesToCustomDTD")
      GoTo ErrorH
    End If
    fncSetProgress Me, 7, 10
    
If Not bolLightMode Then
    If Not objAttributeDTD.fncDocumentAttributesContentValid( _
      objReport, objDom, enFileType, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.objAttributeDTD.fncDocumentAttributesContentValid")
      GoTo ErrorH
    End If
End If 'Not bolLightMode
    fncSetProgress Me, 8, 10
    If bolCancelValidation Then fncValidate = True: GoTo ErrorH

If Not bolLightMode Then
    If Not objAttributeDTD.fncIntraElementAttributesCorelationValid( _
      objReport, objDom, enFileType, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.objAttributeDTD.fncIntraElementAttributesCorelationValid")
      GoTo ErrorH
    End If
End If 'Not bolLightMode
    fncSetProgress Me, 9, 10
    If bolCancelValidation Then fncValidate = True: GoTo ErrorH
  
    If Not fncHeadingsAreCorrectlyNested( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncHeadingsAreCorrectlyNested")
      GoTo ErrorH
    End If

If Not bolLightMode Then
    If Not fncNumPageNormalsCorrespondsToFile( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncNumPageNormalsCorrespondsToFile")
      GoTo ErrorH
    End If
End If 'Not bolLightMode

If Not bolLightMode Then
    If Not fncNumPageSpecialsCorrespondsToFile( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncNumPageSpecialsCorrespondsToFile")
      GoTo ErrorH
    End If
End If 'Not bolLightMode

If Not bolLightMode Then
    If Not fncNumPageFrontsCorrespondsToFile( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncNumPageFrontsCorrespondsToFile")
      GoTo ErrorH
    End If
End If 'Not bolLightMode

If Not bolLightMode Then
    If Not fncNumFootnotesCorrespondsToFile( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncNumFootnotesCorrespondsToFile")
      GoTo ErrorH
    End If
End If 'Not bolLightMode

If Not bolLightMode Then
    If Not fncNumProdNotesCorrespondsToFile( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncNumProdNotesCorrespondsToFile")
      GoTo ErrorH
    End If
End If 'Not bolLightMode

If Not bolLightMode Then
    If Not fncNumSidebarsCorrespondsToFile( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncNumSidebarsCorrespondsToFile")
      GoTo ErrorH
    End If
End If 'Not bolLightMode

If Not bolLightMode Then
    If Not fncHeadingDepthIsRight( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncHeadingDepthIsRight")
      GoTo ErrorH
    End If
End If 'Not bolLightMode

If Not bolLightMode Then
    If Not fncNccTocItemsCorrespondsToFile( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncNccTocItemsCorrespondsToFile")
      GoTo ErrorH
    End If
  End If
End If 'Not bolLightMode


    If Not fncfirstBodyDescIsH1ClassTitle( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateNcc.fncValidate.fncfirstBodyDescIsH1ClassTitle")
      GoTo ErrorH
    End If
  


  fncValidate = True
ErrorH:
  objReport.fncAddContext "ncc"
  fncSetProgress Me, 10, 10

  Set objDom = Nothing
  Set objLastLCNode = Nothing
  Set objFileIntegrity = Nothing
  Set objXmlIntegrity = Nothing
  Set objXMLIntegrityMSXML = Nothing
  Set objAttributeDTD = Nothing
  
  'fncInsertTime "oValidateNcc.fncValidate"
End Function

Private Function fncHeadingsAreCorrectlyNested(iobjReport As oReport, iobjDOMStructure _
  As Object, isAbsPath As String)
  
  fncHeadingsAreCorrectlyNested = False
  On Error GoTo ErrH
  
  Dim objNode As Object, objNodeList As Object
  
  Set objNode = iobjDOMStructure.documentElement.selectSingleNode("body")
  If objNode Is Nothing Then GoTo ErrH
  Set objNodeList = objNode.selectNodes("h1 | h2 | h3 | h4 | h5 | h6")
    
  Dim lLevel As Long, lValue As Long
  lLevel = 1
  For Each objNode In objNodeList
    If objNode Is Nothing Then
        objEvent.subLog ("failure in oValidateNcc.fncHeadingsAreCorrectlyNested")
        GoTo ErrH
    End If
    lValue = Right$(objNode.nodeName, 1)
    If Not (lValue <= lLevel + 1) Then
      fncInsFail2Report iobjReport, objNode, _
        "nccSpecific.headingsAreCorrectlyNested", isAbsPath
    Else
      iobjReport.subInsertSucceededTest
    End If
    lLevel = lValue
  Next
  
  fncHeadingsAreCorrectlyNested = True
ErrH:
  Set objNode = Nothing
  Set objNodeList = Nothing
  
  If Not fncHeadingsAreCorrectlyNested Then
    iobjReport.fncInsertFailedTest "nccSpecific.fncHeadingsAreCorrectlyNested", _
      isAbsPath, -1, -1
  End If
End Function

Private Function fncfirstBodyDescIsH1ClassTitle( _
    iobjReport As oReport, _
    iobjDOMStructure As MSXML2.DOMDocument40, _
    isAbsPath As String _
    ) As Boolean
Dim oTestNode As IXMLDOMNode
Dim objH1Node As IXMLDOMNode
Dim bTestFailed As Boolean
    
    fncfirstBodyDescIsH1ClassTitle = False
    bTestFailed = False
    
    Set objH1Node = iobjDOMStructure.documentElement.selectSingleNode("body/*[position()=1]")
    If Not objH1Node Is Nothing Then
      If objH1Node.nodeName = "h1" Then
        'we know per ncccustomdtd if first body child is h1
        'test if first h1 has class="title"
        Set oTestNode = objH1Node.selectSingleNode("@class")
        'Stop
        If oTestNode Is Nothing Then
          bTestFailed = True
        Else
          If oTestNode.text <> "title" Then
            bTestFailed = True
          End If
        End If
      Else
        bTestFailed = True
      End If 'Not objH1Node.nodeName = "h1" Then
    Else
      bTestFailed = True
    End If 'Not objH1Node Is Nothing Then
    
    If bTestFailed Then
     fncInsFail2Report iobjReport, objH1Node, "nccSpecific.firstBodyDescIsH1ClassTitle", _
     isAbsPath, "suggested value: class='title'"
    Else
      iobjReport.subInsertSucceededTest
    End If
        
    fncfirstBodyDescIsH1ClassTitle = True

End Function

Private Function fncNumPageNormalsCorrespondsToFile(iobjReport As oReport, _
  iobjDOMStructure As Object, isAbsPath As String)
  fncNumPageNormalsCorrespondsToFile = fncMetaCorrespondsToFile( _
    iobjReport, "nccSpecific.numPageNormalsCorrespondsToFile", _
    "//meta[@name = 'ncc:pageNormal']/@content", _
    "//span[@class = 'page-normal']", iobjDOMStructure, isAbsPath _
    )
End Function

Private Function fncNumPageSpecialsCorrespondsToFile(iobjReport As oReport, _
  iobjDOMStructure As Object, isAbsPath As String)
  
  fncNumPageSpecialsCorrespondsToFile = fncMetaCorrespondsToFile( _
    iobjReport, "nccSpecific.numPageSpecialsCorrespondsToFile", _
    "//meta[@name = 'ncc:pageSpecial']/@content", _
    "//span[@class = 'page-special']", iobjDOMStructure, isAbsPath _
    )
End Function

Private Function fncNumPageFrontsCorrespondsToFile(iobjReport As oReport, _
  iobjDOMStructure As Object, isAbsPath As String)
  
  fncNumPageFrontsCorrespondsToFile = fncMetaCorrespondsToFile( _
    iobjReport, "nccSpecific.numPageFrontsCorrespondsToFile", _
    "//meta[@name = 'ncc:pageFront']/@content", _
    "//span[@class = 'page-front']", iobjDOMStructure, isAbsPath _
    )
End Function

Private Function fncNumFootnotesCorrespondsToFile(iobjReport As oReport, _
  iobjDOMStructure As Object, isAbsPath As String)
  
  fncNumFootnotesCorrespondsToFile = fncMetaCorrespondsToFile( _
    iobjReport, "nccSpecific.numFootnotesCorrespondsToFile", _
    "//meta[@name = 'ncc:footnotes']/@content", _
    "//span[@class = 'noteref']", iobjDOMStructure, isAbsPath _
    )
End Function

Private Function fncNumProdNotesCorrespondsToFile(iobjReport As oReport, _
  iobjDOMStructure As Object, isAbsPath As String)
  
  fncNumProdNotesCorrespondsToFile = fncMetaCorrespondsToFile( _
    iobjReport, "nccSpecific.numProdNotesCorrespondsToFile", _
    "//meta[@name = 'ncc:prodNotes']/@content", _
    "//span[@class = 'optional-prodnote']", iobjDOMStructure, isAbsPath _
    )
End Function

Private Function fncNumSidebarsCorrespondsToFile(iobjReport As oReport, _
  iobjDOMStructure As Object, isAbsPath As String)
  
  fncNumSidebarsCorrespondsToFile = fncMetaCorrespondsToFile( _
    iobjReport, "nccSpecific.numSidebarsCorrespondsToFile", _
    "//meta[@name = 'ncc:sidebars']/@content", _
    "//span[@class = 'sidebar']", iobjDOMStructure, isAbsPath _
    )
End Function

Private Function fncHeadingDepthIsRight(iobjReport As oReport, _
  iobjDOMStructure As Object, isAbsPath As String)
  
  Dim objNode As Object, objNodeList As Object
  Dim lHeadingDepth As Long, lRetrievedHD As Long, lDeepestFound As Long
  
  fncHeadingDepthIsRight = False
  
  Set objNode = iobjDOMStructure.selectSingleNode( _
    "//meta[@name = 'ncc:depth']/@content")
  If Not objNode Is Nothing Then 'if the elm doesnt exist this test should not be run: the dtd reports its abscence before this
    lRetrievedHD = fncString2Integer(objNode.nodeValue)

'    mg20030314:
'    For lHeadingDepth = 1 To 6
'      Set objNodeList = iobjDOMStructure.selectNodes("//h" & lHeadingDepth)
'      If objNodeList.length = 0 Then lHeadingDepth = lHeadingDepth - 1: Exit For
'    Next lHeadingDepth
'    mg20030313
'    If lHeadingDepth = 7 Then lHeadingDepth = 6

'    mg20030314: do like this instead since sometimes incorrect sequence of heading levels
    For lHeadingDepth = 1 To 6
      Set objNodeList = iobjDOMStructure.selectNodes("//h" & lHeadingDepth)
      If Not objNodeList Is Nothing Then
        If objNodeList.length > 0 Then lDeepestFound = lHeadingDepth
      End If
    Next lHeadingDepth
        
    If Not lDeepestFound = lRetrievedHD Then
'    If Not lHeadingDepth = lRetrievedHD Then
      Set objNode = objNode.selectSingleNode("..")
      fncInsFail2Report iobjReport, objNode, "nccSpecific.headingDepth", _
        isAbsPath, "suggested value: " & lDeepestFound
    Else
      iobjReport.subInsertSucceededTest
    End If
  End If
  fncHeadingDepthIsRight = True
  
  Set objNode = Nothing
  Set objNodeList = Nothing
End Function

Private Function fncMetaCorrespondsToFile(iobjReport As oReport, isTestId As String, _
  isTagHasValue As String, isCorrespondsToNumberOf As String, _
  iobjDOMStructure As Object, isAbsPath As String) As Boolean
  
  fncMetaCorrespondsToFile = False
  
  Dim objNode As Object, objNodeList As Object
    
  Set objNode = iobjDOMStructure.selectSingleNode(isTagHasValue)
  
  If Not objNode Is Nothing Then 'if the elm doesnt exist this test should not be run: the dtd reports its abscence before this
    Set objNodeList = iobjDOMStructure.selectNodes(isCorrespondsToNumberOf)
  
    If Not fncString2Integer(objNode.nodeValue) = objNodeList.length Then
      Set objNode = objNode.selectSingleNode("..")
      fncInsFail2Report iobjReport, objNode, isTestId, isAbsPath, _
        "suggested value: " & objNodeList.length
    Else
        iobjReport.subInsertSucceededTest
    End If
  End If
  
  Set objNode = Nothing
  Set objNodeList = Nothing
  
  fncMetaCorrespondsToFile = True
End Function

Private Function fncNccTocItemsCorrespondsToFile( _
    iobjReport As oReport, iobjDom As Object, isAbsPath As String _
    ) As Boolean
    
  Dim objNode As Object, lMetaTocItems As Long, lCompTocItems As Long
  Dim objNodeList As Object
  
  fncNccTocItemsCorrespondsToFile = True
  
  Set objNode = iobjDom.selectSingleNode("//head/meta[@name='ncc:tocItems']" & _
    "/@content")
  If objNode Is Nothing Then Exit Function
  
  lMetaTocItems = fncString2Integer(objNode.nodeValue)
  
  Set objNodeList = iobjDom.selectNodes("//a")
  If lMetaTocItems = objNodeList.length Then
    iobjReport.subInsertSucceededTest
  Else
    Set objNode = objNode.selectSingleNode("..")
    fncInsFail2Report iobjReport, objNode, "nccSpecific.tocItems", isAbsPath, _
      "suggested value: " & objNodeList.length
  End If
  
  Set objNode = Nothing
  Set objNodeList = Nothing
End Function

'Private Function fncSetDcIdentifierProp(objDom As MSXML2.DOMDocument40)
'Dim oNode As IXMLDOMNode
'
'  Set oNode = objDom.selectSingleNode("//meta[@name='dc:identifier']@content")
'  If Not oNode Is Nothing Then
'  xxx
'  End If
'
'End Function
