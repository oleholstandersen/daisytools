VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oValidateSmil"
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
' This object is for validation of smil files
' objReport = report object that collects all files reported by this object

Public objReport As oReport

' The main function. Arguments work as follows:
' isAbsPath = Full path + filename of file to validate
' isContext = optional context for external file validation, see oFileIntegrity
' iobjNode = optional node for external file validation, see oFileIntegrity
' Return value is "false" if an internal error occurs, otherwise it is "true"
'
Public Function fncValidate(isAbsPath As String, _
  Optional isContext As String, Optional iobjNode As Object) As Boolean
  
  Dim bolResult As Boolean
  Dim objDom As Object
  
  Dim objFileIntegrity As New oFileIntegrity            'for tests: exists, has valid name, is readable
  Dim objXmlIntegrity As New oXmlIntegrity              'for tests: is wellformed, has and is doctype
  Dim objXMLIntegrityMSXML As New oXmlIntegrityMSXML    'for dtd (given,custom) validation
  Dim objAttributeDTD As New oAttributeDTD              'for adtd validation
  
  'fncInsertTime "oValidateSmil.fncValidate"
  
  Set objReport = New oReport                           'instantiate report object for this class
  
  fncSetProgress Me, 0, 12
  
  Dim objNode As Object, sContext As String
  If Not IsMissing(iobjNode) Then Set objNode = iobjNode
  If Not IsMissing(isContext) Then sContext = isContext
  
  If Not objFileIntegrity.fncFileExists( _
    objReport, isAbsPath, sContext, objNode _
    ) Then
    objEvent.subLog ("failure in oValidateSmil.fncValidate.objFileIntegrity.fncFileExists")
    GoTo ErrorH
  End If
  fncSetProgress Me, 1, 12
  
  If Not objFileIntegrity.fncFileHasValidName( _
    objReport, isAbsPath, smil _
    ) Then
    objEvent.subLog ("failure in oValidateSmil.fncValidate.objFileIntegrity.fncFileHasValidName")
    GoTo ErrorH
  End If
  fncSetProgress Me, 2, 12
  
  If Not objFileIntegrity.fncFileIsReadable( _
    objReport, isAbsPath _
    ) Then
    objEvent.subLog ("failure in oValidateSmil.fncValidate.objFileIntegrity.fncFileIsReadable")
    GoTo ErrorH
  End If
  fncSetProgress Me, 3, 12
  
  If Not objXmlIntegrity.fncIsWellformedXML( _
    objReport, isAbsPath, objDom _
    ) Then
    objEvent.subLog ("failure in oValidateSmil.fncValidate.objXMLIntegrity.fncIsWellformedXML")
    GoTo ErrorH
  End If
  fncSetProgress Me, 4, 12
  
  If Not objDom Is Nothing Then
    If Not objXmlIntegrity.fncHasAndIsDocumenttype( _
      objReport, objDom, smil10, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateSmil.fncValidate.objXMLIntegrity.fncHasAndIsDocumenttype")
      GoTo ErrorH
    End If
    fncSetProgress Me, 5, 12
  
    If Not objXMLIntegrityMSXML.documentValidatesToGivenDTD( _
      objReport, isAbsPath, "smil10.dtd", smil _
      ) Then
      objEvent.subLog ("failure in oValidateSmil.fncValidate.objXMLIntegrityElcel.documentValidatesToGivenDTD")
      GoTo ErrorH
    End If
    fncSetProgress Me, 6, 12
  
    If Not objXMLIntegrityMSXML.documentValidatesToCustomDTD( _
      objReport, isAbsPath, "d202smil.dtd", smil _
      ) Then
      objEvent.subLog ("failure in oValidateSmil.fncValidate.objXMLIntegrityElcel.documentValidatesToCustomDTD")
      GoTo ErrorH
    End If
    fncSetProgress Me, 7, 12

If Not bolLightMode Then
    If Not objAttributeDTD.fncDocumentAttributesContentValid( _
      objReport, objDom, smil, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateSmil.fncValidate.objAttributeDTD.fncDocumentAttributesContentValid")
      GoTo ErrorH
    End If
End If 'Not bolLightMode
    fncSetProgress Me, 8, 12
    If bolCancelValidation Then fncValidate = True: GoTo ErrorH

If Not bolLightMode Then
    If Not objAttributeDTD.fncIntraElementAttributesCorelationValid( _
      objReport, objDom, smil, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateSmil.fncValidate.objAttributeDTD.fncIntraElementAttributesCorelationValid")
      GoTo ErrorH
    End If
End If 'Not bolLightMode
    fncSetProgress Me, 9, 12
    If bolCancelValidation Then fncValidate = True: GoTo ErrorH
    
    If Not fncMotherSeqDurExists( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateSmil.fncValidate.fncMotherSeqDurExists")
      GoTo ErrorH
    End If
    fncSetProgress Me, 10, 12

If Not bolLightMode Then
    If Not fncSmilInternalDurationValues( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateSmil.fncValidate.fncMotherSeqDurValue")
      GoTo ErrorH
    End If
End If 'Not bolLightMode
    fncSetProgress Me, 11, 12
    
    If Not fncOneTextPerPar( _
      objReport, objDom, isAbsPath _
      ) Then
      objEvent.subLog ("failure in oValidateSmil.fncValidate.fncOneTextPerPar")
      GoTo ErrorH
    End If
  End If

  fncValidate = True
    
ErrorH:
  objReport.fncAddContext "smil"
  fncSetProgress Me, 12, 12
  
  Set objDom = Nothing
  Set objLastLCNode = Nothing
  Set objFileIntegrity = Nothing
  Set objXmlIntegrity = Nothing
  Set objXMLIntegrityMSXML = Nothing
  Set objAttributeDTD = Nothing
  
  'fncInsertTime "oValidateSmil.fncValidate"
End Function

' This function verifies that the mother 'seq' element has a 'dur' attribute
Private Function fncMotherSeqDurExists( _
  ByRef iobjReport As oReport, ByVal oDOMStructure As Object, _
  ByVal isAbsPath As String _
  ) As Boolean
  
  'checks that mother seq has a dur attribute
  Dim oNodeElm As Object
  Dim oNodeAtt As Object
  
  'fncInsertTime "oValidateSmil.fncMotherSeqDurExists"
  
  fncMotherSeqDurExists = False
  
  Set oNodeElm = oDOMStructure.selectSingleNode("//body/seq")
  If oNodeElm Is Nothing Then
    iobjReport.fncInsertFailedTest "smilSpecific.mSeqDurAttr", isAbsPath, _
      -1, -1
  Else
    Set oNodeElm = oNodeElm.selectSingleNode("@dur")
    If oNodeElm Is Nothing Then
      iobjReport.fncInsertFailedTest "smilSpecific.mSeqDurAttr", isAbsPath, _
        -1, -1
    End If
    iobjReport.subInsertSucceededTest
  End If
  
  Set oNodeElm = Nothing
  Set oNodeAtt = Nothing
  
  fncMotherSeqDurExists = True
  
  'fncInsertTime "oValidateSmil.fncMotherSeqDurExists"
End Function

' This function verifies that the internal duration value ('ncc:timeInThisSmil')
' of the smil file is correct.
Private Function fncSmilInternalDurationValues( _
  ByRef iobjReport As oReport, ByVal oDOMStructure As Object, _
  ByVal isAbsPath As String _
  ) As Boolean
  
  Dim oNodeList As Object
  Dim oNode As Object
  Dim oNodeAtt As Object
  Dim oNodeMap As Object

  Dim lTimeComputed As Long, lTimeGiven As Long
  Dim objCE As Object, objCB As Object
  
  'fncInsertTime "oValidateSmil.fncSmilInternalDurationValues"
  
  fncSmilInternalDurationValues = False
    
  'check that mother seq dur attr of smilfile attr corresponds to length of children
  'create sum of time in this smil: lngTimeCount to smil totaltime in ms
  Set oNodeList = oDOMStructure.getElementsByTagName("audio")
    
  For Each oNode In oNodeList
    Set oNodeMap = oNode.Attributes
      
    Set objCE = Nothing
    Set objCB = Nothing
      
    Set objCE = oNodeMap.getNamedItem("clip-end")
    Set objCB = oNodeMap.getNamedItem("clip-begin")
      
    If Not (objCE Is Nothing Or objCB Is Nothing) Then _
      lTimeComputed = lTimeComputed + fncConvertSmilClockVal2Ms(objCE.text) - _
        fncConvertSmilClockVal2Ms(objCB.text)
       
  Next
    
  'set path to mother seq dur attribute
  Set oNodeAtt = oDOMStructure.documentElement.selectSingleNode("body/seq/@dur")
    
  'convert mother seq dur attr value to ms
  If Not oNodeAtt Is Nothing Then
    lTimeGiven = fncConvertSmilClockVal2Ms(oNodeAtt.text)
      
    'compare computed value with given value
'      If lTimeGiven <> lTimeComputed Then
    If Not fncTimeCompare(lTimeComputed, lTimeGiven) Then
      Set oNode = oNodeAtt.selectSingleNode("..")
      fncInsFail2Report iobjReport, oNode, "smilSpecific.mSeqDurAttrVal", _
        isAbsPath, "suggested value: " & fncConvertMS2SmilClockVal( _
        lTimeComputed, SCV_TimeCount_s)
    Else
        iobjReport.subInsertSucceededTest
    End If
    
  'checks that meta ncc:TimeInThisSmil corresponds to length of children
    Set oNode = oDOMStructure.selectSingleNode( _
      "//meta[@name='ncc:timeInThisSmil']/@content")
    If Not oNode Is Nothing Then
      If fncSmilClockIsEqual(CStr(lTimeComputed) & "ms", oNode.nodeValue) Then
        iobjReport.subInsertSucceededTest
      Else
        Set oNode = oNode.selectSingleNode("..")
        fncInsFail2Report iobjReport, oNode, "smilSpecific.timeInThisSmil", _
          isAbsPath, "suggested value: " & fncConvertMS2SmilClockVal( _
          lTimeComputed, SCV_FullClock)
      End If
    End If
  End If

  Set oNode = Nothing
  Set oNodeAtt = Nothing
  Set oNodeMap = Nothing
  Set oNodeList = Nothing

  fncSmilInternalDurationValues = True
    
  'fncInsertTime "oValidateSmil.fncSmilInternalDurationValues"
End Function

Private Function fncOneTextPerPar(ByRef iobjReport As oReport, _
  ByVal oDOMStructure As Object, ByVal isAbsPath As String) As Boolean

  Dim oNode As Object
  Dim oPars As Object
  Dim oTexts As Object
  
  On Error GoTo ErrorH
  
  Set oPars = oDOMStructure.selectNodes("//par")
  For Each oNode In oPars
    Set oTexts = oNode.selectNodes("text")
    If oTexts.length > 1 Then fncInsFail2Report iobjReport, oNode, _
      "smilSpecific.oneTextPerPar", isAbsPath
  Next oNode
  
  Set oTexts = Nothing
  Set oNode = Nothing
  Set oPars = Nothing
  
  fncOneTextPerPar = True
ErrorH:
End Function
