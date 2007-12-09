VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oValidateContent"
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
' This object is for validation of content documents
' objReport = report object that collects all files reported by this object

Public objReport As oReport

' The main function. Arguments work as follows:
' isAbsPath = Full path + filename of file to validate
' isContext = optional context for external file validation, see oFileIntegrity
' iobjNode = optional node for external file validation, see oFileIntegrity
' Return value is "false" if an internal error occurs, otherwise it is "true"
'
Public Function fncValidate(isAbsPath As String, _
  Optional isContext As Variant, Optional iobjNode As Variant) As Boolean
  fncValidate = False
  
  Dim objFileIntegrity As New oFileIntegrity, objXmlIntegrity As New oXmlIntegrity
  Dim objDom As Object, objAdtd As New oAttributeDTD
  Dim objXMLIntegrityMSXML As New oXmlIntegrityMSXML
  
  Set objReport = New oReport
  fncSetProgress Me, 0, 7
  
  Dim objNode As Object, sContext As String
  If Not IsMissing(iobjNode) Then Set objNode = iobjNode
  If Not IsMissing(isContext) Then sContext = isContext
  
  If Not objFileIntegrity.fncFileExists( _
    objReport, isAbsPath, sContext, objNode _
    ) Then
    objEvent.subLog ("failure in oValidateContent.fncValidate.objFileIntegrity.fncFileExists")
    GoTo ErrorH
  End If
  fncSetProgress Me, 1, 8
  
  If Not objFileIntegrity.fncFileHasValidName( _
    objReport, isAbsPath, smilMediaObText _
    ) Then GoTo ErrorH
  fncSetProgress Me, 2, 8
  
  If Not objFileIntegrity.fncFileIsReadable( _
    objReport, isAbsPath _
    ) Then GoTo ErrorH
  fncSetProgress Me, 3, 8
    
  If Not objXmlIntegrity.fncIsWellformedXML( _
    objReport, isAbsPath, objDom _
    ) Then GoTo ErrorH
    
  If objDom Is Nothing Then 'if it couldnt be parsed no sense in continuing
    fncValidate = True
    GoTo ErrorH
  End If
  
  fncSetProgress Me, 4, 8
  
  If Not objXmlIntegrity.fncHasAndIsDocumenttype( _
    objReport, objDom, xhtml10, isAbsPath _
    ) Then GoTo ErrorH
  fncSetProgress Me, 5, 8

  'mg20040804, check for bodyref
  'Stop
  Dim sDtdToUse As String
  If bDTDIsExtendedInternallyForSkippableStructures(objDom.xml) Then
    sDtdToUse = "xhtml1-transitional-bodyref.dtd"
  Else
    sDtdToUse = "xhtml1-transitional.dtd"
  End If
  
  If Not objXMLIntegrityMSXML.documentValidatesToGivenDTD( _
    objReport, isAbsPath, sDtdToUse, smilMediaObText _
    ) Then GoTo ErrorH
  fncSetProgress Me, 6, 8

  If Not bolLightMode Then
    If Not objAdtd.fncDocumentAttributesContentValid( _
      objReport, objDom, smilMediaObText, isAbsPath _
      ) Then GoTo ErrorH
  End If 'Not bolLightMode
  
  fncSetProgress Me, 7, 8
  If bolCancelValidation Then fncValidate = True: GoTo ErrorH
      
If Not bolLightMode Then
  If Not objAdtd.fncIntraElementAttributesCorelationValid( _
    objReport, objDom, smilMediaObText, isAbsPath _
    ) Then GoTo ErrorH
End If 'Not bolLightMode

  fncValidate = True
ErrorH:
  objReport.fncAddContext "content"
  fncSetProgress Me, 8, 8
  Set objLastLCNode = Nothing
End Function

Private Function bDTDIsExtendedInternallyForSkippableStructures( _
  ByRef sXml As String) As Boolean
  Dim lTest1 As Long, lTest2 As Long, lTest3 As Long, lTest4 As Long, lTest5 As Long
  
  On Error GoTo ErrH
  
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
 
ErrH:
End Function

