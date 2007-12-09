VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oValidateDiscinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Daisy 2.02 Validator Engine
' Copyright (C) 2002 Daisy Consortium
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
'
' For information about the Daisy Consortium, visit www.daisy.org or contact
' info@mail.daisy.org. For development issues, contact markus.gylling@tpb.se or
' karl.ekdahl@tpb.se.

' *** Validate DiscInfo ***
' This object is for validation of discinfo files
' objReport = report object that collects all files reported by this object

Public objReport As oReport

' The main function. Arguments work as follows:
' isAbsPath = Full path + filename of file to validate
' Return value is "false" if an internal error occurs, otherwise it is "true"
'
Public Function fncValidate(isAbsPath As String) As Boolean
  
  Dim bolResult As Boolean
  
  Dim objFileIntegrity As New oFileIntegrity, objXmlIntegrity As New oXmlIntegrity
  Dim objXMLIntegrityMSXML As New oXmlIntegrityMSXML
  Dim objAttributeDTD As New oAttributeDTD
  Dim objDom As Object
  
  Set objReport = New oReport
  fncSetProgress Me, 0, 10
  
  If Not objFileIntegrity.fncFileExists( _
    objReport, isAbsPath) Then
    objEvent.subLog ("failure in oValidateDiscinfo.fncValidate.objFileIntegrity.fncFileExists")
    GoTo ErrorH
  End If
  fncSetProgress Me, 1, 10
    
  If Not objFileIntegrity.fncFileHasValidName( _
    objReport, isAbsPath, discinfo _
    ) Then
    objEvent.subLog ("failure in oValidateDiscinfo.fncValidate.objFileIntegrity.fncFileHasValidName")
    GoTo ErrorH
  End If
  fncSetProgress Me, 2, 10
    
  If Not objFileIntegrity.fncFileIsReadable( _
    objReport, isAbsPath _
    ) Then
    objEvent.subLog ("failure in oValidateDiscinfo.fncValidate.objFileIntegrity.fncFileIsReadable")
    GoTo ErrorH
  End If
  fncSetProgress Me, 3, 10
  
  If Not objXmlIntegrity.fncIsWellformedXML( _
    objReport, isAbsPath, objDom _
    ) Then
    objEvent.subLog ("failure in oValidateDiscinfo.fncValidate.objXMLIntegrity.fncIsWellformedXML")
    GoTo ErrorH
  End If
  fncSetProgress Me, 4, 10
  
  If objDom Is Nothing Then fncValidate = True: GoTo ErrorH
  
  If Not objXmlIntegrity.fncHasAndIsDocumenttype( _
    objReport, objDom, xhtml10, isAbsPath _
    ) Then
    objEvent.subLog ("failure in oValidateDiscinfo.fncValidate.objXMLIntegrity.fncHasAndIsDocumenttype")
    GoTo ErrorH
  End If
  fncSetProgress Me, 5, 10
  
  If Not objXMLIntegrityMSXML.documentValidatesToGivenDTD( _
    objReport, isAbsPath, "xhtml1-transitional.dtd", discinfo _
    ) Then
    objEvent.subLog ("failure in oValidateDiscinfo.fncValidate.objXMLIntegrityElcel.documentValidatesToGivenDTD")
    GoTo ErrorH
  End If
  fncSetProgress Me, 6, 10
  
  If Not objXMLIntegrityMSXML.documentValidatesToCustomDTD( _
    objReport, isAbsPath, "d202discinfo.dtd", discinfo _
    ) Then
    objEvent.subLog ("failure in oValidateDiscinfo.fncValidate.objXMLIntegrityElcel.documentValidatesToCustomDTD")
    GoTo ErrorH
  End If
  fncSetProgress Me, 7, 10

If Not bolLightMode Then
  If Not objAttributeDTD.fncDocumentAttributesContentValid( _
    objReport, objDom, discinfo, isAbsPath _
    ) Then
    objEvent.subLog ("failure in oValidateDiscinfo.fncValidate.objAttributeDTD.fncDocumentAttributesContentValid")
    GoTo ErrorH
  End If
End If 'Not bolLightMode
  fncSetProgress Me, 8, 10
  
If Not bolLightMode Then
  If Not objAttributeDTD.fncIntraElementAttributesCorelationValid( _
    objReport, objDom, discinfo, isAbsPath _
    ) Then
    objEvent.subLog ("failure in oValidateDiscinfo.fncValidate.objAttributeDTD.fncIntraElementAttributesCorelationValid")
    GoTo ErrorH
  End If
End If 'Not bolLightMode

  fncSetProgress Me, 9, 10
  If bolCancelValidation Then fncValidate = True: GoTo ErrorH
  
  If Not fncAllBooksExists( _
    objReport, isAbsPath, objDom _
    ) Then
    objEvent.subLog ("failure in oValidateDiscInfo.fncValidate.fncAllBooksExists")
    GoTo ErrorH
  End If
      
  fncValidate = True
ErrorH:
  objReport.fncAddContext "discinfo"
  fncSetProgress Me, 10, 10
  Set objLastLCNode = Nothing
End Function

' This function verifies that all books given exists
Private Function fncAllBooksExists( _
  iobjReport As oReport, isAbsPath As String, iobjDom As Object _
  ) As Boolean
  
  fncAllBooksExists = False
  
  Dim objNodeList As Object, objNode As Object
  Dim sDrive As String, sPath As String, sFile As String, sId As String
  Dim objFileIntegrity As oFileIntegrity, objDistribution As New oDistribution
  Dim objNode2 As Object
  
  Set objNodeList = iobjDom.selectNodes("//a")
  For Each objNode2 In objNodeList
    Set objNode = objNode2.selectSingleNode("@href")
  
    objDistribution.fncIsRelativeUri iobjReport, objNode.nodeValue, objNode
    fncParseURI objNode.nodeValue, sDrive, sPath, sFile, sId, isAbsPath
    
    Set objFileIntegrity = New oFileIntegrity
    objFileIntegrity.fncFileExists objReport, sDrive & sPath & sFile, _
      isAbsPath, objNode2
  Next
  
  fncAllBooksExists = True
End Function
