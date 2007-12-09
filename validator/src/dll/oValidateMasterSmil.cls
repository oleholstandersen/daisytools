VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oValidateMasterSmil"
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
' This object is for validation of master smil files
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
  Dim objAttributeDTD As New oAttributeDTD, objDom As Object
  
  fncValidate = False
  
  Set objReport = New oReport
  fncSetProgress Me, 0, 9
  
  If Not objFileIntegrity.fncFileExists( _
    objReport, isAbsPath) Then
    objEvent.subLog ("failure in oValidateMasterSmil.fncValidate.objFileIntegrity.fncFileExists")
    GoTo ErrorH
  End If
  fncSetProgress Me, 1, 9
      
  'Since the master smil is optional: if it doesnt exist: do no more tests
  
  Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
  If oFSO.fileexists(isAbsPath) Then
  
      If Not objFileIntegrity.fncFileHasValidName( _
        objReport, isAbsPath, mastersmil _
        ) Then
        objEvent.subLog ("failure in oValidateMasterSmil.fncValidate.objFileIntegrity.fncFileHasValidName")
        GoTo ErrorH
      End If
      fncSetProgress Me, 2, 9
      
      If Not objFileIntegrity.fncFileIsReadable( _
        objReport, isAbsPath _
        ) Then
        objEvent.subLog ("failure in oValidateMasterSmil.fncValidate.objFileIntegrity.fncFileIsReadable")
        GoTo ErrorH
      End If
      fncSetProgress Me, 3, 9
      
      If Not objXmlIntegrity.fncIsWellformedXML( _
        objReport, isAbsPath, objDom _
        ) Then
        objEvent.subLog ("failure in oValidateMasterSmil.fncValidate.objXMLIntegrity.fncIsWellformedXML")
        GoTo ErrorH
      End If
      fncSetProgress Me, 4, 9
      
      If objDom Is Nothing Then 'if it couldnt be parsed no sense in continuing
        fncValidate = True
        GoTo ErrorH
      End If
      
      If Not objXmlIntegrity.fncHasAndIsDocumenttype( _
        objReport, objDom, smil10, isAbsPath _
        ) Then
        objEvent.subLog ("failure in oValidateMasterSmil.fncValidate.objXMLIntegrity.fncHasAndIsDocumenttype")
      End If
      fncSetProgress Me, 5, 9
      
      If Not objXMLIntegrityMSXML.documentValidatesToGivenDTD( _
        objReport, isAbsPath, "smil10.dtd", smil _
        ) Then
        objEvent.subLog ("failure in oValidateMasterSmil.fncValidate.objXMLIntegrityElcel.documentValidatesToGivenDTD")
      End If
      fncSetProgress Me, 6, 9
      
      If Not objXMLIntegrityMSXML.documentValidatesToCustomDTD( _
        objReport, isAbsPath, "d202msmil.dtd", mastersmil _
        ) Then
        objEvent.subLog ("failure in oValidateMasterSmil.fncValidate.objXMLIntegrityElcel.documentValidatesToCustomDTD")
      End If
      fncSetProgress Me, 7, 9
If Not bolLightMode Then
      If Not objAttributeDTD.fncDocumentAttributesContentValid( _
        objReport, objDom, mastersmil, isAbsPath _
        ) Then
        objEvent.subLog ("failure in oValidateMasterSmil.fncValidate.objAttributeDTD.fncDocumentAttributesContentValid")
      End If
End If 'Not bolLightMode
      fncSetProgress Me, 8, 9
      If bolCancelValidation Then fncValidate = True: GoTo ErrorH

If Not bolLightMode Then
      If Not objAttributeDTD.fncIntraElementAttributesCorelationValid( _
        objReport, objDom, mastersmil, isAbsPath _
        ) Then
        objEvent.subLog ("failure in oValidateMasterSmil.fncValidate.objAttributeDTD.fncIntraElementAttributesCorelationValid")
      End If
End If 'Not bolLightMode
  End If
  fncValidate = True
ErrorH:
  objReport.fncAddContext "mastersmil"
  fncSetProgress Me, 9, 9

  Set objDom = Nothing
  Set oFSO = Nothing
  Set objLastLCNode = Nothing
End Function
