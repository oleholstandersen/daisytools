VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oValidateMultivolume"
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

' *** Validate Content ***
' This object is for validation of multivolume DTBs and their whole filesets
' objReport = report object that collects all files reported by this object

Public objReport As oReport

' The main function. Arguments work as follows:
' isAbsPath() = Full path + filename of all volumes of the DTB
' Return value is "false" if an internal error occurs, otherwise it is "true"
'
Public Function fncValidate(isAbsPath() As String) As Boolean
  Dim lCounter As Long, lDTBCount As Long
  
  Set objReport = New oReport
  
  lDTBCount = UBound(isAbsPath) + 1
  
  fncSetProgress Me, 0, lDTBCount
  For lCounter = 0 To lDTBCount - 1
    If Not fncValidateVolume(isAbsPath(lCounter)) Then Exit For
    If bolCancelValidation Then Exit For
    
    If lCounter + 1 < lDTBCount Then fncSetProgress Me, lCounter + 1, lDTBCount
  Next lCounter
  fncSetProgress Me, lDTBCount, lDTBCount

  fncCheckSetInfo isAbsPath, lDTBCount
  Dim objDistribution As New oDistribution
  objDistribution.fncTestMultivolumeDTBFileSet objReport, isAbsPath
  
  objReport.fncAddContext "multivolume"
  fncValidate = True
  Set objLastLCNode = Nothing
End Function

' This function validates the given volume
Private Function fncValidateVolume(isAbsPath As String) As Boolean
  Dim objValidateNcc As New oValidateNcc
  Dim objNccSmilRel As New oNccSmilRelations
  Dim objValidateMasterSmil As New oValidateMasterSmil
  Dim objMasterSmilSmilRel As New oMasterSmilRel
  
  Dim sNccName As String, sMasterSmilName As String
  
  fncValidateVolume = False
  
  Dim objReport2 As New oReport
  'Set objReport = New oReport
  
  sNccName = "ncc.html"
  fncGetPreferedFileName isAbsPath, sNccName, "ncc.html", "ncc.htm"
  
'First we validate the ncc internally
  fncSetProgress Me, 0, 6
  
  If Not objValidateNcc.fncValidate(isAbsPath & sNccName, True) Then GoTo ErrorH
  objReport2.fncMergeReportsWithContext objValidateNcc.objReport
  fncSetProgress Me, 1, 6
  If bolCancelValidation Then fncValidateVolume = True: GoTo ErrorH
  
'Then we validate it's external references, (see oValNccInter.Validate for info)
  objNccSmilRel.fncValidate objReport2, isAbsPath & sNccName, True
  Set objNccSmilRel = Nothing
  fncSetProgress Me, 2, 6
  If bolCancelValidation Then fncValidateVolume = True: GoTo ErrorH

'Now it's time to validate the master smil file
  If Not objValidateMasterSmil.fncValidate( _
    isAbsPath & "master.smil") Then GoTo ErrorH
  fncSetProgress Me, 3, 6
  If bolCancelValidation Then fncValidateVolume = True: GoTo ErrorH
     
  objReport2.fncMergeReportsWithContext objValidateMasterSmil.objReport
  Set objValidateMasterSmil = Nothing
  
  If fncGetPreferedFileName(isAbsPath, sMasterSmilName, "master.smil", _
    "master.smi") Then
  
    If Not objMasterSmilSmilRel.fncValidate( _
      objReport2, isAbsPath & sMasterSmilName, True) Then GoTo ErrorH
  End If
  
  fncSetProgress Me, 4, 6
  If bolCancelValidation Then fncValidateVolume = True: GoTo ErrorH

  'fncDistTests objReport, isAbsPath, sNccName
  fncSetProgress Me, 5, 6
  
  Debug.Print "No distribution tests."
  
  fncValidateVolume = True
ErrorH:
  fncSetProgress Me, 6, 6
  'objReport.fncAddContext "dtb"
  objReport.fncMergeReportsWithContext objReport2, "dtb"
End Function

Private Function fncCheckSetInfo( _
  isAbsPath() As String, lNumVolumes As Long _
  ) As Boolean
  
  Dim objDom As Object, objNode As Object
  Dim objXmlIntegrity As oXmlIntegrity, objBogusReport As New oReport

  Dim lValue1 As Long, lValue2 As Long, lCounter As Long, bolFound() As Boolean
  Dim lCNV As Long, lCTV As Long, objNode2 As Object
  
  ReDim bolFound(lNumVolumes)

  On Error Resume Next

  For lCounter = 0 To lNumVolumes - 1
    Set objDom = CreateObject("Msxml2.DOMDocument.4.0")
    Set objXmlIntegrity = New oXmlIntegrity
    objXmlIntegrity.fncIsWellformedXML objBogusReport, isAbsPath(lCounter) & "ncc.html", _
      objDom
  
    Set objNode = objDom.selectSingleNode( _
      "//meta[@name = 'ncc:setInfo']/@content")
    Set objNode2 = objNode.selectSingleNode("..")
    
    If Not objNode Is Nothing Then
      lValue1 = InStr(1, objNode.nodeValue, "of", vbBinaryCompare)
      lCNV = CLng(Right$(objNode.nodeValue, Len(objNode.nodeValue) - lValue1 - 2))
      If lCNV = lNumVolumes Then
        objReport.subInsertSucceededTest
      Else
        fncInsFail2Report objReport, objNode2, "multivolumeSpecific.numVolumes", _
          isAbsPath(lCounter), "suggested value: " & CStr(lNumVolumes)
      End If
      
      lValue2 = CLng(Left$(objNode.nodeValue, lValue1 - 1))
      If bolFound(lValue2) = False Then
        bolFound(lValue2) = True
        objReport.subInsertSucceededTest
      Else
        fncInsFail2Report objReport, objNode2, _
          "multivolumeSpecific.allreadyExist", isAbsPath(lCounter), _
          "volume already exist: " & CStr(lValue2)
      End If
    End If
  Next lCounter

  For lCounter = 1 To lNumVolumes
    If Not bolFound(lCounter) Then
      objReport.fncInsertFailedTest "multivolumeSpecific.allVolumesFound", _
        , , , lCounter
    Else
      objReport.subInsertSucceededTest
    End If
  Next

  fncCheckSetInfo = True
End Function
