VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oValidateDTB"
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
' This object is for validation of full dtb filesets
' objReport = report object that collects all files reported by this object

Public objReport As oReport

' The main function. Arguments work as follows:
' isAbsPath = Full path of the DTB directory
' Return value is "false" if an internal error occurs, otherwise it is "true"
'
Public Function fncValidate(isAbsPath As String) As Boolean
  Dim objValidateNcc As New oValidateNcc
  Dim objNccSmilRel As New oNccSmilRelations
  Dim objValidateMasterSmil As New oValidateMasterSmil
  Dim objMasterSmilSmilRel As New oMasterSmilRel
  
  Dim sNccName As String, sMasterSmilName As String
  
  fncValidate = False
  'fncInsertTime "oValidateDTB.fncValidate"
  
  Set objReport = New oReport

' if fncGetPreferedFileName fails (see mGlobal) the error will be reported
' for "ncc.html"
  sNccName = "ncc.html"
  fncGetPreferedFileName isAbsPath, sNccName, "ncc.html", "ncc.htm"
  
'First we validate the ncc internally
  fncSetProgress Me, 0, 6
  
  If Not objValidateNcc.fncValidate(isAbsPath & sNccName) Then GoTo ErrorH
  objReport.fncMergeReportsWithContext objValidateNcc.objReport
  fncSetProgress Me, 1, 6
  If bolCancelValidation Then fncValidate = True: GoTo ErrorH

'Then we validate it's external references, (see oValNccSmilRelations.Validate for info)

  objNccSmilRel.fncValidate objReport, isAbsPath & sNccName
  Set objNccSmilRel = Nothing
  fncSetProgress Me, 2, 6
  If bolCancelValidation Then fncValidate = True: GoTo ErrorH

'Now it's time to validate the master smil file
  'mg 20030827, no mastersmil tests in light mode
  If Not bolLightMode Then
    If Not objValidateMasterSmil.fncValidate( _
      isAbsPath & "master.smil") Then GoTo ErrorH
    fncSetProgress Me, 3, 6
    If bolCancelValidation Then fncValidate = True: GoTo ErrorH
       
    objReport.fncMergeReportsWithContext objValidateMasterSmil.objReport
    Set objValidateMasterSmil = Nothing
    
    If fncGetPreferedFileName(isAbsPath, sMasterSmilName, "master.smil", _
      "master.smi") Then
    
      If Not objMasterSmilSmilRel.fncValidate( _
        objReport, isAbsPath & sMasterSmilName) Then GoTo ErrorH
    End If
  End If ' not bolLightMode
  
  fncSetProgress Me, 4, 6
  If bolCancelValidation Then fncValidate = True: GoTo ErrorH

  'fncDistTests objReport, isAbsPath, sNccName
  Dim objDistribution As New oDistribution
  objDistribution.fncTestSingleDTBFileSet objReport, isAbsPath, sNccName
  fncSetProgress Me, 5, 6
    
  fncValidate = True
ErrorH:
  fncSetProgress Me, 6, 6
  objReport.fncAddContext "dtb"
  
  Set objLastLCNode = Nothing
  'fncInsertTime "oValidateDTB.fncValidate"
  'fncFinalizeStatistics
  Debug.Print "fncValidate exit normally"
End Function
