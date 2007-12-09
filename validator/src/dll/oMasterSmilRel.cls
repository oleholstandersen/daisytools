VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oMasterSmilRel"
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

' This is the main validation function for the Master Smil relationship tests,
' The reported errors will be inserted into the supplied "iobjReport" that must
' allready be initialized. The isAbsPath must contain the full path and filename to
' the master smil file
'
Public Function fncValidate( _
  iobjReport As oReport, isAbsPath As String, Optional ibolMultiVolume As Variant _
  ) As Boolean

  Dim objMasterSmilDom As Object
  Dim objXmlIntegrity As New oXmlIntegrity, objFileIntegrity As New oFileIntegrity
  Dim objDistribution As New oDistribution
  Dim objNode As Object, objNodeRef As Object
  
  Dim objRefList As Object
  Dim objSmilDom As Object
  Dim objLocalReport As oReport
  
  Dim objBogusReport As New oReport
  Dim objNode2 As Object
  Dim lSmil_ComputedTotalElapsedTime As Long
  
  Dim lTemp As Long, sCurrSrc As String
  
  'fncInsertTime "oMasterSmilRelations.fncValidate"
  
  Dim bolMultiVolume As Boolean
  bolMultiVolume = False
  If Not IsMissing(ibolMultiVolume) Then bolMultiVolume = ibolMultiVolume
    
  'Since the master smil is optional: if it doesnt exist: do no more tests
  Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
  If oFSO.fileexists(isAbsPath) Then
  
    If Not objXmlIntegrity.fncIsWellformedXML( _
      objBogusReport, isAbsPath, objMasterSmilDom) Then fncValidate = True: Exit Function
    
    ' Select all "ref" elements and check all SMIL files that they're pointing at
    If Not objMasterSmilDom Is Nothing Then
        Set objRefList = objMasterSmilDom.selectNodes("//ref")
        
        Dim lProgress As Long
        fncSetProgress Me, 0, objRefList.length
        For Each objNode2 In objRefList
          Set objNodeRef = objNode2.selectSingleNode("@src")
        
          objDistribution.fncIsRelativeUri iobjReport, objNodeRef.nodeValue, objNodeRef
          sCurrSrc = fncStripIdAddPath(objNodeRef.nodeValue, isAbsPath)
        
          Set objLocalReport = New oReport
          
          objFileIntegrity.fncFileExists objLocalReport, sCurrSrc, isAbsPath, _
            objNode2
          If (objLocalReport.lFailedTestCount > 0) And (bolMultiVolume) Then _
            GoTo SkipThis
            
          objFileIntegrity.fncFileHasValidName objLocalReport, sCurrSrc, smil
          objFileIntegrity.fncFileIsReadable objLocalReport, sCurrSrc
          
          iobjReport.fncMergeReportsWithContext objLocalReport, "masterSmilSmilRel"
          
          objXmlIntegrity.fncIsWellformedXML objBogusReport, sCurrSrc, objSmilDom
                                                         
          If Not objSmilDom Is Nothing Then _
            fncAddToTotalElapsedTime objSmilDom, lSmil_ComputedTotalElapsedTime
          
          lProgress = lProgress + 1
          If Not lProgress = objRefList.length Then _
            fncSetProgress Me, lProgress, objRefList.length
            
          If bolCancelValidation Then
            fncSetProgress Me, objRefList.length, objRefList.length
            fncValidate = True
            Exit Function
          End If
SkipThis:
        Next
        fncSetProgress Me, objRefList.length, objRefList.length
      
        'Compare the computed total time of smils with the meta data found in ncc
        If bolMultiVolume Then fncValidate = True: Exit Function
If Not bolLightMode Then
        Set objNode = objMasterSmilDom.selectSingleNode( _
          "//meta[@name='ncc:timeInThisSmil']/@content")
        If Not objNode Is Nothing Then
          Dim lSmilMeta_TotalElapsedTime As Long
          lSmilMeta_TotalElapsedTime = fncConvertSmilClockVal2Ms(objNode.nodeValue)
          If fncSmilClockIsEqual(CStr(lSmilMeta_TotalElapsedTime / 1000), _
               CStr(lSmil_ComputedTotalElapsedTime / 1000)) Then
            iobjReport.subInsertSucceededTest
          Else
            Set objNode = objNode.selectSingleNode("..")
            fncInsFail2Report iobjReport, objNode, "masterSmilSmilRel.timeInThisSmil", _
              isAbsPath, "suggested value: " & fncConvertMS2SmilClockVal( _
              lSmil_ComputedTotalElapsedTime, SCV_FullClock)
          End If
        End If
End If 'Not bolLightMode
    End If 'Not objMasterSmilDom Is Nothing
  End If 'fso.fileexists
  fncValidate = True
  
  'fncInsertTime "oMasterSmilRelations.fncValidate"
End Function

' This function checks wheter a file is allready in the file array
'
Private Function isObjectInArray( _
  isAbsPath As String, aAlreadyCheckedObjects() As String, lObjectsInArray As Long _
  ) As Boolean
  
  Dim i As Long
  
    isObjectInArray = False
    For i = 0 To lObjectsInArray - 1
        If aAlreadyCheckedObjects(i) = isAbsPath Then
            isObjectInArray = True
            Exit For
        End If
    Next i
    
    If Not isObjectInArray Then
        ReDim Preserve aAlreadyCheckedObjects(lObjectsInArray)
        aAlreadyCheckedObjects(lObjectsInArray) = isAbsPath
        lObjectsInArray = lObjectsInArray + 1
    End If
    
End Function

' This function takes all (clipend - clipbegin) and adds them to the SMILs total
' time
'
Private Function fncAddToTotalElapsedTime( _
  ByVal iobjDomSmil As Object, _
  ByRef lSmil_ComputedTotalElapsedTime As Long _
  ) As Boolean
  
  Dim objNode As Object, lTemp As Long, objAudioList As Object
  Dim objNodeCE As Object, objNodeCS As Object
  Dim lCS As Long, lCE As Long
  
  Set objAudioList = iobjDomSmil.selectNodes("//audio")
  
  For Each objNode In objAudioList
    Set objNodeCS = objNode.selectSingleNode("@clip-begin")
    If objNodeCS Is Nothing Then _
      lCS = 0 Else lCS = fncConvertSmilClockVal2Ms(objNodeCS.nodeValue)
      
    Set objNodeCE = objNode.selectSingleNode("@clip-end")
    If objNodeCE Is Nothing Then
      Dim objTempNode As Object, objFGM As New FilgraphManager
      Dim objMP As IMediaPosition
      
      Set objTempNode = objNode.selectSingleNode("@src")
      If Not objTempNode Is Nothing Then
        On Error GoTo AudioErr
        objFGM.RenderFile fncStripIdAddPath(objTempNode.nodeValue, "")
        Set objMP = objFGM
        lCE = fncRound2(objMP.Duration * 1000)
      Else
AudioErr:
        Exit Function
      End If
    Else
      lCE = fncConvertSmilClockVal2Ms(objNodeCE.nodeValue)
    End If
    
    lTemp = lTemp + (lCE - lCS)
  Next objNode
  
  lSmil_ComputedTotalElapsedTime = lSmil_ComputedTotalElapsedTime + lTemp
End Function
