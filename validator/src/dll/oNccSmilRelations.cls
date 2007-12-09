VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oNccSmilRelations"
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

' This is the main function for the Ncc Smil Relations tests. "isAbsPath" must be
' the full path + filename to the NCC, all tests will be inserted in "iobjReport"
' that must be an allready initialized variable of class "oReport"
'
Public Function fncValidate( _
  iobjReport As oReport, isAbsPath As String, Optional ibolMultiVolume As Variant _
  ) As Boolean

  Dim objNccDom As Object
  Dim objXmlIntegrity As New oXmlIntegrity, objDistribution As New oDistribution
  Dim objNode As Object, objNodeHref As Object
  Dim objNode2 As Object
  
  Dim objHrefList As Object

  Dim aAlreadyCheckedObjects() As String, lObjectsInArray As Long
  Dim sCurrSrc As String
  
  Dim objValidateSmil As oValidateSmil
  Dim objSmilMediaRel As New oSmilMediaRelations
  Dim objLocalReport As New oReport
  Dim objSmilDom As Object

  Dim objBogusReport As New oReport
  
'  Dim sinSmil_ComputedTotalElapsedTime As Single
  Dim lSmil_ComputedTotalElapsedTime As Long
  Dim bolIsRel As Boolean
  
  'fncInsertTime "oNccSmilRelations.fncValidate"
  
  Dim bolMultiVolume As Boolean
  bolMultiVolume = False
  If Not IsMissing(ibolMultiVolume) Then bolMultiVolume = ibolMultiVolume

  'parse the ncc
  If Not objXmlIntegrity.fncIsWellformedXML( _
    objBogusReport, isAbsPath, objNccDom) Then fncValidate = True: GoTo ErrorH
  
  Dim objBogusNode As Object
  
  If objNccDom Is Nothing Then GoTo ErrorH
  'get all href nodes of allowed elements
  Set objHrefList = objNccDom.selectNodes("//h1/a" & _
                                          "| //h2/a" & _
                                          "| //h3/a" & _
                                          "| //h4/a" & _
                                          "| //h5/a" & _
                                          "| //h6/a" & _
                                          "| //span/a" & _
                                          "| //div/a")
  Dim lProgress As Long
  fncSetProgress Me, 0, objHrefList.length
  
  For Each objNode2 In objHrefList
    bolIsRel = False
    Set objNodeHref = objNode2.selectSingleNode("@rel")
    If Not objNodeHref Is Nothing Then bolIsRel = True
    
    Set objNodeHref = objNode2.selectSingleNode("@href")
  
    objDistribution.fncIsRelativeUri iobjReport, objNodeHref.nodeValue, objNodeHref
    sCurrSrc = fncStripIdAddPath( _
        objNodeHref.nodeValue, isAbsPath _
        ) 'get the filename from the nodevalue
    
'    DoEvents '**BugTrace
    
    If Not isObjectInArray( _
        sCurrSrc, aAlreadyCheckedObjects, lObjectsInArray _
        ) Then 'if this file is not already checked, then
        
        'validate this smil internally:
        Set objValidateSmil = Nothing
        Set objValidateSmil = New oValidateSmil
        objValidateSmil.fncValidate sCurrSrc, isAbsPath, objNode2
        DoEvents '**BugTrace
        iobjReport.fncMergeReportsWithContext objValidateSmil.objReport, _
            "nccSmilRel"
      
        'then validate this smilfiles media objects relations
        Set objLocalReport = Nothing
        Set objLocalReport = New oReport
        
        If Not objSmilMediaRel.fncValidate(objLocalReport, sCurrSrc, _
          bolMultiVolume) Then _
          Exit Function
        DoEvents '**BugTrace
        iobjReport.fncMergeReportsWithContext objLocalReport, "nccSmilRel"
      
        'then do the specific ncc-smil relations tests:
      
        'parse the smilfile
        Set objSmilDom = Nothing
        objXmlIntegrity.fncIsWellformedXML objBogusReport, sCurrSrc, objSmilDom
      
        If (Not objSmilDom Is Nothing) And (Not bolMultiVolume) Then
If Not bolLightMode Then
            'context="totalTimeIsCorrect"
            fncCheckTotalElapsedTime iobjReport, sCurrSrc, objSmilDom, _
              lSmil_ComputedTotalElapsedTime
            'sinSmil_ComputedTotalElapsedTime
                     
            fncAddToTotalElapsedTime objSmilDom, lSmil_ComputedTotalElapsedTime
            'sinSmil_ComputedTotalElapsedTime
End If 'Not bolLightMode
            'context="pointsToAcceptedElement"
            fncCheckFirstParOrText iobjReport, sCurrSrc, objSmilDom, objNccDom
        End If
        DoEvents '**BugTrace
     End If
        If Not objSmilDom Is Nothing Then
            'context="allIDsExist" + context="pointsToAcceptedElement":
            If Not aAlreadyCheckedObjects(lObjectsInArray - 1) = sCurrSrc Then
'               iobjReport.fncInsertFailedTest "nccSmilRel.smilSequenceIsSequential", isAbsPath, fncGetDOMLine(objNodeHref), fncGetDOMColumn(objNodeHref)
               Set objBogusNode = objNodeHref.selectSingleNode("..")
               fncInsFail2Report iobjReport, objBogusNode, _
                 "nccSmilRel.smilSequenceIsSequential", isAbsPath
               
'               Exit Function '???
               GoTo ErrorH
             Else
               fncCheckSmilFragment iobjReport, objNode2, objSmilDom, _
                 isAbsPath, bolMultiVolume
             End If
             DoEvents '**BugTrace
        End If
        
        lProgress = lProgress + 1
        If Not lProgress = objHrefList.length Then _
          fncSetProgress Me, lProgress, objHrefList.length
          
        If bolCancelValidation Then
          fncValidate = True
          fncSetProgress Me, objHrefList.length, objHrefList.length
'          Exit Function
          GoTo ErrorH
        End If
  
    DoEvents '**BugTrace
  Next
      fncSetProgress Me, objHrefList.length, objHrefList.length
      
      If bolMultiVolume Then GoTo ErrorH

If Not bolLightMode Then
      'Compare the computed total time of smils with the meta data found in ncc
      Set objNode = objNccDom.selectSingleNode( _
        "//meta[@name='ncc:totalTime']/@content")
      If Not objNode Is Nothing Then
        Dim lSmilMeta_TotalElapsedTime As Long
        lSmilMeta_TotalElapsedTime = fncConvertSmilClockVal2Ms(objNode.nodeValue)
        If fncSmilClockIsEqual(CStr(lSmilMeta_TotalElapsedTime / 1000), _
             CLng(lSmil_ComputedTotalElapsedTime / 1000)) Then
          iobjReport.subInsertSucceededTest
        Else
          Set objNode = objNode.selectSingleNode("..")
          fncInsFail2Report iobjReport, objNode, "nccSmilRel.totalTimeIsCorrect", _
            isAbsPath, "suggested value: " & fncConvertMS2SmilClockVal( _
              lSmil_ComputedTotalElapsedTime, SCV_FullClock, False)
        End If
      End If
End If 'not bolLightMode

  fncValidate = True
ErrorH:
  
  Set objNccDom = Nothing
  Set objXmlIntegrity = Nothing
  Set objDistribution = Nothing
  Set objNode = Nothing
  Set objNodeHref = Nothing
  Set objNode2 = Nothing
  Set objHrefList = Nothing
  Set objValidateSmil = Nothing
  Set objSmilMediaRel = Nothing
  Set objLocalReport = Nothing
  Set objSmilDom = Nothing
  Set objBogusReport = Nothing
  
  'fncInsertTime "oNccSmilRelations.fncValidate"
End Function

' This object check wheter a file exists in the given file array, if not it inserts
' the given file.
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

' This function compares the 'ncc:totalElapsedTime' value within the give smil DOM
' with the value calculated from previously opened SMIL files.
Private Function fncCheckTotalElapsedTime( _
  ByRef iobjReport As oReport, ByVal isSmilAbsPath As String, _
  ByVal iobjDomSmil As Object, _
  ByVal lSmil_ComputedTotalElapsedTime As Long _
  ) As Boolean
  
  Dim objNode As Object, lTemp As Long
    
  'fncInsertTime "oNccSmilRelations.fncCheckTotalElapsedTime"
    
  Set objNode = iobjDomSmil.selectSingleNode( _
    "//meta[@name='ncc:totalElapsedTime']/@content")
      
  If Not objNode Is Nothing Then
    If fncSmilClockIsEqual( _
      objNode.nodeValue, CStr(lSmil_ComputedTotalElapsedTime) & "ms") Then
          
      iobjReport.subInsertSucceededTest
    Else
    
      Set objNode = objNode.selectSingleNode("..")
      fncInsFail2Report iobjReport, objNode, "nccSmilRel.totalElapsedTimeIsValid", _
        isSmilAbsPath, "suggested value: " & fncConvertMS2SmilClockVal( _
          lSmil_ComputedTotalElapsedTime, SCV_FullClock)
    End If
  End If
  
  Set objNode = Nothing
  'fncInsertTime "oNccSmilRelations.fncCheckTotalElapsedTime"
End Function

' This function adds the total play time of the given smil file to the
' 'lSmil_ComputedTotalElapsedTime' variable supplied.
Private Function fncAddToTotalElapsedTime( _
  ByVal iobjDomSmil As Object, _
  ByRef lSmil_ComputedTotalElapsedTime As Long _
  ) As Boolean
  
  Dim objNode As Object, lTemp As Long, objAudioList As Object
  Dim objNodeCE As Object, objNodeCS As Object
  Dim lCS As Long, lCE As Long
  
  'fncInsertTime "oNccSmilRelations.fncAddToTotalElapsedTime"
  
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
        lCE = objMP.Duration * 1000
      End If
    Else
      lCE = fncConvertSmilClockVal2Ms(objNodeCE.nodeValue)
    End If
    
    'sinTemp = sinTemp + (sinCE - sinCS)
    lTemp = lTemp + lCE - lCS 'fncAddSingleWOFPErr(sinTemp, sinCE, -sinCS)
  Next objNode
  
  'sinSmil_ComputedTotalElapsedTime = sinSmil_ComputedTotalElapsedTime + sinTemp
  lSmil_ComputedTotalElapsedTime = lSmil_ComputedTotalElapsedTime + lTemp
AudioErr:
  Set objNode = Nothing
  Set objNodeCE = Nothing
  Set objNodeCS = Nothing
  Set objAudioList = Nothing
  
  'fncInsertTime "oNccSmilRelations.fncAddToTotalElapsedTime"
End Function

'Check that the first <par> or <text> element in the text media object points at
'an ID that corresponds to an ID in the Ncc
Private Function fncCheckFirstParOrText( _
  iobjReport As oReport, isAbsPath As String, _
  iobjSmilDOM As Object, iobjNccDom As Object _
  ) As Boolean
  
  Dim objNodeList As Object, objNode As Object
  Dim objHeadingNode As Object, sId As String
  
  Dim objDomStructure As Object
  
  'fncInsertTime "oNccSmilRelations.fncCheckFirstParOrText"
  
  fncCheckFirstParOrText = True
  Set objNodeList = iobjSmilDOM.selectNodes( _
    "//par[position()=1]/@id | //par[position()=1]/text[position()=1]/@id")
  
  Dim sT As String
  
  For Each objNode In objNodeList
    sId = objNode.nodeValue
    sT = "/a[@href='" & fncGetFileName(isAbsPath) & "#" & sId & "']"

    Set objHeadingNode = iobjNccDom.selectSingleNode( _
      "//h1" & sT & "| //h2" & sT & "| //h3" & sT & "| //h4" & sT & "| //h5" & _
      sT & "| //h6" & sT)

    If Not objHeadingNode Is Nothing Then
      iobjReport.subInsertSucceededTest
      GoTo ErrorH
    Else
     'Stop
    End If
  Next
  
  Set objNode = Nothing
  Set objHeadingNode = Nothing
  Set objNodeList = Nothing
  Set objDomStructure = Nothing
  
  iobjReport.fncInsertFailedTest "nccSmilRel.checkFirstParOrText", _
    isAbsPath, -1, -1, sId

ErrorH:
  'fncInsertTime "oNccSmilRelations.fncCheckFirstParOrText"
End Function

' This function checks wheter the given HREF nodes ID reference exists in the
' give smil DOM.
Private Function fncCheckSmilFragment( _
    iobjReport As oReport, iobjANode As Object, _
    iobjSmilDOM As Object, isNccAbsPath As String, _
    Optional ibolMultiVolume As Variant _
    ) As Boolean
    
  Dim sId As String
  Dim objNode As Object, bolMultiVolume As Boolean, objHrefNode As Object
  
  'fncInsertTime "oNccSmilRelations.fncCheckSmilFragment"
  
  Set objHrefNode = iobjANode.selectSingleNode("@href")
  
  bolMultiVolume = False
  If Not IsMissing(ibolMultiVolume) Then bolMultiVolume = ibolMultiVolume
  
    'first check that the id exists in the smilfile
    sId = fncGetId(objHrefNode.nodeValue)
    If sId = "" Then iobjReport.subInsertSucceededTest: GoTo ErrorH:
    
    Set objNode = iobjSmilDOM.selectSingleNode("//*[@id='" & sId & "']")
    If objNode Is Nothing Then  'if the id does not exist
      fncInsFail2Report iobjReport, iobjANode, _
        "nccSmilRel.hrefFragmentExists", isNccAbsPath, "id doesn't exist: " & sId
    Else                        'if the id does exist
        iobjReport.subInsertSucceededTest
        'and then check that it points to an accepted element (if not multivolume)
      If Not bolMultiVolume Then
        Select Case objNode.nodeName
            Case "par", "text" ',"audio" '??? for mvb promptsmils?
                iobjReport.subInsertSucceededTest
            
            Case Else
                Set objNode = objNode.selectSingleNode("..")
                fncInsFail2Report iobjReport, objNode, _
                  "nccSmilRel.hrefFragmentPointsToAllowedElement", _
                  fncStripIdAddPath(objHrefNode.nodeValue, isNccAbsPath), _
                  "id found on invalid element: " & sId
        End Select
      End If
    End If

ErrorH:
    Set objNode = Nothing
    Set objHrefNode = Nothing
  
  'fncInsertTime "oNccSmilRelations.fncCheckSmilFragment"
End Function

