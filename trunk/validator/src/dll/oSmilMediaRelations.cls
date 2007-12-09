VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oSmilMediaRelations"
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

' "aAudioObjects" is all allready checked audio files, "lAudioObjectCount" is the
' length of "aAudioObjects"
Dim aAudioObjects() As String, lAudioObjectCount As Long

' "aTextObjects" is all allready checked audio files, "lTextObjectCount" is the
' length of "aTextObjects"
Dim aTextObjects() As String, lTextObjectCount As Long

' "objFGM" is a DirectShow object for validating clip times in audio files,
' "sCurrentParseAudioFile" is the audiofile currently parsed
Dim objFGM As Object, sCurrentParsedAudioFile As String
' "objContentDom" is the last parsed text document, "sCurrentParsedContentDoc" is
' the filename of the currently parsed document
Dim objContentDom As Object, sCurrentParsedContentDoc As String

' "objFGM" and "objContentDom" are destroyed when the owning object is, by doing
' this we can cache the last parsed files trough many validations.

' This is the main function in the Smil Media Relations tests object
' Inputs are, "isAbsPath" (must contain a full path + filename to a smil file) and
' iobjReport (must be an allready initialized variable of class "oReport")
'
Public Function fncValidate( _
  iobjReport As oReport, isAbsPath As String, _
  Optional ibolMultiVolume As Variant _
  ) As Boolean
  
  Dim objSmilDom As Object
      
  Dim objXmlIntegrity As New oXmlIntegrity, objFileIntegrity As New oFileIntegrity
  Dim objDistribution As New oDistribution, objBogusReport As New oReport
  Dim objMediaObjectNodeList As Object
  Dim objNode As Object, objNodeSrc As Object
  
  Dim objDocumentSmilRelations As New oDocumentSmilRelations
  
  Dim sCurrSrc As String
  Dim bolIsLabyrintenAdpcm As Boolean
  Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
  If oFSO Is Nothing Then objEvent.subLog "Error in oSmilMediaRelations.fncValidate: " & _
    "couldn't create filesystemobject": Exit Function
  
  'fncInsertTime "oSmilMediaRelations.fncValidate"
  
  Dim bolMultiVolume As Boolean
  bolMultiVolume = False
  If Not IsMissing(ibolMultiVolume) Then bolMultiVolume = ibolMultiVolume
  
  'parse smilfile
  If Not objXmlIntegrity.fncIsWellformedXML(objBogusReport, isAbsPath, objSmilDom) Then GoTo ErrH

  If Not objSmilDom Is Nothing Then
      'create nodelist with all file references
      Set objMediaObjectNodeList = objSmilDom.selectNodes("//text | //audio | //img | //ref | //animation | //textstream | //video")
      
      Dim lProgress As Long
      fncSetProgress Me, 0, objMediaObjectNodeList.length
    
      Dim objValidateContent As oValidateContent
    
      For Each objNode In objMediaObjectNodeList
        Set objNodeSrc = objNode.selectSingleNode("@src")
        
        objDistribution.fncIsRelativeUri iobjReport, objNodeSrc.nodeValue, objNodeSrc
        sCurrSrc = fncStripIdAddPath(objNodeSrc.nodeValue, isAbsPath)
        
        Select Case objNode.nodeName
            Case "text"
                ' Do not do these tests if they have already been made on this
                ' file
                If Not isObjectInArray(sCurrSrc, aTextObjects, lTextObjectCount) Then
                    Set objValidateContent = Nothing
                    Set objValidateContent = New oValidateContent
                    
                    ' run oValidateContent on all content docs
                    objValidateContent.fncValidate sCurrSrc, isAbsPath, objNode
                    iobjReport.fncMergeReportsWithContext _
                      objValidateContent.objReport
                      
                    Set objValidateContent = Nothing
                      
                    objDocumentSmilRelations.fncCheckLinkback iobjReport, _
                      sCurrSrc, bolMultiVolume
                End If
                
                ' parse contentdoc, check that fragment ids exist
                If Not sCurrSrc = sCurrentParsedContentDoc Then
                    If Not objXmlIntegrity.fncIsWellformedXML( _
                        objBogusReport, sCurrSrc, objContentDom) Then GoTo SkipToNextMediaObject
                    sCurrentParsedContentDoc = sCurrSrc
                End If
                
                 ' if the contentdoc doesnt exist (this should already have been
                 ' detected by objvalidatecontentdoc) don't try this test
                If Not objContentDom Is Nothing Then
                    fncTextObjectIDExists iobjReport, isAbsPath, objNodeSrc, objContentDom
                End If
                
            Case "audio"
              If Not bolDisableAudioTests Then
                If Not isObjectInArray(sCurrSrc, aAudioObjects, lAudioObjectCount) Then
                    ' file exists + is readable
                    objFileIntegrity.fncFileExists iobjReport, sCurrSrc, _
                      isAbsPath, objNode
                    
                      ' file is valid audiofile + audiofile has valid extension
                      ' test if it exists first (else freeFile will create it)
                      If oFSO.FileExists(sCurrSrc) Then
                        'mg 20050330 added  hasValidName
                        objFileIntegrity.fncFileHasValidName iobjReport, sCurrSrc, smilMediaObAudio
                        fncFileIsValidAudioObject iobjReport, isAbsPath, objNode, sCurrSrc, bolIsLabyrintenAdpcm
                      Else
                        'be silent; non-existance is already reported
                      End If
                    
                End If 'Not isObjectInArray
              End If 'Not bolDisableAudioTests

              If (Not bolLightMode) Then
                If (Not bolDisableAudioTests) Then
                  ' clip exists in audiofile
                  If Not bolIsLabyrintenAdpcm Then
                    If oFSO.FileExists(sCurrSrc) Then
                      If Not fncAudioObjectClipExists( _
                        iobjReport, isAbsPath, objNode, sCurrSrc, objFGM, sCurrentParsedAudioFile _
                        ) Then GoTo ErrH
                    Else
                        'be silent; non-existance is already reported
                    End If
                  Else
                    'objEvent.subLog ("audiofile is adpcm2: clip test cancelled")
                  End If
                End If 'not bolDisableAudioTests
              End If 'Not bolLightMode
            Case Else
                    ' file exists + is readable
                    objFileIntegrity.fncFileExists iobjReport, sCurrSrc, _
                      isAbsPath, objNode
                    objFileIntegrity.fncFileIsReadable iobjReport, sCurrSrc
        End Select
SkipToNextMediaObject:
        
        lProgress = lProgress + 1
        If Not lProgress = objMediaObjectNodeList.length Then _
          fncSetProgress Me, lProgress, objMediaObjectNodeList.length
          
        If bolCancelValidation Then
          fncValidate = True
          fncSetProgress Me, objMediaObjectNodeList.length, _
            objMediaObjectNodeList.length
          GoTo ErrH
        End If
      Next
  
  fncSetProgress Me, objMediaObjectNodeList.length, objMediaObjectNodeList.length
  
  End If

  
  fncValidate = True
ErrH:
  Set objXmlIntegrity = Nothing
  Set objFileIntegrity = Nothing
  Set objDistribution = Nothing
  Set objBogusReport = Nothing
  Set objNode = Nothing
  Set objNodeSrc = Nothing
  Set objMediaObjectNodeList = Nothing
  Set objSmilDom = Nothing
  Set oFSO = Nothing
  
  iobjReport.fncAddContext "smilMediaRel"
  
  'fncInsertTime "oSmilMediaRelations.fncValidate"
End Function

' This function checks if the given file already exists in an array, if not
' it inserts the file as a new array instance.
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

' This function checks if the given ID exists in the content document
'
Private Function fncTextObjectIDExists( _
  ByRef iobjReport As oReport, ByVal isAbsPath As String, ByVal iobjNode As Object, _
  ByRef iobjDom As Object _
  ) As Boolean
    
  Dim sId As String, sTemp As String, objNode As Object
  
  'fncInsertTime "oSmilMediaRelations.fncTextObjectIDExists"
  
  fncParseURI iobjNode.nodeValue, sTemp, sTemp, sTemp, sId
  If sId = "" Then fncTextObjectIDExists = True: Exit Function
  
  Set objNode = iobjDom.selectSingleNode( _
    "//*[@id = '" & sId & "']")
  If objNode Is Nothing Then
    Set iobjNode = iobjNode.selectSingleNode("..")
    fncInsFail2Report iobjReport, iobjNode, "textObjectIDExists", isAbsPath, _
      "id doesn't exist: " & sId
  Else
    iobjReport.subInsertSucceededTest
  End If
  
  Set objNode = Nothing
    
  fncTextObjectIDExists = True
  
  'fncInsertTime "oSmilMediaRelations.fncTextObjectIDExists"
End Function

' Check the the audio mediafile is valid, (is readable, is accepted format,
' has valid extension according to format)
'
Private Function fncFileIsValidAudioObject( _
  iobjReport As oReport, isAbsPath As String, ByVal objNode As Object, _
  ByVal isMediaAbsPath As String, ByRef bolIsLabyrintenAdpcm As Boolean _
  ) As Boolean
  
  'fncInsertTime "oSmilMediaRelations.fncFileIsValidAudioObject"
  
  Err.Clear
  On Error GoTo ErrH
  'isMediaAbsPath= media object abs path; isAbsPath is smilfile path
  
  Dim iFF As Integer, aData(25) As Byte, bolResult As Boolean, sFileName As String
  Dim lSyncWord As Integer, lIdexID As Byte, lLayer As Byte, sType As String
  
  isMediaAbsPath = LCase$(isMediaAbsPath)
  
'Read the first 26 bytes of the file
  Dim lTemp As Long, lFile As Long
  lFile = FreeFile
  Open isMediaAbsPath For Binary Access Read As #lFile
  Get #lFile, , aData
  Close #lFile
'  Do
'    Get #lFile, , lByte1
'Again:
'    If lByte1 = 255 Then
'      Get #lFile, , lByte2
'      If (fncShr(lByte2, 4) And 15) = 15 Then Exit Do
'      lByte1 = lByte2
'      GoTo Again
'    End If
'  Loop Until EOF(lFile)
'  Seek #lFile, Seek(lFile) - 2
'  If Not EOF(lFile) Then Get #lFile, , aData
'  Close #lFile

  lSyncWord = fncShl(aData(0), 3) Or (fncShr(aData(1), 5))
  If lSyncWord = 2047 Then
    lIdexID = fncShr(aData(1), 3) And 3
    lLayer = fncShr(aData(1), 1) And 3
    If (lIdexID = 3 Or lIdexID = 2) And (lLayer = 1 Or lLayer = 2) Then sType = "mpeg"
  Else
    lLayer = fncShl(aData(21), 8) Or aData(20)
    If lLayer = 55811 Or lLayer = 1 Then sType = "wav" '55811= adpcm2, 1=pcm
    'add mg 20030220:
    If lLayer = 55811 Then
      bolIsLabyrintenAdpcm = True
    Else
      bolIsLabyrintenAdpcm = False
    End If
  End If
 
  bolResult = False
  fncFileIsValidAudioObject = True

'Check that the filename has an acceptable extension
'If previous test didn't say that this were either an 'wav' or 'mp3' file, this
'is not an accepted format.
  Select Case sType
    Case "wav"
      If Right$(isMediaAbsPath, 3) = "wav" Then bolResult = True
      iobjReport.subInsertSucceededTest
    Case "mpeg"
      If (Right$(isMediaAbsPath, 3) = "mp3" Or _
        Right$(isMediaAbsPath, 3) = "mp2") Then bolResult = True
      iobjReport.subInsertSucceededTest
    Case Else
      Set objNode = objNode.selectSingleNode("..")
      fncInsFail2Report iobjReport, objNode, "fileIsValidAudioObject", isAbsPath, _
        isMediaAbsPath
      fncFileIsValidAudioObject = False
  End Select
  
  If Not bolResult Then
    Set objNode = objNode.selectSingleNode("..")
    fncInsFail2Report iobjReport, objNode, "fileIsValidAudioObject", isAbsPath
  Else
    iobjReport.subInsertSucceededTest
  End If
ErrH:
  If Err.Number <> 0 Then
    Set objNode = objNode.selectSingleNode("..")
    fncInsFail2Report iobjReport, objNode, "fileIsValidAudioObject", isAbsPath
  End If
  
  fncFileIsValidAudioObject = True
  
  'fncInsertTime "oSmilMediaRelations.fncFileIsValidAudioObject"
End Function

' This function checks wheter the clip time given exists within the given audiofile
'
Private Function fncAudioObjectClipExists( _
  iobjReport As oReport, isAbsPath As String, ByVal iobjNode As Object, _
  isMediaAbsPath As String, iobjFGM As Object, isLastRenderedFile As String _
  ) As Boolean
  
  Dim lClipBegin As Long, lClipEnd As Long, objNode As Object
  Dim objMP As IMediaPosition, sClipBegin As String, sClipEnd As String
  
  'fncInsertTime "oSmilMediaRelations.fncAudioClipExists"
  
  fncAudioObjectClipExists = True
  
  Set objNode = iobjNode.selectSingleNode("@clip-begin")
  If Not objNode Is Nothing Then _
    lClipBegin = fncConvertSmilClockVal2Ms(objNode.nodeValue) Else Exit Function
  sClipBegin = objNode.nodeValue
  
  Set objNode = iobjNode.selectSingleNode("@clip-end")
  If Not objNode Is Nothing Then _
    lClipEnd = fncConvertSmilClockVal2Ms(objNode.nodeValue) Else Exit Function
  sClipEnd = objNode.nodeValue
  
  On Error Resume Next
  Err.Clear
  
  If Not (isMediaAbsPath = isLastRenderedFile) Then
    isLastRenderedFile = isMediaAbsPath
    Set iobjFGM = Nothing
    Set iobjFGM = New FilgraphManager
    iobjFGM.RenderFile isLastRenderedFile
  End If
  
  ' The following queries the iobjFGM object (that is a FilterGraphManager) for a
  ' IMediaPosition interface, the IMediaPosition interfaces handles duration,
  ' rate and position
  Set objMP = iobjFGM
  
  If Not Err.Number = 0 Then Exit Function
  
  ' Optimize by calculating the following values ONCE
  Dim lCVal As Long, lSuggested As Long
  lCVal = CLng(objMP.Duration * 1000)
  
  'lSuggested = Round(objMP.Duration, 3)
  'mg 20030827 the below instead ??
  lSuggested = lCVal
  
  If (Not fncTimeCompareLog(lClipBegin, lCVal, False)) Then
'    fncInsFail2Report iobjReport, iobjNode, "audioObjectClipBeginExists", _
'      isAbsPath, "suggested value: " & lSuggested & "s"
' mg20030910: did the below instead
     fncInsFail2Report _
       iobjReport, _
       iobjNode, _
       "audioObjectClipBeginExists", _
       isAbsPath, _
       "maximum possible value is: " & fncConvertMS2SmilClockVal(lSuggested, SCV_Npt) & "s"
  Else
    iobjReport.subInsertSucceededTest
  End If

  If (Not fncTimeCompareLog(lClipEnd, lCVal, False)) Then
'    fncInsFail2Report iobjReport, iobjNode, "audioObjectClipEndExists", _
'      isAbsPath, "suggested value: " & lSuggested & "s"
' mg20030910: did the below instead
    fncInsFail2Report _
       iobjReport, _
       iobjNode, _
       "audioObjectClipEndExists", _
       isAbsPath, _
       "maximum possible value is: " & fncConvertMS2SmilClockVal(lSuggested, SCV_Npt) & "s"

  Else
    iobjReport.subInsertSucceededTest
  End If
  
  If (lClipEnd < lClipBegin) Then
    fncInsFail2Report iobjReport, iobjNode, "audioObjectClipBeginVsEnd", _
      isAbsPath
  Else
    iobjReport.subInsertSucceededTest
  End If
  
  Set objNode = Nothing
  Set objMP = Nothing
  
  'fncInsertTime "oSmilMediaRelations.fncAudioClipExists"
End Function

Private Sub Class_Initialize()
  Set objContentDom = Nothing
  Set objFGM = Nothing
End Sub

Private Sub Class_Terminate()
 Set objContentDom = Nothing
 Set objFGM = Nothing
End Sub

Private Function fncGetBitrate(bIdexID As Byte, bLayer As Byte, _
  bBitRate As Byte) As Long
  
  Dim lTemp As Long
  
    If (bIdexID = 1 And bLayer = 3) Then
      lTemp = 32 * bBitRate
    ElseIf (bIdexID = 1 And bLayer = 2) Then
      If bBitRate < 3 Then
        lTemp = 16 + 16 * bBitRate
      ElseIf bBitRate < 4 Then
        lTemp = 56 + (8 * (bBitRate - 3))
      ElseIf bBitRate < 9 Then
        lTemp = 64 + (16 * (bBitRate - 4))
      ElseIf bBitRate < 13 Then
        lTemp = 128 + (32 * (bBitRate - 8))
      Else
        lTemp = 256 + (64 * (bBitRate - 12))
      End If
    ElseIf (bIdexID = 1 And bLayer = 1) Then
      If bBitRate < 6 Then
        lTemp = 24 + 8 * bBitRate
      ElseIf bBitRate < 9 Then
        lTemp = 64 + (16 * (bBitRate - 5))
      ElseIf bBitRate < 14 Then
        lTemp = 128 + (32 * (bBitRate - 9))
      ElseIf bBitRate = 14 Then
        lTemp = 320
      End If
    ElseIf (bIdexID = 0 And bLayer = 3) Then
      If bBitRate < 3 Then
        lTemp = 16 + (16 * bBitRate)
      ElseIf bBitRate < 5 Then
        lTemp = 48 + (8 * (bBitRate - 2))
      ElseIf bBitRate < 13 Then
        lTemp = 80 + (16 * (bBitRate - 5))
      Else
        lTemp = 224 + (32 * (bBitRate - 13))
      End If
    ElseIf (bIdexID = 0 And (bLayer = 2 Or bLayer = 1)) Then
      If bBitRate < 9 Then
        lTemp = bBitRate * 8
      Else
        lTemp = 64 + (16 * (bBitRate - 8))
      End If
    End If
      
    lTemp = lTemp * 1000
      
    If bBitRate = 0 Then lTemp = 0
    If bBitRate = 15 Then lTemp = -1
    
    fncGetBitrate = lTemp
End Function

Private Function fncGetSampleFreq(bIdexID As Byte, bSampleFreq As Byte) As Long
  Select Case bIdexID
  Case 1
    If bSampleFreq = 0 Then fncGetSampleFreq = 44100
    If bSampleFreq = 1 Then fncGetSampleFreq = 48000
    If bSampleFreq = 2 Then fncGetSampleFreq = 32000
  Case Else
    If bSampleFreq = 0 Then fncGetSampleFreq = 22050
    If bSampleFreq = 1 Then fncGetSampleFreq = 24000
    If bSampleFreq = 2 Then fncGetSampleFreq = 16000
  End Select
End Function
