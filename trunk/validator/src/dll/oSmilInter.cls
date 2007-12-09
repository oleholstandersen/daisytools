VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oSmilInter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objFGM As FilgraphManager, objMP As IMediaPosition, sAudioFile As String
Private objNccDom As MSXML2.DOMDocument40

Private Type typCM
  sAbsPath As String
  objDom As MSXML2.DOMDocument40
  enFileType As enuFileType
End Type

Public Function fncValidate( _
  iobjReport As oReport, isAbsPathNcc As String, isAbsPath() As String, _
  lAbsPathCount As Long _
  ) As Boolean
  
  'lCounter = current smil file processing in array
  'siTotalElapsedTime = total time of all processed 'seq' elements 'dur' attributes
  'siCalculated = value of current smil files 'ncc:timeInThisSmil' attribute
  'objDom = current smil files dom
  'sMediaAbsPath = current media elements media files absloute path
  'aCheckedMedia() = cache of all processed media files
  'lCMCount = count of entries in aCheckedMedia()
  
  Dim lCounter As Long, siTotalElapsedTime As Single, siCalculated As Single
  Dim objNode As IXMLDOMNode
  Dim objNodeList As IXMLDOMNodeList
  
  Dim objDom As MSXML2.DOMDocument40, objXmlIntegrity As oXmlIntegrity
  Dim sMediaAbsPath As String
  
  Dim sId As String, sDrive As String, sPath As String, sFileName As String
  Dim objNode2 As IXMLDOMNode, aCheckedMedia() As typCM, lCMCount As Long
  Dim lCounter2 As Long, objBogusReport As New oReport
  
  fncValidate = False
  
  'parse ncc file to use for checking verses smil files
  Set objXmlIntegrity = New oXmlIntegrity
  If Not objXmlIntegrity.fncIsWellformedXML( _
    objBogusReport, isAbsPathNcc, objNccDom _
    ) Then Exit Function
  
  Set objFGM = New FilgraphManager
  Set objMP = objFGM
  sAudioFile = ""
  
  fncSetProgress Me, 0, lAbsPathCount
  For lCounter = 0 To lAbsPathCount - 1
'parse current smil file
    Set objXmlIntegrity = New oXmlIntegrity
    If Not objXmlIntegrity.fncIsWellformedXML(objBogusReport, isAbsPath(lCounter), _
      objDom) Then GoTo Skip2
            
'Get ncc:totalElapsedTime to see that it matches the calculated values of all
'previous 'seq' elements 'dur' attributes
    Set objNode = objDom.selectSingleNode( _
      "//meta[name = 'ncc:totalElapsedTime']/@content")
    If Not objNode Is Nothing Then
      siCalculated = fncConvertHHMMSS2S(objNode.nodeValue)
        
      If Not CLng(siCalculated) = CLng(siTotalElapsedTime) Then
        iobjReport.fncInsertFailedTest "smilInter.totalElapsedTimeIsValid", _
          isAbsPath(lCounter), fncGetDOMLine(objNode), fncGetDOMColumn(objNode), _
          siTotalElapsedTime
      Else
        iobjReport.subInsertSucceededTest
      End If
    End If
      
'Check that first <par> or <text> element have the same ID as a heading in the NCC
    If Not fncCheckFirstParOrText( _
      iobjReport, isAbsPath(lCounter), objDom _
      ) Then GoTo Skip2
      
    Set objNodeList = objDom.selectNodes("//audio | //text")
      
    For Each objNode In objNodeList
'Derive the path to the media object
      Set objNode2 = objNode.selectSingleNode("@src")
      fncParseURI objNode2.nodeValue, sDrive, sPath, sFileName, sId, isAbsPathNcc
      sMediaAbsPath = sDrive & sPath & sFileName

'Check if it has been processed before, if so don't do the same things again
      For lCounter2 = 0 To lCMCount - 1
        If aCheckedMedia(lCounter2).sAbsPath = LCase$(sMediaAbsPath) Then Exit For
      Next lCounter2
      If (lCounter2 > lCMCount - 1) Then
        ReDim Preserve aCheckedMedia(lCMCount)
        Set aCheckedMedia(lCMCount).objDom = New MSXML2.DOMDocument40
        aCheckedMedia(lCMCount).sAbsPath = LCase$(sMediaAbsPath)
        
        lCMCount = lCMCount + 1

'Check that the media file that the media elements 'src' attribute points at exists
        If fncFileIntegrity( _
          iobjReport, isAbsPath(lCounter), objNode, _
          aCheckedMedia(lCMCount - 1).enFileType, sMediaAbsPath _
          ) Then
        
'Run file integrity tests on media file according to the type of media object
          If aCheckedMedia(lCMCount - 1).enFileType = smilMediaObAudio Then
            If Not fncFileIsValidAudioObject( _
              iobjReport, isAbsPath(lCounter), objNode, sMediaAbsPath _
              ) Then GoTo Skip
          Else
            If Not fncFileIsValidTextObject( _
              iobjReport, sMediaAbsPath, aCheckedMedia(lCMCount - 1).objDom _
              ) Then GoTo Skip
          End If
        End If
      End If

      If aCheckedMedia(lCounter2).enFileType = smilMediaObAudio Then
'Check that the audio clip pointed at from the media elements exists in the media file
        fncAudioObjectClipExists iobjReport, isAbsPath(lCounter), sMediaAbsPath, _
          objNode
      Else
'Check that the ID pointed at from the media element exists in the media file
        fncTextObjectIDExists iobjReport, isAbsPath(lCounter), objNode, _
          aCheckedMedia(lCounter2).objDom
      End If
        
Skip:
    Next
      
'Add first <seq> elements 'dur' attribute to totaltime
      Set objNode = objDom.selectSingleNode( _
        "//seq/@dur")
      If Not objNode Is Nothing Then
        siTotalElapsedTime = siTotalElapsedTime + Val(Left$(objNode.nodeValue, _
          Len(objNode.nodeValue) - 1))
      End If
Skip2:
    If Not (lCounter + 1 = lAbsPathCount) Then
      fncSetProgress Me, lCounter + 1, lAbsPathCount
    End If
    
    If bolCancelValidation Then Exit For
  Next lCounter
  fncSetProgress Me, lAbsPathCount, lAbsPathCount

'clear the filecache
  For lCounter = 0 To lCMCount - 1
    Set aCheckedMedia(lCounter).objDom = Nothing
  Next lCounter
  ReDim aCheckedMedia(0)
  
  Set objMP = Nothing
  Set objFGM = Nothing
  Set objNccDom = Nothing
  
  fncValidate = True
End Function

'File integrity tests on mediafile
Private Function fncFileIntegrity( _
  iobjReport As oReport, ByVal isAbsPath As String, ByVal iobjNode As IXMLDOMNode, _
  ByRef ienFileType As enuFileType, ByRef isMediaAbsPath As String _
  ) As Boolean
  
  fncFileIntegrity = False
  
  Dim objNode As IXMLDOMNode
  
  Select Case iobjNode.nodeName
    Case "audio"
      ienFileType = smilMediaObAudio
    Case "text"
      ienFileType = smilmediaobtext
  End Select
  
  Dim objFileIntegrity As New oFileIntegrity
  Dim objLocalReport As New oReport
  

'Check that the media file exists
  If Not objFileIntegrity.fncFileExists( _
    objLocalReport, isMediaAbsPath _
    ) Then GoTo Skip
'Check that the mediafile has valid name
  If Not objFileIntegrity.fncFileHasValidName( _
    objLocalReport, fncGetFileName(isMediaAbsPath), ienFileType _
    ) Then GoTo Skip

'If the mediafile is a textfile, check that it is readable
  If ienFileType = smilmediaobtext Then
    Dim objValidateContent As New oValidateContent
    If Not objValidateContent.fncValidate(isMediaAbsPath) Then GoTo Skip
    objLocalReport.fncMergeReportsWithContext objValidateContent.objReport
  End If
  
  fncFileIntegrity = True
Skip:
  iobjReport.fncMergeReportsWithContext objLocalReport, "smilInter"
End Function

'Check that the first <par> or <text> element in the text media object points at
'an ID that corresponds to an ID in the Ncc
Private Function fncCheckFirstParOrText( _
  iobjReport As oReport, isAbsPath As String, _
  iobjMediaDom As MSXML2.DOMDocument40 _
  ) As Boolean
  
  Dim objNodeList As IXMLDOMNodeList, objNode As IXMLDOMNode
  Dim lParCount As Long, lTextCount As Long, objNode2 As IXMLDOMNode
  Dim objNode3 As IXMLDOMNode, sId As String, sTemp As String
  
  Dim objDomStructure As MSXML2.DOMDocument40
  
  fncCheckFirstParOrText = False
  
  Set objNodeList = iobjMediaDom.selectNodes("//par | //text")
  For Each objNode In objNodeList
    If objNode.nodeName = "par" Then
      lParCount = lParCount + 1
      If lParCount = 2 Then GoTo Skip
    Else
      lTextCount = lTextCount + 1
      If lTextCount = 2 Then GoTo Skip
    End If
    
    Set objNode2 = objNode.selectSingleNode("@id")
    If Not objNode2 Is Nothing Then
      sId = objNode2.nodeValue
      Set objNode3 = objNccDom.selectSingleNode( _
        "//a[@href = '" & fncGetFileName(isAbsPath) & "#" & sId & "']")
           
      If Not objNode3 Is Nothing Then
        sTemp = objNode3.parentNode.nodeName
        If (sTemp = "h1" Or sTemp = "h2" Or sTemp = "h3" Or sTemp = "h4" Or sTemp = "h5" Or _
          sTemp = "h6") Then _
          iobjReport.subInsertSucceededTest: fncCheckFirstParOrText = True: Exit Function
      End If
    End If
Skip:
    If (lTextCount = 2) And (lParCount = 2) Then Exit For
  Next
  
  iobjReport.fncInsertFailedTest "smilInter.checkFirstParOrText", _
    isAbsPath, fncGetDOMLine(objNode), fncGetDOMColumn(objNode), sId
End Function

'Check that the text mediafile is a valid xhtml file
Private Function fncFileIsValidTextObject( _
  iobjReport As oReport, isMediaAbsPath As String, _
  ByRef iobjMediaDom As MSXML2.DOMDocument40 _
  ) As Boolean
  
  Dim objXmlIntegrity As New oXmlIntegrity, objLocalReport As New oReport
  Dim objXMLIntegrityElcel As New oXmlIntegrityElcel
  Dim objAttributeDTD As New oAttributeDTD
  
  fncFileIsValidTextObject = False
  

  If Not objXmlIntegrity.fncIsWellformedXML( _
    objLocalReport, isMediaAbsPath, iobjMediaDom _
    ) Then GoTo Skip

  If Not objXmlIntegrity.fncHasAndIsDocumenttype( _
    objLocalReport, iobjMediaDom, xhtml10, isMediaAbsPath _
    ) Then GoTo Skip
  
  If Not objXMLIntegrityElcel.documentValidatesToGivenDTD( _
    objLocalReport, isMediaAbsPath, "xhtml1-transitional.dtd" _
    ) Then GoTo Skip
  
  If Not objAttributeDTD.fncDocumentAttributesContentValid( _
    objLocalReport, iobjMediaDom, smilmediaobtext, isMediaAbsPath _
    ) Then GoTo Skip
  
  If Not objAttributeDTD.fncIntraElementAttributesCorelationValid( _
    objLocalReport, iobjMediaDom, smilmediaobtext, isMediaAbsPath _
    ) Then GoTo Skip
  
  fncFileIsValidTextObject = True
Skip:
  iobjReport.fncMergeReportsWithContext objLocalReport, "smilInter"
End Function

'Check that the ID pointed at by the media element exists in the mediafile
Private Function fncTextObjectIDExists( _
  iobjReport As oReport, isAbsPath As String, iobjNode As IXMLDOMNode, _
  iobjDom As MSXML2.DOMDocument40 _
  ) As Boolean
    
  Dim sId As String, sTemp As String, objNode As IXMLDOMNode
  
  Set objNode = iobjNode.selectSingleNode("@src")
  fncParseURI objNode.nodeValue, sTemp, sTemp, sTemp, sId
  
  Set objNode = iobjDom.selectSingleNode( _
    "//*[@id = '" & sId & "']")
  If objNode Is Nothing Then
    iobjReport.fncInsertFailedTest "smilInter.textObjectIDExists", _
      isAbsPath, fncGetDOMLine(iobjNode), fncGetDOMColumn(iobjNode)
  Else
    iobjReport.subInsertSucceededTest
  End If
End Function

'Check the the audio mediafile is valid, (is readable, is accepted format,
'has valid extension according to format)

Private Function fncFileIsValidAudioObject( _
  iobjReport As oReport, isAbsPath As String, objNode As IXMLDOMNode, _
  isMediaAbsPath As String _
  ) As Boolean
  
  Dim iFF As Integer, aData(25) As Byte, bolResult As Boolean, sFileName As String
  Dim lSyncWord As Long, lIdexID As Long, lLayer As Long, sType As String
  
'Read the first 26 bytes of the file
  iFF = FreeFile
  Open isMediaAbsPath For Binary As #iFF
  Get #iFF, , aData
  Close #iFF
  
'Check some header bits to decide wich format the file is encoded in
  lSyncWord = fncShl(aData(0), 3) Or (fncShr(aData(1), 5))
  If lSyncWord = 2047 Then
    lIdexID = fncShr(aData(1), 3) And 3
    lLayer = fncShr(aData(1), 1) And 3
    If (lIdexID = 3 Or lIdexID = 2) And (lLayer = 1 Or lLayer = 2) Then sType = "mpeg"
  Else
    lLayer = fncShl(aData(21), 8) Or aData(20)
    If lLayer = 55811 Or lLayer = 1 Then sType = "wav"
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
      If (Right$(isMediaAbsPath, 3) = "mpg" Or Right$(isMediaAbsPath, 3) = "mp3" Or Right$(isMediaAbsPath, 3) = "mp2") Or _
        (Right$(isMediaAbsPath, 4) = "mpeg") Then bolResult = True
      iobjReport.subInsertSucceededTest
    Case Else
      iobjReport.fncInsertFailedTest "smilInter.fileIsValidAudioObject", _
        isAbsPath, fncGetDOMLine(objNode), fncGetDOMColumn(objNode)
      fncFileIsValidAudioObject = False
  End Select
  
  If Not bolResult Then
    iobjReport.fncInsertFailedTest "smilInter.fileHasValidExtension", _
    isAbsPath, fncGetDOMLine(objNode), fncGetDOMColumn(objNode)
  Else
    iobjReport.subInsertSucceededTest
  End If
End Function

'Check so that the audioclip pointed at by the media object in the mediafile
'exists
Private Function fncAudioObjectClipExists( _
  iobjReport As oReport, isAbsPath As String, isMediaAbsPath As String, _
  iobjNode As IXMLDOMNode _
  ) As Boolean
  
  Dim siClipBegin As Single, siClipEnd As Single, objNode As IXMLDOMNode
  
  Set objNode = iobjNode.selectSingleNode("@clip-begin")
  If Not objNode Is Nothing Then _
    siClipBegin = fncConvertNPT2SSs(objNode.nodeValue) Else siClipBegin = -1
  
  Set objNode = iobjNode.selectSingleNode("@clip-end")
  If Not objNode Is Nothing Then _
    siClipEnd = fncConvertNPT2SSs(objNode.nodeValue) Else siClipEnd = -1
  
  On Error Resume Next
  Err.Clear
  
  If Not (isMediaAbsPath = sAudioFile) Then
    sAudioFile = isMediaAbsPath
    objFGM.RenderFile sAudioFile
  End If
  
  If Not Err.Number = 0 Then Exit Function
  
  If (siClipBegin > objMP.Duration) Or (siClipEnd > objMP.Duration) Or _
    (siClipBegin = -1) Or (siClipEnd = -1) Then
    
    iobjReport.fncInsertFailedTest "smilInter.audioObjectClipExists", _
      isAbsPath, fncGetDOMLine(iobjNode), fncGetDOMColumn(iobjNode), _
      objMP.Duration & "s"
  Else
    iobjReport.subInsertSucceededTest
  End If
End Function

Private Sub Class_Initialize()
  Set objMP = objFGM
End Sub
