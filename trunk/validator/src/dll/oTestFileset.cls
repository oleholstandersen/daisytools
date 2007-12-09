VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oTestFileset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type tFile
  sAbsPath As String
  lFileSize As Long
  eFileType As enuFileType
  bolParsed As Boolean
End Type

Public Function fncTestSingleDTBFileSet( _
  iObjReport As oReport, isAbsPath As String, isNccName As String) As Boolean
  
  Dim aFiles() As tFile, lFileCount As Long, lCounter As Long
  lFileCount = 0
  
'*** Insert NCC file
  fncInsertFileInArray aFiles, lFileCount, isAbsPath & isNccName, False, ncc, 0
  
'*** Insert Master smil if existing
  Dim sMasterSmil As String
  fncGetPreferedFileName isAbsPath, sMasterSmil, "master.smil", "master.smi"
  If Not sMasterSmil = "" Then _
    fncInsertFileInArray aFiles, lFileCount, isAbsPath & sMasterSmil, False, _
      mastersmil, 0
      
  fncExtractAll aFiles, lFileCount
  
  fncCompareSmilList iObjReport, isAbsPath & isNccName, aFiles, lFileCount, ncc
  fncCompareSmilList iObjReport, isAbsPath & sMasterSmil, aFiles, lFileCount, mastersmil

  fncCompareNccMetaData iObjReport, isAbsPath & isNccName, aFiles, lFileCount
End Function

Public Function fncTestMultivolumeDTBFileSet( _
  iObjReport As oReport, isAbsPath() As String) As Boolean
  
  Dim aFiles() As tFile, lFileCount As Long, lCounter As Long
  Dim sMasterSmil As String, lDTBCount As Long
  Dim sNcc As String
  
  lFileCount = 0
  lDTBCount = UBound(isAbsPath) + 1
  
  For lCounter = 0 To lDTBCount - 1
'*** Insert NCC file
    fncGetPreferedFileName isAbsPath(lCounter), sNcc, "ncc.html", "ncc.htm"
    fncInsertFileInArray aFiles, lFileCount, isAbsPath(lCounter) & sNcc, False, ncc, 0
  
'*** Insert Master smil if existing
    fncGetPreferedFileName isAbsPath(lCounter), sMasterSmil, "master.smil", _
      "master.smi"
    If Not sMasterSmil = "" Then _
      fncInsertFileInArray aFiles, lFileCount, isAbsPath(lCounter) & sMasterSmil, False, _
        mastersmil, 0
  Next lCounter
      
  fncExtractAll aFiles, lFileCount
  
  For lCounter = 0 To lDTBCount - 1
    fncGetPreferedFileName isAbsPath(lCounter), sNcc, "ncc.html", "ncc.htm"
    fncCompareNccMetaData iObjReport, isAbsPath(lCounter) & sNcc, aFiles, _
      lFileCount
 Next lCounter
End Function

Private Function fncExtractAll( _
  aFiles() As tFile, lFileCount As Long) As Boolean

  Dim objDom As MSXML2.DOMDocument40, lCounter As Long
  
  Do
    If Not aFiles(lCounter).bolParsed Then
      fncParseFile aFiles(lCounter).sAbsPath, objDom, _
        aFiles(lCounter).lFileSize, aFiles(lCounter).eFileType
      
      Select Case aFiles(lCounter).eFileType
        Case ncc
          fncNccExtract objDom, aFiles, lFileCount, aFiles(lCounter).sAbsPath
      
        Case smil
          fncSmilExtract objDom, aFiles, lFileCount, aFiles(lCounter).sAbsPath
        
        Case mastersmil
          fncMasterSmilExtract objDom, aFiles, lFileCount, _
            aFiles(lCounter).sAbsPath
          
        Case smilMediaObText
          fncContentDocExtract objDom, aFiles, lFileCount, _
            aFiles(lCounter).sAbsPath
          
      End Select
      
      aFiles(lCounter).bolParsed = True
    End If
  
    lCounter = lCounter + 1
  Loop Until lCounter = lFileCount
End Function

Private Function fncCompareNccMetaData(iObjReport As oReport, _
  isAbsPath As String, aArray() As tFile, lArrayCount As Long) As Boolean
  
  Dim objDom As MSXML2.DOMDocument40, objNode As IXMLDOMNode, lTemp As Long
  
  fncParseFile isAbsPath, objDom, lTemp, ncc
  
  Set objNode = objDom.selectSingleNode("//meta[@name='ncc:files']/@content")
  If Not objNode Is Nothing Then
    If Not CLng(objNode.nodeValue) = lArrayCount Then
      iObjReport.fncInsertFailedTest "nccMetaFiles", isAbsPath, _
        fncGetDOMLine(objNode), fncGetDOMColumn(objNode), lArrayCount
    End If
  End If
  
  Set objNode = objDom.selectSingleNode("//meta[@name='ncc:kByteSize']/@content")
  If Not objNode Is Nothing Then
    Dim lSize As Long
    For lTemp = 0 To lArrayCount - 1
      lSize = lSize + aArray(lTemp).lFileSize
    Next lTemp
    
    If Not CLng(objNode.nodeValue) = Int(lSize / 1024) Then
      iObjReport.fncInsertFailedTest "nccMetakByteSize", isAbsPath, _
        fncGetDOMLine(objNode), fncGetDOMColumn(objNode), Int(lSize / 1024)
    End If
  End If
End Function

Private Function fncCompareSmilList(iObjReport As oReport, _
  isAbsPath As String, aArray() As tFile, lArrayCount As Long, _
  enTestFile As enuFileType) As Boolean
  
  Dim objDom As MSXML2.DOMDocument40, lTemp As Long
  Dim objNode As IXMLDOMNode, objNodeList As IXMLDOMNodeList
  Dim sAbsPath As String, bolFound As Boolean
  
  fncParseFile isAbsPath, objDom, lTemp, ncc
  
  If enTestFile = ncc Then
    Set objNodeList = objDom.selectNodes("//a/@href")
  ElseIf enTestFile = mastersmil Then
    Set objNodeList = objDom.selectNodes("//ref/@src")
  Else
    Exit Function
  End If
  
  bolFound = True
  
  lTemp = 0
  Do Until aArray(lTemp).eFileType = smil
    lTemp = lTemp + 1
    If lTemp = lArrayCount Then Exit Do
  Loop
  
  For Each objNode In objNodeList
    sAbsPath = fncGetPath(isAbsPath, objNode.nodeValue)
    If lTemp = lArrayCount Then
      bolFound = False: Exit For
    ElseIf sAbsPath = fncGetPath(isAbsPath, aArray(lTemp).sAbsPath) Then
      Do
        lTemp = lTemp + 1
        If lTemp = lArrayCount Then Exit Do
      Loop Until aArray(lTemp).eFileType = smil
    ElseIf sAbsPath = fncGetPath(isAbsPath, aArray(lTemp - 1).sAbsPath) And _
      enTestFile = ncc Then
'Do nothing
    Else
      bolFound = False: Exit For
    End If
  Next objNode
  
  If bolFound Then
    lTemp = lTemp + 1
    Do Until lTemp >= lArrayCount
      If lTemp = aArray(lTemp).eFileType = smil Then bolFound = False: Exit Do
      lTemp = lTemp + 1
    Loop
  End If
  
  If Not bolFound And enTestFile = ncc Then
    iObjReport.fncInsertFailedTest "nccSmilListCorrect", isAbsPath
  ElseIf Not bolFound And enTestFile = mastersmil Then
    iObjReport.fncInsertFailedTest "masterSmilSmilListCorrect", isAbsPath
  End If
End Function

Private Function fncExistInArray(aArray() As tFile, lArrayCount As Long, _
  sValue As String, Optional lOutput As Variant) As Boolean

  Dim lCounter As Long
  
  For lCounter = 0 To lArrayCount - 1
    If aArray(lCounter).sAbsPath = sValue Then
      fncExistInArray = True
      If Not IsMissing(lOutput) Then lOutput = lCounter
      Exit Function
    End If
  Next lCounter
End Function

Private Function fncInsertFileInArray( _
  iaFiles() As tFile, ilFileCount As Long, isAbsPath As String, _
  ibolParsed As String, ieFileType As enuFileType, ilFileSize As Long)
  
  ReDim Preserve iaFiles(ilFileCount)
  iaFiles(ilFileCount).sAbsPath = isAbsPath
  iaFiles(ilFileCount).bolParsed = ibolParsed
  iaFiles(ilFileCount).eFileType = ieFileType
  iaFiles(ilFileCount).lFileSize = ilFileSize
  
  ilFileCount = ilFileCount + 1
End Function

Private Function fncParseFile( _
  ByVal isAbsPath As String, ByRef iobjDom As MSXML2.DOMDocument40, _
  ByRef ilFileSize As Long, ByVal enFileType As enuFileType) As Boolean
  
  Dim objXmlIntegrity As New oXmlIntegrity, objBogusReport As New oReport
  Dim oFSO As Object, oFile As Object
  
  If enFileType = ncc Or enFileType = smil Or enFileType = mastersmil Or _
    enFileType = smilMediaObText Then _
    objXmlIntegrity.fncIsWellformedXML objBogusReport, isAbsPath, iobjDom
  
  On Error Resume Next
  Set oFSO = CreateObject("scripting.FileSystemObject")
  Set oFile = oFSO.GetFile(isAbsPath)
  ilFileSize = oFile.Size
  
  If Err.Number <> 0 Then fncParseFile = False Else fncParseFile = True
End Function

Private Function fncGetPath(isBasePath As String, isAbsPath As String) As String
  Dim sD As String, sP As String, sF As String, sId As String
  
  fncParseURI isAbsPath, sD, sP, sF, sId, isBasePath
  fncGetPath = sD & sP & sF
End Function

Private Function fncNccExtract( _
  ByVal iobjDom As MSXML2.DOMDocument, ByRef iaFiles() As tFile, _
  ByRef ilFileCount As Long, ByVal isBasePath As String _
  ) As Boolean
    
  Dim objNode As IXMLDOMNode, objNodeList As IXMLDOMNodeList
  Dim sAbsPath As String, lFileSize As Long, lFile As Long
  
  Set objNodeList = iobjDom.selectNodes("//a/@href")
        
  For Each objNode In objNodeList
    sAbsPath = fncGetPath(isBasePath, objNode.nodeValue)
    
    If Not fncExistInArray(iaFiles, ilFileCount, sAbsPath) Then _
      fncInsertFileInArray iaFiles, ilFileCount, sAbsPath, False, smil, 0
  Next
End Function

Private Function fncSmilExtract( _
  ByVal iobjDom As MSXML2.DOMDocument, ByRef iaFiles() As tFile, _
  ByRef ilFileCount As Long, ByVal isBasePath As String _
  ) As Boolean
  
  Dim objNode As IXMLDOMNode, objNodeList As IXMLDOMNodeList, sAbsPath As String
  Dim lFileSize As Long
  Dim eFileType As enuFileType, lFile As Long, objSrc As IXMLDOMNode
  
  Set objNodeList = iobjDom.selectNodes("//text | //audio | //img | //ref " & _
    "| //animation | //textstream | //video")
  
  For Each objNode In objNodeList
    Set objSrc = objNode.selectSingleNode("@src")
    If objSrc Is Nothing Then GoTo Skip
    sAbsPath = fncGetPath(isBasePath, objSrc.nodeValue)
  
    If Not fncExistInArray(iaFiles, ilFileCount, sAbsPath) Then
      Select Case objNode.nodeName
        Case "text": eFileType = smilMediaObText
        Case "audio": eFileType = smilMediaObAudio
        Case "img": eFileType = smilMediaObImg
        Case Else: eFileType = smilMediaObOther
      End Select
      
      If Not fncExistInArray(iaFiles, ilFileCount, sAbsPath) Then _
        fncInsertFileInArray iaFiles, ilFileCount, sAbsPath, False, eFileType, 0
    End If
Skip:
  Next objNode
End Function

Private Function fncContentDocExtract( _
  ByVal iobjDom As MSXML2.DOMDocument, ByRef iaFiles() As tFile, _
  ByRef ilFileCount As Long, ByVal isBasePath As String _
  ) As Boolean
    
  Dim objNode As IXMLDOMNode, objNodeList As IXMLDOMNodeList
  Dim sAbsPath As String, lFileSize As Long, lFile As Long
  
  Set objNodeList = iobjDom.selectNodes("//img/@src | //link[@type='text/css']/@href")
        
  For Each objNode In objNodeList
    sAbsPath = fncGetPath(isBasePath, objNode.nodeValue)
    
    If Not fncExistInArray(iaFiles, ilFileCount, sAbsPath) Then _
      fncInsertFileInArray iaFiles, ilFileCount, sAbsPath, False, _
      smilMediaObOther, 0
  Next
End Function

Private Function fncMasterSmilExtract( _
  ByVal iobjDom As MSXML2.DOMDocument, ByRef iaFiles() As tFile, _
  ByRef ilFileCount As Long, ByVal isBasePath As String _
  ) As Boolean
    
  Dim objNode As IXMLDOMNode, objNodeList As IXMLDOMNodeList
  Dim sAbsPath As String, lFileSize As Long, lFile As Long
  
  Set objNodeList = iobjDom.selectNodes("//ref/@src")
        
  For Each objNode In objNodeList
    sAbsPath = fncGetPath(isBasePath, objNode.nodeValue)
    
    If Not fncExistInArray(iaFiles, ilFileCount, sAbsPath) Then _
      fncInsertFileInArray iaFiles, ilFileCount, sAbsPath, False, smil, 0
  Next
End Function

