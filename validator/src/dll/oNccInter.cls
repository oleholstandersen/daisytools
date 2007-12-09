VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oNccInter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type typCM
  sAbsPath As String
  objDom As MSXML2.DOMDocument40
  enFileType As enuFileType
End Type

Public Function fncValidate( _
  iobjReport As oReport, isAbsPath As String _
  ) As Boolean
  
  Dim objValidateSmil As oValidateSmil
  Dim objNode2 As IXMLDOMNode, objNodeList As IXMLDOMNodeList, objNode As IXMLDOMNode
  
  Dim sFile As String, sPath As String, sDrive As String, sID As String
  Dim bolResult As Boolean, siTotalTime As Single, siCalculated As Single
  Dim lCounter As Long
  
  Dim objDom As MSXML2.DOMDocument40, objXmlIntegrity As New oXmlIntegrity
  Dim aSmilCache() As typCM, lSCCount As Long, lCounter2 As Long
  Dim sSmilFile As String, objSmilDom As MSXML2.DOMDocument40
  
  Dim objBogusReport As New oReport
  
  objXmlIntegrity.fncIsWellformedXML objBogusReport, isAbsPath, objDom
  
'Get all 'href' attributes on all 'a' elements in the current ncc
  Set objNodeList = objDom.selectNodes("//a/@href")
  
  On Error Resume Next
  
  Dim lItem As Long
  
  fncSetProgress Me, 0, objNodeList.length
  For Each objNode In objNodeList

'Parse the URI of the current 'href' attribute
    fncParseURI objNode.nodeValue, sDrive, sPath, sFile, sID, isAbsPath
    sSmilFile = sDrive & sPath & sFile

'If file has allready been parsed, skip the validation
    For lCounter2 = 0 To lSCCount - 1
      If aSmilCache(lCounter2).sAbsPath = LCase$(sSmilFile) Then Exit For
    Next lCounter2
    If lCounter2 > lSCCount - 1 Then
      ReDim Preserve aSmilCache(lCounter2)
      aSmilCache(lCounter2).sAbsPath = LCase$(sSmilFile)
      lSCCount = lSCCount + 1
    Else
      GoTo Skip
    End If
    
    Set objValidateSmil = New oValidateSmil
'To be able to continue with the forthcoming test, we'll have to validate the files
    If Not objValidateSmil.fncValidate(sSmilFile) Then GoTo Skip
    
'Store the targets domstructure in another place
    Set objXmlIntegrity = New oXmlIntegrity
    objXmlIntegrity.fncIsWellformedXML objBogusReport, aSmilCache(lCounter2).sAbsPath, _
      aSmilCache(lCounter2).objDom
    
'Get the first 'seq' elements 'dur' attribute
    Set objNode2 = aSmilCache(lCounter2).objDom.selectSingleNode("//seq/@dur")
    siCalculated = siCalculated + Val(Left$(objNode2.nodeValue, _
      Len(objNode2.nodeValue) - 1))
    
Skip:
'Check so that the ID is pointing to a <text>, <audio> or <par> element
    If Not (sID = "" Or aSmilCache(lCounter2).objDom Is Nothing) Then
      
      Set objNode2 = aSmilCache(lCounter2).objDom.selectSingleNode( _
        "//*[@id = '" & sID & "']")
      
      If objNode2 Is Nothing Then
        iobjReport.fncInsertFailedTest "nccInter.allIDsExists", _
          aSmilCache(lCounter2).sAbsPath, fncGetDOMLine(objNode), _
          fncGetDOMColumn(objNode), sID
      
      Else
        iobjReport.subInsertSucceededTest
        bolResult = False
      
        If (objNode2.nodeName = "text" Or objNode2.nodeName = "audio" Or _
          objNode2.nodeName = "par") Then bolResult = True
        
        If Not bolResult Then
          iobjReport.fncInsertFailedTest "nccInter.pointsToAcceptedElement", _
            aSmilCache(lCounter2).sAbsPath, fncGetDOMLine(objNode), _
            fncGetDOMColumn(objNode)
        Else
          iobjReport.subInsertSucceededTest
        End If
      End If
    End If

'Add all failed tests from the objvalidatesmil class to my own oReport object
    iobjReport.fncMergeReportsWithContext objValidateSmil.objReport
    
    Set objValidateSmil = Nothing
    Set objNode = Nothing
    Set objNode2 = Nothing
    
    lItem = lItem + 1
    If Not lItem = objNodeList.length Then fncSetProgress Me, lItem, objNodeList.length
    
    If bolCancelValidation Then Exit For
  Next
  fncSetProgress Me, objNodeList.length, objNodeList.length
  
  For lCounter2 = 0 To lSCCount - 1
    Set aSmilCache(lCounter2).objDom = Nothing
  Next lCounter2
  ReDim aSmilCache(0)
  
'Get the contents of the 'ncc:totaltime' attribute
  Set objNode = objDom.selectSingleNode( _
    "//meta[@name = 'ncc:totalTime']/@content")
  If Not objNode Is Nothing Then siTotalTime = fncConvertHHMMSS2S(objNode.nodeValue)

  If Not CLng(siCalculated) = CLng(siTotalTime) Then
    iobjReport.fncInsertFailedTest "nccInter.totalTimeIsCorrect", _
      isAbsPath, fncGetDOMLine(objNode), fncGetDOMColumn(objNode), _
      fncConvertSSs2HHMMSS(siCalculated)
  Else
    iobjReport.subInsertSucceededTest
  End If
End Function
