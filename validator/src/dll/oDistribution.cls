VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private sThePath As String

Private Type tFile
  sAbsPath As String
  lFileSize As Long
  eFileType As enuFileType
  bolParsed As Boolean
End Type

' This function gathers the ncc and master smil (if existing) for a single DTB
' and runs two tests on them
'
Public Function fncTestSingleDTBFileSet( _
  iobjReport As oReport, isAbsPath As String, isNccName As String) As Boolean
  
  Dim aFiles() As tFile, lFileCount As Long, lCounter As Long
  
  'fncInsertTime "oDistribution.fncTestSingleDTBFileSet"
  
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
  
  fncCompareSmilList iobjReport, isAbsPath & isNccName, aFiles, lFileCount, ncc
  If Not sMasterSmil = "" Then fncCompareSmilList iobjReport, isAbsPath & _
    sMasterSmil, aFiles, lFileCount, mastersmil

If Not bolLightMode Then
  fncCompareNccMetaData iobjReport, isAbsPath & isNccName, aFiles, lFileCount
End If

  'fncInsertTime "oDistribution.fncTestSingleDTBFileSet"
End Function

' This function gathers the ncc(s) and master smil(s) (if existing) for a
' multivolume DTB and runs two tests on them
'
Public Function fncTestMultivolumeDTBFileSet( _
  iobjReport As oReport, isAbsPath() As String) As Boolean
  
  Dim aFiles() As tFile, lFileCount As Long, lCounter As Long
  Dim sMasterSmil As String, lDTBCount As Long
  Dim sNcc As String
  
  'fncInsertTime "oDistribution.fncTestMultivolumeDTBFileSet"
  
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
    fncCompareNccMetaData iobjReport, isAbsPath(lCounter) & sNcc, aFiles, _
      lFileCount
  Next lCounter
 
  'fncInsertTime "oDistribution.fncTestMultivolumeDTBFileSet"
End Function

' This file extract all file references from a file and inserts them into an
' array
'
Private Function fncExtractAll( _
  aFiles() As tFile, lFileCount As Long) As Boolean

  Dim objDom As Object, lCounter As Long
  
  Do
    If Not aFiles(lCounter).bolParsed Then
      fncParseFile aFiles(lCounter).sAbsPath, objDom, _
        aFiles(lCounter).lFileSize, aFiles(lCounter).eFileType
      
      If Not objDom Is Nothing Then
      Select Case aFiles(lCounter).eFileType
        Case ncc '!!!
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
      End If
      
      aFiles(lCounter).bolParsed = True
    End If
  
    lCounter = lCounter + 1
  Loop Until lCounter = lFileCount
  
  'debug: loop and print all found files:
  'Dim i As Long
  'For i = 0 To lFileCount - 1
  '  Debug.Print aFiles(i).sAbsPath
  'Next i
  'Stop
  'end debug
    
End Function

' This function compares the 'ncc:files' and 'ncc:kByteSize' meta data found in
' the ncc with the filelist for the whole DTB
'
Private Function fncCompareNccMetaData(iobjReport As oReport, _
  isAbsPath As String, aArray() As tFile, lArrayCount As Long) As Boolean
  
  Dim objDom As Object, objNode As Object, lTemp As Long
  
  fncParseFile isAbsPath, objDom, lTemp, ncc
  
  If objDom Is Nothing Then Exit Function
  
  If Not bolDisableAudioTests Then
    Set objNode = objDom.selectSingleNode("//meta[@name='ncc:files']/@content")
    If Not objNode Is Nothing Then
      If Not fncString2Integer(objNode.nodeValue) = lArrayCount Then
        Set objNode = objNode.selectSingleNode("..")
        fncInsFail2Report iobjReport, objNode, "dist.nccMetaFiles", isAbsPath, _
        "suggested value: " & CStr(lArrayCount)
      End If
    End If
  End If
  
  If Not bolDisableAudioTests Then
    Set objNode = objDom.selectSingleNode("//meta[@name='ncc:kByteSize']/@content")
    If Not objNode Is Nothing Then
      Dim lSize As Double
      For lTemp = 0 To lArrayCount - 1
        lSize = lSize + aArray(lTemp).lFileSize
      Next lTemp
    
      Dim sMetaValue As String, sCalcValue As String, lDecimals As Long
      sMetaValue = objNode.nodeValue
      sMetaValue = Replace(sMetaValue, ".", ",")
      lDecimals = Len(sMetaValue) - InStr(sMetaValue, ",")
      If (lDecimals < 0) Or lDecimals >= Len(sMetaValue) Then lDecimals = 0
  '    lDecimals = 0
    
      If Round(sMetaValue, lDecimals) <> Round((lSize / 1024), lDecimals) Then
    
        Set objNode = objNode.selectSingleNode("..")
        fncInsFail2Report iobjReport, objNode, "dist.nccMetakByteSize", isAbsPath, _
          "suggested value: " & Round(lSize / 1024, lDecimals)
      End If
    End If
  End If 'Not bolDisableAudioTests
End Function

' This function compares a file with a given filelist so that (a) all files are
' refered to and (b) all files that are refered to more than once are refered to
' continousley.
'
Private Function fncCompareSmilList(iobjReport As oReport, _
  isAbsPath As String, aArray() As tFile, lArrayCount As Long, _
  enTestFile As enuFileType) As Boolean
  
  Dim objDom As Object, lTemp As Long
  Dim objNode As Object, objNodeList As Object
  Dim sAbsPath As String, bolFound As Boolean, lBackup As Long
  Dim sMissing As String
  
  fncParseFile isAbsPath, objDom, lTemp, ncc
  
  If objDom Is Nothing Then Exit Function
  
  If enTestFile = ncc Then
    Set objNodeList = objDom.selectNodes("//a/@href")
  ElseIf enTestFile = mastersmil Then
    Set objNodeList = objDom.selectNodes("//ref/@src")
  Else
    Exit Function
  End If
  
  bolFound = True
  lBackup = 0
  
  lTemp = 0
  Do Until aArray(lTemp).eFileType = smil
    lTemp = lTemp + 1
    If lTemp = lArrayCount Then Exit Do
  Loop
  bolFound = True
  
  For Each objNode In objNodeList
    sAbsPath = fncStripIdAddPath(objNode.nodeValue, isAbsPath)
    
    If lTemp = lArrayCount Then
      bolFound = False: Exit For
    ElseIf sAbsPath = fncStripIdAddPath(aArray(lTemp).sAbsPath, isAbsPath) Then
      lBackup = lTemp
      Do
        lTemp = lTemp + 1
        If lTemp = lArrayCount Then lTemp = lBackup: Exit Do
      Loop Until aArray(lTemp).eFileType = smil
    ElseIf sAbsPath = fncStripIdAddPath(aArray(lTemp - 1).sAbsPath, isAbsPath) And _
      enTestFile = ncc Then
'Do nothing
    Else
      If Not sMissing = "" Then sMissing = sMissing & ", "
      sMissing = sMissing & sAbsPath
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
    iobjReport.fncInsertFailedTest "dist.nccSmilListCorrect", isAbsPath, , , _
      "missing: " & sMissing
  ElseIf Not bolFound And enTestFile = mastersmil Then
    iobjReport.fncInsertFailedTest "dist.masterSmilSmilListCorrect", isAbsPath, , _
      , "missing: " & sMissing
  End If
End Function

' This function checks if a file reference exists in the given array
'
Private Function fncExistInArray(aArray() As tFile, lArrayCount As Long, _
  sValue As String, Optional lOutput As Variant) As Boolean

  Dim lCounter As Long
  
  For lCounter = 0 To lArrayCount - 1
    'If aArray(lCounter).sAbsPath = sValue Then
    'mg20050310:
    If aArray(lCounter).sAbsPath = Replace$(sValue, "\", "/") Then
      fncExistInArray = True
      If Not IsMissing(lOutput) Then lOutput = lCounter
      Exit Function
    End If
  Next lCounter
End Function

' This function inserts a file reference in the given array
'
Private Function fncInsertFileInArray( _
  iaFiles() As tFile, ilFileCount As Long, isAbsPath As String, _
  ibolParsed As String, ieFileType As enuFileType, ilFileSize As Long)
  
  ReDim Preserve iaFiles(ilFileCount)
  iaFiles(ilFileCount).sAbsPath = Replace$(isAbsPath, "\", "/")
  iaFiles(ilFileCount).bolParsed = ibolParsed
  iaFiles(ilFileCount).eFileType = ieFileType
  iaFiles(ilFileCount).lFileSize = ilFileSize
  
  ilFileCount = ilFileCount + 1
End Function

' This function parses a file using MSXML and returns the DOM and size
'
Private Function fncParseFile( _
  ByVal isAbsPath As String, ByRef iobjDom As Object, _
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

' This function extracts all file references from a NCC file and puts them in the
' given array.
'
Private Function fncNccExtract( _
  ByVal iobjDom As Object, ByRef iaFiles() As tFile, _
  ByRef ilFileCount As Long, ByVal isBasePath As String _
  ) As Boolean
    
  Dim objNode As Object, objNodeList As Object
  Dim sAbsPath As String, lFileSize As Long, lFile As Long
  
  Set objNodeList = iobjDom.selectNodes("//a/@href") '[not(@rel)]
        
  For Each objNode In objNodeList
    sAbsPath = fncStripIdAddPath(objNode.nodeValue, isBasePath)
    
    If Not fncExistInArray(iaFiles, ilFileCount, sAbsPath) Then _
      fncInsertFileInArray iaFiles, ilFileCount, sAbsPath, False, smil, 0
  Next
  
  'mg20030916: ncc may reference css
  Set objNodeList = iobjDom.selectNodes("//img/@src | //link[@type='text/css']/@href")
        
  For Each objNode In objNodeList
    sAbsPath = fncStripIdAddPath(objNode.nodeValue, isBasePath)
    If Not fncExistInArray(iaFiles, ilFileCount, sAbsPath) Then _
      fncInsertFileInArray iaFiles, ilFileCount, sAbsPath, False, _
      xhtmlExtEnt, 0
  Next
  
End Function

' This function extracts all file references from a SMIL file and puts them in the
' given array.
'
Private Function fncSmilExtract( _
  ByVal iobjDom As Object, ByRef iaFiles() As tFile, _
  ByRef ilFileCount As Long, ByVal isBasePath As String _
  ) As Boolean
  
  Dim objNode As Object, objNodeList As Object, sAbsPath As String
  Dim lFileSize As Long
  Dim eFileType As enuFileType, lFile As Long, objSrc As Object
  
  Set objNodeList = iobjDom.selectNodes("//text | //audio | //img | //ref " & _
    "| //animation | //textstream | //video")
  
  For Each objNode In objNodeList
    Set objSrc = objNode.selectSingleNode("@src")
    If objSrc Is Nothing Then GoTo Skip
    sAbsPath = fncStripIdAddPath(objSrc.nodeValue, isBasePath)
  
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

' This function extracts all file references from a content document (xhtml) file
' and puts them in the given array.
'
Private Function fncContentDocExtract( _
  ByVal iobjDom As Object, ByRef iaFiles() As tFile, _
  ByRef ilFileCount As Long, ByVal isBasePath As String _
  ) As Boolean
    
  Dim objNode As Object, objNodeList As Object
  Dim sAbsPath As String, lFileSize As Long, lFile As Long
  
  Set objNodeList = iobjDom.selectNodes("//img/@src | //link[@type='text/css']/@href")
        
  For Each objNode In objNodeList
    sAbsPath = fncStripIdAddPath(objNode.nodeValue, isBasePath)
        
    If Not fncExistInArray(iaFiles, ilFileCount, sAbsPath) Then _
      fncInsertFileInArray iaFiles, ilFileCount, sAbsPath, False, xhtmlExtEnt, 0
      'mg20050310: parse css for url()
      If (LCase$(fncGetExtension(sAbsPath)) = "css") And (InStr(1, sAbsPath, "http:") < 1) Then
        fncExtractCssAuxilliary iaFiles, ilFileCount, sAbsPath
      End If
  Next
End Function

Private Function fncExtractCssAuxilliary( _
    ByRef iaFiles() As tFile, _
    ByRef ilFileCount As Long, _
    ByRef sCurrSrc As String) _
    As Boolean
    
Dim sCssUrlRefs() As String
Dim sCssUrlRefFullPath As String

  On Error GoTo ErrHandler
  fncExtractCssAuxilliary = False
  
  'check css for url() entries
  If (fncGetCssUrlValues(sCurrSrc, sCssUrlRefs)) Then
    Dim i As Long
    For i = 1 To UBound(sCssUrlRefs)
       sCssUrlRefFullPath = fncGetPathName(sCurrSrc) & sCssUrlRefs(i)
      'check integrity and add
      If fncFileExists(LCase$(sCssUrlRefFullPath)) Then
        If Not fncExistInArray(iaFiles, ilFileCount, sCssUrlRefFullPath) Then
          fncInsertFileInArray iaFiles, ilFileCount, sCssUrlRefFullPath, False, xhtmlExtEnt, 0
        End If
      Else
        
      End If '.fncFileExists
    Next i
  Else
    'no url values found, or error
  End If
  fncExtractCssAuxilliary = True
ErrHandler:
End Function

Private Function fncGetCssUrlValues(ByRef sCurrSrc As String, ByRef sCssUrlRefs() As String) As Boolean
Dim sCssFile As String
Dim bUrlFound As Boolean
Dim lCurrentUrlPos As Long
Dim lUrlOpenParen As Long
Dim lUrlCloseParen As Long
Dim sUrlString As String
Dim sCssUrlRefsCount As Long
'populate the byref input array with any found url() values
'return false on error or if array length = 0
  fncGetCssUrlValues = False
  On Error GoTo ErrHandler
  
  sCssFile = fncGetFileAsString(sCurrSrc)
  If sCssFile <> "" Then
    lCurrentUrlPos = 1
    Do
      bUrlFound = False
      lCurrentUrlPos = InStr(lCurrentUrlPos, sCssFile, "url", vbTextCompare)
      If lCurrentUrlPos > 0 Then
        bUrlFound = True
        lUrlOpenParen = InStr(lCurrentUrlPos, sCssFile, "(", vbTextCompare)
        lUrlCloseParen = InStr(lCurrentUrlPos, sCssFile, ")", vbTextCompare)
        'get the string within the paren
        sUrlString = Mid(sCssFile, lUrlOpenParen + 1, lUrlCloseParen - (lUrlOpenParen + 1))
        'Stop
        'remove single or double quotes
        sUrlString = Trim$(sUrlString)
        If (Mid(sUrlString, 1, 1) = Chr(39)) Or (Mid(sUrlString, 1, 1) = Chr(34)) Then
          'remove the leading char
          sUrlString = Mid(sUrlString, 2, Len(sUrlString))
        End If
        
        If (Mid(sUrlString, Len(sUrlString), 1) = Chr(39)) Or (Mid(sUrlString, Len(sUrlString), 1) = Chr(34)) Then
          'remove the trailing char
          sUrlString = Mid(sUrlString, 1, Len(sUrlString) - 1)
        End If
        
        'add it to array
        sCssUrlRefsCount = sCssUrlRefsCount + 1
        ReDim Preserve sCssUrlRefs(sCssUrlRefsCount)
        sCssUrlRefs(sCssUrlRefsCount) = sUrlString
        'indicate there is content on the array
        fncGetCssUrlValues = True
        'move cursor forward
        lCurrentUrlPos = lUrlCloseParen
      End If
    Loop Until bUrlFound = False
  
  End If 'sCssFile <> ""
  
  Exit Function
ErrHandler:
  fncGetCssUrlValues = False
End Function


' This function extracts all file references from a master SMIL file and puts
' them in the given array.
'
Private Function fncMasterSmilExtract( _
  ByVal iobjDom As Object, ByRef iaFiles() As tFile, _
  ByRef ilFileCount As Long, ByVal isBasePath As String _
  ) As Boolean
    
  Dim objNode As Object, objNodeList As Object
  Dim sAbsPath As String, lFileSize As Long, lFile As Long
  
  Set objNodeList = iobjDom.selectNodes("//ref/@src")
        
  For Each objNode In objNodeList
    sAbsPath = fncStripIdAddPath(objNode.nodeValue, isBasePath)
    
    If Not fncExistInArray(iaFiles, ilFileCount, sAbsPath) Then _
      fncInsertFileInArray iaFiles, ilFileCount, sAbsPath, False, smil, 0
  Next
  
  Debug.Print "fncMasterSmilExtract wasn't implemented, if errors: check here"
End Function

' This function checks if an URI is relative or absolute
'
Public Function fncIsRelativeUri( _
  iobjReport As oReport, isUri As String, iobjNode As Object _
  ) As Boolean

  Dim sDrive As String, sPath As String, sFile As String, sTemp As String

  fncParseURI isUri, sDrive, sPath, sFile, sTemp

  If Not sDrive = "" Then
    fncInsFail2Report iobjReport, iobjNode, "dist.noUrisUseAbsolutePath", isUri, _
      "not absolute path: " & isUri
  Else
    iobjReport.subInsertSucceededTest
  End If
End Function
