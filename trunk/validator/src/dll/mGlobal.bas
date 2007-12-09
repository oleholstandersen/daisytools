Attribute VB_Name = "mGlobal"
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



' The Daisy 2.02 Validator Engine tries to use as little physical DLL bindings as
' possible, this is to avoid stupid VB messages like "ActiveX component can't
' create object" but instead give meaningful messages like "MSXML4 not installed".
' Currently the only direct bindings are MSXML4 (msxml4.dll) and ActiveMovie
' (quartz.dll). These two must be included because we're using interfaces that are
' needed during compiletime so the dll should give meaningful messages even if the
' dll:s are missing on a system.
'
' Non directly referenced dll:s are dtdparser.dll which is a custom component for
' using DTD like rule systems.
'
' The Validator Engine has several properties that can be modified trough the
' oUserControl object. This is a "non creatable global" object that can't be
' instantiated but looks like a global module from outside the dll.
'
' To receive events from the Validator Engine, an event object pointer must be
' created and have the same address as the xobjEvent found in the oUserControl
' object.


' lists all types of files in a smil fileset,
' xhtmlExtEnt is css, images etc;
' all which has an external URI
Public Enum enuFileType
    ncc
    mastersmil
    smil
    smilMediaObText
    smilMediaObAudio
    smilMediaObImg
    smilMediaObOther
    xhtmlExtEnt
    discinfo
    nccmultivolume
End Enum

' equals root element name of smil 1.0 and xhtml 1.0
Public Enum enuDocuType
    smil10
    xhtml10
End Enum

' Type for calculating the current progress
Public Type typProgressItem
  objObject As Object
  lProgress As Long
  lProgressSpan As Long
End Type

' The VTM items array and the array counter
Public aVtmItems() As oVtmItem
Public lVtmItemsCount As Long

' Different paths and options setup by the user
Public sVtmPath As String, bolVtmIsLoaded As Boolean
Public sDtdPath As String, sTempPath As String
Public sAdtdPath As String, bolAdvancedADTD As Boolean

' This flags decides wheter the doctype of all documents should be modified to
' point at a local dtd instead of validating against an online dtd
Public bolUseOnlineDtds As Boolean
    
' The event object used by the dll
Public objEvent As New oEvent

' The progress array and the counter
Public aProgressItem() As typProgressItem, lProgressCount As Long
' The calculated progress and cancel flag
Public lTotalProgress As Long, bolCancelValidation As Boolean

' The rules for the different ADTDs, loaded on dll initialization and destroyed
' on deinitialization to save parsing time
Public objRulesSmil As Object, objRulesNcc As Object
Public objRulesMasterSmil As Object, objRulesContent As Object
Public objRulesMultivolume As Object, objRulesDiscinfo As Object
Public bolRulesAreLoaded As Boolean

' Cache data for fncInsFail2Report
Public objLastLCNode As Object, lLastLine As Long, lLastColumn As Long

' the DLL physical path
Public sAppPath As String

' boolean that indicates wheter the validator is initialized or not
' **CURRENTLY UNUSED**
Public bolInitialized As Boolean

Public bolLightMode As Boolean
Public bolDisableAudioTests As Boolean

' Sets some default values
Sub Main()

    sAppPath = App.Path
    If Not Right$(sAppPath, 1) = "\" Then sAppPath = sAppPath & "\"

    'these vars may also be set from ui in oUserControl. below is default values
        
    'set the default path to the dtds
    sDtdPath = sAppPath & "externals\"
    
    'set the default path for adtd parser
    sAdtdPath = sAppPath & "externals\"
       
' *** IGNORE *** Elcel currently unused
'    'set the default online/local, proxy settings for elcel
'    bolUseOnlineDtds = False
    
    'set default value for vtm (vtm.xml, external test content)
    sVtmPath = sAppPath
    
    'set default temp path for where the validator can save temporary files
    sTempPath = sAppPath
    'end oUserControl defaults
    
    bolRulesAreLoaded = False
    bolVtmIsLoaded = False

'  ****************************************************************
'  *********************** for dll debug **************************
'  ****************************************************************
'  ***** change the properties project type to active x exe  ******
'  *****       change component starttype to standalone      ******
'  *****          decomment and mod the lines below          ******
'
'
'    sDtdPath = sAppPath & "..\ui\externals\"
'    sAdtdPath = sAppPath & "..\ui\externals\"
'    sVtmPath = sAppPath & "..\ui\"
'    fncLoadRulesFiles
'    fncParseVtm
'
'    Dim objUC As New oUserControl
'    objUC.fncDeinitializeValidator
'    Set objUC = Nothing
'
'    Stop
'
'    Dim objVal As New oValidateDTB
'    'Dim objVal As New oValidateContent
'
'    objUC.fncSetLightMode (True)
'
'    objVal.fncValidate "E:\dtbs\temp\"
'    Stop
'
'
'    Stop
'    Set objVal.objReport = Nothing
'    Stop
'    Set objVal = Nothing
'    Stop
'
'    Dim tjo As New oUserControl
'    objUC.fncDeinitializeValidator
'    Stop
'*** END TEST CODE ****

End Sub

' Loads all ADTDs
Public Function fncLoadRulesFiles()

If Not bolLightMode Then
  Set objRulesSmil = CreateObject("DTDParser.cDTDData")
  Set objRulesNcc = CreateObject("DTDParser.cDTDData")
  Set objRulesMasterSmil = CreateObject("DTDParser.cDTDData")
  Set objRulesContent = CreateObject("DTDParser.cDTDData")
  Set objRulesMultivolume = CreateObject("DTDParser.cDTDData")
  Set objRulesDiscinfo = CreateObject("DTDParser.cDTDData")

  If Not objRulesSmil.fncLoadFile(sAdtdPath & "smil.adtd") Then GoTo ErrorH
  DoEvents
  If Not objRulesNcc.fncLoadFile(sAdtdPath & "ncc.adtd") Then GoTo ErrorH
  DoEvents
  If Not objRulesMasterSmil.fncLoadFile(sAdtdPath & "msmil.adtd") Then GoTo ErrorH
  DoEvents
  If Not objRulesContent.fncLoadFile(sAdtdPath & "content.adtd") Then GoTo ErrorH
  DoEvents
  If Not objRulesMultivolume.fncLoadFile(sAdtdPath & "multivolume.adtd") Then GoTo ErrorH
  DoEvents
  If Not objRulesDiscinfo.fncLoadFile(sAdtdPath & "discinfo.adtd") Then GoTo ErrorH
  DoEvents
End If 'Not bolLightMode

  fncLoadRulesFiles = True
  bolRulesAreLoaded = True
    
  Exit Function
ErrorH:
  objEvent.subLog "Error loading ADTD files"

  Set objRulesSmil = Nothing
  Set objRulesNcc = Nothing
  Set objRulesMasterSmil = Nothing
  Set objRulesContent = Nothing
  Set objRulesMultivolume = Nothing
  Set objRulesDiscinfo = Nothing
End Function

' This function is to be called for all functions within a object that takes any
' noticable time to process. Each object identifies itself, reports the current
' test and how many total tests that are going to performed. If the previous
' object that reported progress was another or no object, the progress hierarchy
' level is increased. Each levels progress information is stored in the
' "aProgressItem" array, the current hierarchy level is stored in the variable
' "lProgressCount". If a second object would be called trough the first object,
' the tests hierarchy level will increase and will not decrease until the
' second object has reported that the last test has been performed. I.E:
'
'                             (Starting at hierarchy level 0)
' Object1.fncSetProgress 0, 2 (Hierarchy level 1)
' Object1.fncSetProgress 1, 2 (Still hierarchy level 1)
' Object2.fncSetProgress 0, 1 (Hierarchy level increases to 2)
' Object2.fncSetProgress 1, 1 (All tests of Object2 has been performed (1/1) so
'                              hierarchy level decreases to 1 again)
' Object1.fncSetProgress 2, 2 (All tests of Object1 has been performed (2/2) so
'                              hierarchy level decreases to 0)

Public Function fncSetProgress( _
  objObject As Object, lProgress As Long, lProgressSpan As Long _
  ) As Boolean
  
  Dim bolCreateNew As Boolean
  
  If lProgressCount = 0 Then
    bolCreateNew = True
  ElseIf Not objObject Is aProgressItem(lProgressCount - 1).objObject Then
    bolCreateNew = True
  End If
  
  If bolCreateNew Then
    ReDim Preserve aProgressItem(lProgressCount)
    Set aProgressItem(lProgressCount).objObject = objObject
    aProgressItem(lProgressCount).lProgressSpan = lProgressSpan
    lProgressCount = lProgressCount + 1
  End If
  
  Dim lOldProgress As Long
  lOldProgress = lTotalProgress
  
  aProgressItem(lProgressCount - 1).lProgress = lProgress
  fncCalculateTotalProgress
  
  If lProgress = aProgressItem(lProgressCount - 1).lProgressSpan Then
    lProgressCount = lProgressCount - 1
    If lProgressCount > 0 Then ReDim Preserve aProgressItem(lProgressCount - 1)
  End If
  
  If Not lTotalProgress = lOldProgress Then objEvent.subProgressChanged
End Function

' This function is calculating the current progress in percent trough going trough
' all the hierarchy levels found in the "aProgressItem" array. The value of each
' test is calculated trough ((previousObjectsPercentagePerTest) / totalTests) *
' currentTest. I.E.
'
' The first object has 4 tests, each test is then worth 25%. Inbetween the 2nd and
' 3rd test a new object is called, this object has 2 tests. Each of these tests
' will be worth 12,5% (25% / 2). The progress flow will look like this:
' Object1:test1 = 25%, Object1:test2 = 50%, Object2:test1 = 62,5%
' Object2:test2 = 75%, Object1:test3 = 75%, Object1:test4 = 100%
' Notice that Object2:test2 and Object1:test3 has the same percentage value, that
' is since the logic demands that test3 is the progress for Object2.
'
Public Function fncCalculateTotalProgress() As Boolean
  Dim lCounter As Long, lProgress As Byte, sDivide As Single
  
  sDivide = 100
  
  For lCounter = 0 To lProgressCount - 1
    If Not (aProgressItem(lCounter).lProgressSpan = 0) Then _
      sDivide = (sDivide / aProgressItem(lCounter).lProgressSpan)
    lProgress = lProgress + (aProgressItem(lCounter).lProgress * sDivide)
  Next lCounter
  
  lTotalProgress = lProgress
End Function

' This function opens the VTM file and calls the fncParseVtmNodeList function
Public Function fncParseVtm() As Boolean
  fncParseVtm = False
  
  Dim objVtmDom As Object 'New object
  Dim objNodeList As Object ' object
  Dim objNode As Object ' object
   
  If bolVtmIsLoaded Then fncParseVtm = True: Exit Function
   
  Set objVtmDom = CreateObject("Msxml2.DOMDocument.4.0")
  
  objVtmDom.async = False
  objVtmDom.validateOnParse = False
  objVtmDom.preserveWhiteSpace = True
  objVtmDom.resolveExternals = False
  objVtmDom.setProperty "SelectionLanguage", "XPath"
  objVtmDom.setProperty "NewParser", True
  
  If Not objVtmDom.Load(sVtmPath & "vtm.xml") Then GoTo ErrorH
  bolVtmIsLoaded = True

  
  If objVtmDom Is Nothing Then GoTo ErrorH
  Set objNode = objVtmDom.selectSingleNode("//validatorTestMap")
  
  Set objNodeList = objNode.selectNodes("child::*[@context]")
  If Not fncParseVtmNodeList(objNodeList, "") Then GoTo ErrorH

  fncParseVtm = True
ErrorH:
  If Not fncParseVtm Then _
    objEvent.subLog "Error loading vtm (" & sVtmPath & "vtm.xml)"
  
  Set objNode = Nothing
  Set objNodeList = Nothing
  Set objVtmDom = Nothing
End Function

' This function takes the VTM DOM and retrieves information for all tests and puts
' them in the right context.                       ' object
Private Function fncParseVtmNodeList(iobjNodeList As Object, _
  isContext As String)
  
  fncParseVtmNodeList = False
  
  Dim objNodeList As Object
  Dim objNode As Object, objTestNode As Object
  Dim objContextAttr As Object, objTestContextAttr As Object
  Dim objAttrNode As Object, lCurrentItem As Long
  
  For Each objNode In iobjNodeList
    Set objContextAttr = objNode.selectSingleNode("@context")
    
    Set objNodeList = objNode.selectNodes("child::test")
  
    For Each objTestNode In objNodeList
      Set objAttrNode = objTestNode.selectSingleNode("@context")
            
      lCurrentItem = fncGetVtmItemIndex( _
        isContext & objContextAttr.nodeValue & "." & objAttrNode.nodeValue)
      
      If lCurrentItem = -1 Then
        ReDim Preserve aVtmItems(lVtmItemsCount)
        Set aVtmItems(lVtmItemsCount) = New oVtmItem
        lCurrentItem = lVtmItemsCount
        lVtmItemsCount = lVtmItemsCount + 1
      End If
    
      If Not objAttrNode Is Nothing Then aVtmItems(lCurrentItem).sContext = _
        isContext & objContextAttr.nodeValue & "." & objAttrNode.nodeValue
    
      Set objAttrNode = objTestNode.selectSingleNode("@failClass")
      If Not objAttrNode Is Nothing Then aVtmItems(lCurrentItem).sFailClass = objAttrNode.nodeValue
    
      Set objAttrNode = objTestNode.selectSingleNode("@failType")
      If Not objAttrNode Is Nothing Then aVtmItems(lCurrentItem).sFailType = objAttrNode.nodeValue
  
      Set objAttrNode = objTestNode.selectSingleNode("@link")
      If Not objAttrNode Is Nothing Then aVtmItems(lCurrentItem).sLink = objAttrNode.nodeValue
    
      Set objAttrNode = objTestNode.selectSingleNode("@longDesc")
      If Not objAttrNode Is Nothing Then aVtmItems(lCurrentItem).sLongDesc = objAttrNode.nodeValue
    
      Set objAttrNode = objTestNode.selectSingleNode("@name")
      If Not objAttrNode Is Nothing Then aVtmItems(lCurrentItem).sName = objAttrNode.nodeValue
    
      Set objAttrNode = objTestNode.selectSingleNode("@shortDesc")
      If Not objAttrNode Is Nothing Then aVtmItems(lCurrentItem).sShortDesc = objAttrNode.nodeValue

  
    Next objTestNode
    
    
    Set objNodeList = objNode.selectNodes("child::*[@context]")
    fncParseVtmNodeList objNodeList, isContext & objContextAttr.nodeValue & "."
  Next objNode

  Set objAttrNode = Nothing
  Set objTestContextAttr = Nothing
  Set objContextAttr = Nothing
  Set objNode = Nothing
  Set objTestNode = Nothing
  Set objNodeList = Nothing
  
  fncParseVtmNodeList = True
End Function

' This function returns the index of the VTM item that has the given context
Private Function fncGetVtmItemIndex(ByVal isContext As String) As Long
  isContext = LCase$(isContext)
  
  fncGetVtmItemIndex = -1
  
  Dim lTempCounter As Long
  For lTempCounter = 0 To lVtmItemsCount - 1
    If LCase$(aVtmItems(lTempCounter).sContext) = isContext Then
      fncGetVtmItemIndex = lTempCounter
      Exit Function
    End If
  Next lTempCounter
End Function

' This function returns the VTM item that has the given context
Public Function fncGetVtmItem(ByVal isContext As String) As oVtmItem
  isContext = LCase$(isContext)
  
  Dim lTempCounter As Long
  For lTempCounter = 0 To lVtmItemsCount - 1
    If LCase$(aVtmItems(lTempCounter).sContext) = isContext Then
      Set fncGetVtmItem = aVtmItems(lTempCounter)
      Exit Function
    End If
  Next lTempCounter
End Function

' This function inserts the given error and error information into the given
' report object. If a DOM Node is given, the function will try to calculate the
' line and column of the DOM node. **NOTICE!** since attribute nodes doesn't
' have a parent node, attribute nodes can NOT be used with this function.
'
Public Function fncInsFail2Report(iobjReport As oReport, iobjNode As Object, _
  isTestId As String, isAbsPath As String, Optional isComment As Variant)

  Dim lLine As Long, lColumn As Long

  Dim objNode As Object, objID As Object
  'Dim reader As New SAXXMLReader40
  Dim reader As Object
  Dim contentHandler As New SAXContentHandler
  
  On Error GoTo ErrorH
  
  Set reader = CreateObject("Msxml2.SAXXMLReader.4.0")
  
  If Not iobjNode Is Nothing Then
    ' search for the nearest ID attribute
    Set objNode = iobjNode
    Set objID = objNode.selectSingleNode("@id")
    Do Until Not objID Is Nothing
      Set objNode = objNode.parentNode
      If objNode Is Nothing Then Exit Do
      Set objID = objNode.selectSingleNode("@id")
    Loop

    Dim bolIDMatch As Boolean, bolAncestorMatch As Boolean, sAncestorChain As String

    ' If no ID attribute was found, we'll have to try to find the location using
    ' the nodes parents and it's attributes. This function will fail if any other
    ' node has the EXACT same parents and EXACT same attributes and attribute
    ' values.
    If objNode Is Nothing Then
      Set objNode = iobjNode
    
      sAncestorChain = objNode.nodeName
      Dim objNodeList As Object, lCounter As Long
      Set objNodeList = objNode.selectNodes("@*")
    
      If objNodeList.length > 0 Then
        sAncestorChain = sAncestorChain & "["
        For lCounter = 0 To objNodeList.length - 1
          sAncestorChain = sAncestorChain & "@" & _
            objNodeList.Item(lCounter).nodeName & "='" & _
            objNodeList.Item(lCounter).nodeValue & "'"
        
          If lCounter < objNodeList.length - 1 Then _
            sAncestorChain = sAncestorChain & " and "
        Next lCounter
        sAncestorChain = sAncestorChain & "]"
      End If
    
      Set objNode = objNode.parentNode
      Do Until (objNode Is Nothing)
        If (objNode.nodeName = "#document") Then Exit Do
        sAncestorChain = objNode.nodeName & "/" & sAncestorChain
        Set objNode = objNode.parentNode
      Loop
    
      Set objNodeList = iobjNode.ownerDocument.selectNodes(sAncestorChain)
      If objNodeList.length = 1 Then _
        bolAncestorMatch = True: contentHandler.sAncestorChain = sAncestorChain
  
      contentHandler.eNodeType = iobjNode.nodeType
      contentHandler.sText = iobjNode.nodeName
    Else
      bolIDMatch = True
      contentHandler.sId = objID.nodeValue
      contentHandler.eNodeType = objID.nodeType
      contentHandler.sText = "id"
    End If

    ' We use the SAX interface to get the line/column of the found node
    Set reader.contentHandler = contentHandler
    reader.parseURL isAbsPath

    lLine = contentHandler.lLine
    lColumn = contentHandler.lColumn
  Else
    lLine = -1
    lColumn = -1
  End If
  Dim sComment As String
  If Not IsMissing(isComment) Then sComment = isComment

ErrorH:
  iobjReport.fncInsertFailedTest isTestId, isAbsPath, lLine, lColumn, sComment
End Function

' This function looks in the given directory (isAbsPath) and returns the first
' file found that matches any of the filenames given (isFileName) and returns it
' (isOutputFileName)
'
Public Function fncGetPreferedFileName( _
  ByVal isAbsPath As String, ByRef isOutputFileName As String, _
  ParamArray isFileName() As Variant _
  ) As Boolean
  
  fncGetPreferedFileName = False
    
  Dim oFSO As Object, lCounter As Long
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  If oFSO Is Nothing Then objEvent.subLog "Error in fncGetNccName: " & _
    "couldn't create filesystemobject": Exit Function

  For lCounter = 0 To UBound(isFileName)
    If oFSO.FileExists(isAbsPath & isFileName(lCounter)) Then
      isOutputFileName = isFileName(lCounter)
      fncGetPreferedFileName = True
      Exit For
    End If
  Next lCounter
  
  Set oFSO = Nothing
End Function
