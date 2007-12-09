Attribute VB_Name = "mMain"
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



Option Explicit

' Job variables
Public aJobItems() As oJobItem
Public lJobCount As Long
Public lCurrentJob As Long
' Progress indicators
Public bolAbort As Boolean
Public bolRegenerating As Boolean
Public lProgress As Long
' Current log file object (NOTHING if not regenerating) and application path
Public objLogFile As Object
Public sAppPath As String
' File system object (created at startup since it's used very often)
Public oFSO As Object
' IANA Charsets
Public aIanaCharset() As String
Public lCharsetCount As Long

' Default job properties
Public lDtbType As Long
Public lCharset As Long
Public lIANACharset As Long
Public bolPreserveMeta As Boolean
Public bolSeqRename As Boolean
Public bolUseNumeric As Boolean
Public bolSameFolder As Boolean
Public sMetaFile As String
Public sSavePath As String
Public sPrefix As String

Public bolPb2kLayoutFix As Boolean
Public bolFixPar As Boolean
Public bolRebuildLinkStructure As Boolean
Public bolDisableBrokenXhtmlLinks As Boolean

Public bolMergeShortPhrases As Boolean
Public lClipSpan As Long
Public lClipLessThan As Long
Public lFirstClipLessThan As Long
Public lNextClipLessThan As Long
'new 20030917
Public bolEstimateBrokenXhtmlLinks As Boolean
Public bolDoVerboseLog As Boolean
Public bolAddCss As Boolean
Public bolMakeTrueNccOnly As Boolean
'Public bolPointTargetsToPar As Boolean
Public lSmilTarget As Long 'uses the SMILTARGET consts
'end new 20030917

Public bolMoveBook As Boolean

' General Settings
Public sDefaultSavePath As String
Public sDefaultMetaPath As String
Public bolHalt As Boolean

' Log Settings
Public sLogPath As String

' Constants for use with BrowseForFolder
Public Const BIF_RETURNONLYFSDIRS   As Long = &H1
Public Const BIF_DONTGOBELOWDOMAIN  As Long = &H2
Public Const BIF_VALIDATE           As Long = &H20
Public Const BIF_EDITBOX            As Long = &H10
Public Const BIF_NEWDIALOGSTYLE     As Long = &H40
Public Const BIF_USENEWUI As Long = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)

Public Const RENDER_REPLACE = 0
Public Const RENDER_NEWFOLDER = 1

Public Const SMILTARGET_NOCHANGE = 0
Public Const SMILTARGET_PAR = 1
Public Const SMILTARGET_TEXT = 2

' Interface object to be used (for future NON-UI interface use)
Public objUI As Object
Public bolValidatorExists As Boolean
Public objValidatorUserControl As Object
Public objRegeneratorUserControl As Object

Public objCurrentLog As oLogItem
Public objJobsLog As New oLogItem

'Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
'Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
'Public Const HIGH_PRIORITY_CLASS = &H80
'Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'Public Type MEMORYSTATUS
'        dwLength As Long
'        dwMemoryLoad As Long
'        dwTotalPhys As Long
'        dwAvailPhys As Long
'        dwTotalPageFile As Long
'        dwAvailPageFile As Long
'        dwTotalVirtual As Long
'        dwAvailVirtual As Long
'End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
   ByVal lpParameters As String, ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Public Sub Main()
  'On Error Resume Next
  On Error GoTo errH

' Set our Application path
  sAppPath = App.Path
  If Not Right$(sAppPath, 1) = "\" Then sAppPath = sAppPath & "\"
  
' Create a fileSystemObject, this one is going to be used many times in the program
  Set oFSO = CreateObject("scripting.fileSystemObject")
  
' Set the current log object
  Set objCurrentLog = objJobsLog

' check msxml4 availability
  Dim bMsxmlPresenceIsTested As Boolean
  Dim oTemp As Object
  bMsxmlPresenceIsTested = True
  Set oTemp = CreateObject("Msxml2.DOMDocument.4.0")
  If oTemp Is Nothing Then
    MsgBox "MSXML 4 not installed, exiting program.", vbOKOnly
    objUI.Unload
    Exit Sub
  End If
  bMsxmlPresenceIsTested = False

' Get the regenerator control object
  Set objRegeneratorUserControl = CreateObject("Regenerator.oUserControl")
  If objRegeneratorUserControl Is Nothing Then
    MsgBox "Regenerator dll not correctly registered!" & vbCrLf & "Exiting program.", vbOKOnly
    objUI.Unload
    Exit Sub
  End If
      
' Set jobitems to (0)
  ReDim aJobItems(0)
  
' Set s to be the main user interface
  Set objUI = frmMain
  
' Load IANA charset list
  fncLoadCharsetList
  objUI.subUpdateIanaCharactersets

' Load registry settings
  fncLoadRegistrySettings
  
  'fncAddMemLog vbCrLf & "Prg. start " & Date & " " & Time
  
  If bolUseValidator Then
    fncInitValidator
    'fncAddMemLog "Validator initialized."
  End If
  
  
  Exit Sub
errH:
  If bMsxmlPresenceIsTested Then
    MsgBox "MSXML 4 is not installed. Exiting program.", vbOKOnly
  Else
    MsgBox "An unknown error occured. Exiting program.", vbOKOnly
  End If
  objUI.Unload
  Exit Sub
End Sub

Public Function fncSetJobProperties()
Dim lJob As Long
  
' This function sets the job properties from
' the global variables found at the top of this file
    
  lJob = lCurrentJob
  If lJob = 0 Then Exit Function
  If lJob > lJobCount Then
    MsgBox "internal error - property setting cancelled"
    Exit Function
  End If
  
  With aJobItems(lJob)
    .lDtbType = lDtbType
    .lInputCharset = lCharset
    .lCharsetOther = lIANACharset
    .bolPreserveMeta = bolPreserveMeta
    .sMetaImport = sMetaFile
    .bolSeqRename = bolSeqRename
    .bolUseNumeric = bolUseNumeric
    .sPrefix = sPrefix
    .bolSaveSame = bolSameFolder
    .sNewFolder = sSavePath
    .bolMoveBook = bolMoveBook
  End With
  
End Function


' ***** Registry data settings ******

Private Function fncLoadRegistrySettings()
  Dim lValue As Long, sValue As String, bolValue As Boolean
  
  sBaseKey = "Software\DaisyWare\Regenerator\Last properties"
  
  'on error Resume Next
  
  fncLoadRegistryData "DTB Type", lDtbType, HKEY_CURRENT_USER, , 1
  fncLoadRegistryData "Input charset", lCharset, HKEY_CURRENT_USER, , 0
  fncLoadRegistryData "IANA charset", lIANACharset, HKEY_CURRENT_USER, , 0
  fncLoadRegistryData "Preserve meta", bolPreserveMeta, HKEY_CURRENT_USER, , True
  fncLoadRegistryData "Sequential rename", bolSeqRename, HKEY_CURRENT_USER, , True
  fncLoadRegistryData "Use numeric portion", bolUseNumeric, HKEY_CURRENT_USER, , False
  fncLoadRegistryData "Same folder", bolSameFolder, HKEY_CURRENT_USER, , False
  fncLoadRegistryData "Move book", bolMoveBook, HKEY_CURRENT_USER, , True
  
  sBaseKey = "Software\DaisyWare\Regenerator\General"
  
  fncLoadRegistryData "Default save path", sDefaultSavePath, HKEY_CURRENT_USER, , "*fulldtbpath*\regenerated\"
  fncLoadRegistryData "Default meta path", sDefaultMetaPath, HKEY_CURRENT_USER, , "*fulldtbpath*\"
  fncLoadRegistryData "Halt 'on error", bolHalt, HKEY_CURRENT_USER, , False
  fncLoadRegistryData "Fix par", bolFixPar, HKEY_CURRENT_USER, , True
  fncLoadRegistryData "Mangle links", bolRebuildLinkStructure, HKEY_CURRENT_USER, , True
  fncLoadRegistryData "Disable broken xhtml links", bolDisableBrokenXhtmlLinks, HKEY_CURRENT_USER, , False
  fncLoadRegistryData "Pb2k Layout fix", bolPb2kLayoutFix, HKEY_CURRENT_USER, , True
  fncLoadRegistryData "Merge phrases", bolMergeShortPhrases, HKEY_CURRENT_USER, , True
  fncLoadRegistryData "Merge short", lClipLessThan, HKEY_CURRENT_USER, , 400
  fncLoadRegistryData "Merge short 2", lFirstClipLessThan, HKEY_CURRENT_USER, , 1000
  fncLoadRegistryData "Merge next less than", lNextClipLessThan, HKEY_CURRENT_USER, , 5000
  fncLoadRegistryData "Merge maximum span", lClipSpan, HKEY_CURRENT_USER, , 100
  'new 20030917
  fncLoadRegistryData "Estimate broken xhtml links", bolEstimateBrokenXhtmlLinks, HKEY_CURRENT_USER, , False
  fncLoadRegistryData "Verbose log", bolDoVerboseLog, HKEY_CURRENT_USER, , True
  fncLoadRegistryData "Add default css", bolAddCss, HKEY_CURRENT_USER, , False
  fncLoadRegistryData "Make true NCC-Only", bolMakeTrueNccOnly, HKEY_CURRENT_USER, , False
  'fncLoadRegistryData "Redirect smil targets to par", bolPointTargetsToPar, , , False
  fncLoadRegistryData "smil targets", lSmilTarget, HKEY_CURRENT_USER, , SMILTARGET_NOCHANGE
  'end new 20030917
    
  sBaseKey = "Software\DaisyWare\Regenerator\Validation"
  
  fncLoadRegistryData "Validate job", bolUseValidator, HKEY_CURRENT_USER, , True
  fncLoadRegistryData "Include warnings", bolIncludeWarnings, HKEY_CURRENT_USER, , False
  fncLoadRegistryData "Include NC errors", bolIncludeNCErrors, HKEY_CURRENT_USER, , True
  fncLoadRegistryData "Include advanced ADTD information", bolIncludeADVADTD, HKEY_CURRENT_USER, , _
    False
  fncLoadRegistryData "Create standalone report", bolCreateStandalone, HKEY_CURRENT_USER, , False
  fncLoadRegistryData "Save standalone report", sStandalonePath, HKEY_CURRENT_USER, , _
    "*fulldtbpath*\regenerated\report.html"
  
  fncLoadRegistryData "Synchronize settings with validator", bolSyncWValidator, HKEY_CURRENT_USER, _
    , True
  
  fncLoadRegistryData "Validator Light Mode", bolValidatorLightMode, HKEY_CURRENT_USER, _
    , False
  
  Dim sValDTDPath As String, sValAppPath As String, sValTempPath As String
  Dim lValTimeFluct As Long, sValVTMPath As String
  
' If there are no registry settings try to load the Daisy Validators registry keys

  sBaseKey = "Software\DaisyWare\Validator\Misc"
  fncLoadRegistryData "Dtd_AdtdPath", sValDTDPath, HKEY_CURRENT_USER, , ""
  fncLoadRegistryData "AppPath", sValVTMPath, HKEY_CURRENT_USER, , ""
  fncLoadRegistryData "TempPath", sValTempPath, HKEY_CURRENT_USER, , sValue
  sBaseKey = "Software\DaisyWare\Validator\Settings"
  fncLoadRegistryData "TimeFluctuation", lValTimeFluct, HKEY_CURRENT_USER, , 0
  If bolSyncWValidator Then
    sExtPath = sValDTDPath
    sVTMPath = sValVTMPath
    sTempPath = sValTempPath
    lTimeFluctuation = lValTimeFluct
  Else
    sBaseKey = "Software\DaisyWare\Regenerator\Validation"
    fncLoadRegistryData "Externals path", sExtPath, HKEY_CURRENT_USER, , sValDTDPath
    fncLoadRegistryData "VTM path", sVTMPath, HKEY_CURRENT_USER, , sValVTMPath
    fncLoadRegistryData "Temp path", sTempPath, HKEY_CURRENT_USER, , sValTempPath
    fncLoadRegistryData "Time fluctuation", lTimeFluctuation, HKEY_CURRENT_USER, , lValTimeFluct
  End If
  
  'if the validator is not installed or something is wrong
  If (sValDTDPath = "" Or sValVTMPath = "" Or sExtPath = "") Then _
    bolUseValidator = False

  sBaseKey = "Software\DaisyWare\Regenerator\Log"
  fncLoadRegistryData "Save log", sLogPath, HKEY_CURRENT_USER, , sAppPath & "log\"
  
  objRegeneratorUserControl.fncSetAudioTemplatesPath sAppPath & "audiotemplates\"
  objRegeneratorUserControl.fncSetResourcePath sAppPath & "resources\"
  objRegeneratorUserControl.fncSetTidyLibPath sAppPath & "tidylib\"
  
End Function



Public Function fncSaveRegistrySettings()
  sBaseKey = "Software\DaisyWare\Regenerator\Last properties"
  
  fncSaveRegistryData "DTB Type", lDtbType, HKEY_CURRENT_USER
  fncSaveRegistryData "Input charset", lCharset, HKEY_CURRENT_USER
  fncSaveRegistryData "IANA charset", lIANACharset, HKEY_CURRENT_USER
  fncSaveRegistryData "Preserve meta", bolPreserveMeta, HKEY_CURRENT_USER
  fncSaveRegistryData "Sequential rename", bolSeqRename, HKEY_CURRENT_USER
  fncSaveRegistryData "Use numeric portion", bolUseNumeric, HKEY_CURRENT_USER
  fncSaveRegistryData "Same folder", bolSameFolder, HKEY_CURRENT_USER
  fncSaveRegistryData "Move book", bolMoveBook, HKEY_CURRENT_USER

  sBaseKey = "Software\DaisyWare\Regenerator\General"
  
  fncSaveRegistryData "Default save path", sDefaultSavePath, HKEY_CURRENT_USER
  fncSaveRegistryData "Default meta path", sDefaultMetaPath, HKEY_CURRENT_USER
  fncSaveRegistryData "Halt 'on error", bolHalt, HKEY_CURRENT_USER
  fncSaveRegistryData "Fix par", bolFixPar, HKEY_CURRENT_USER
  fncSaveRegistryData "Pb2k Layout fix", bolPb2kLayoutFix, HKEY_CURRENT_USER
  fncSaveRegistryData "Mangle links", bolRebuildLinkStructure, HKEY_CURRENT_USER
  fncSaveRegistryData "Disable broken xhtml links", bolDisableBrokenXhtmlLinks, HKEY_CURRENT_USER
  fncSaveRegistryData "Merge phrases", bolMergeShortPhrases, HKEY_CURRENT_USER
  fncSaveRegistryData "Merge short", lClipLessThan, HKEY_CURRENT_USER
  fncSaveRegistryData "Merge short 2", lFirstClipLessThan, HKEY_CURRENT_USER
  fncSaveRegistryData "Merge next less than", lNextClipLessThan, HKEY_CURRENT_USER
  fncSaveRegistryData "Merge maximum span", lClipSpan, HKEY_CURRENT_USER
  
  'new 20030917
  fncSaveRegistryData "Estimate broken xhtml links", bolEstimateBrokenXhtmlLinks, HKEY_CURRENT_USER
  fncSaveRegistryData "Verbose log", bolDoVerboseLog, HKEY_CURRENT_USER
  fncSaveRegistryData "Add default css", bolAddCss, HKEY_CURRENT_USER
  fncSaveRegistryData "Make true NCC-Only", bolMakeTrueNccOnly, HKEY_CURRENT_USER
  'fncSaveRegistryData "Redirect smil targets to par", bolPointTargetsToPar
  fncSaveRegistryData "smil targets", lSmilTarget, HKEY_CURRENT_USER
  'end new 20030917

  sBaseKey = "Software\DaisyWare\Regenerator\Validation"

  fncSaveRegistryData "Validate job", bolUseValidator, HKEY_CURRENT_USER
  fncSaveRegistryData "Include warnings", bolIncludeWarnings, HKEY_CURRENT_USER
  fncSaveRegistryData "Include NC errors", bolIncludeNCErrors, HKEY_CURRENT_USER
  fncSaveRegistryData "Include advanced ADTD information", bolIncludeADVADTD, HKEY_CURRENT_USER
  fncSaveRegistryData "Create standalone report", bolCreateStandalone, HKEY_CURRENT_USER
  fncSaveRegistryData "Save standalone report", sStandalonePath, HKEY_CURRENT_USER
  
  fncSaveRegistryData "Synchronize settings with validator", bolSyncWValidator, HKEY_CURRENT_USER
  fncSaveRegistryData "Validator Light Mode", bolValidatorLightMode, HKEY_CURRENT_USER
  
  fncSaveRegistryData "Externals path", sExtPath, HKEY_CURRENT_USER
  fncSaveRegistryData "VTM path", sVTMPath, HKEY_CURRENT_USER
  fncSaveRegistryData "Temp path", sTempPath, HKEY_CURRENT_USER
  fncSaveRegistryData "Time fluctuation", lTimeFluctuation, HKEY_CURRENT_USER

  sBaseKey = "Software\DaisyWare\Regenerator\Log"
  fncSaveRegistryData "Save log", sLogPath, HKEY_CURRENT_USER
End Function

Public Function fncInsertJob(isFileName As String)
' This function adds a job to the job queue
  
  If Not fncValidatePath(isFileName) Then
    MsgBox ("Cannot resolve path " & isFileName & "!")
    Exit Function
  End If
  
  lJobCount = lJobCount + 1
  ReDim Preserve aJobItems(lJobCount)
  Set aJobItems(lJobCount) = New oJobItem
  
  Dim sPath As String
  
  With aJobItems(lJobCount)
    .sPath = isFileName
    .sID = "unknown"
    .bolRegResult = False
    .bolRendered = False
    .bolRegRun = False
    .bolValResult = False
    Set .objLog = New oLogItem
  End With
  
  sSavePath = sDefaultSavePath
  sMetaFile = sDefaultMetaPath
  lCurrentJob = lJobCount
  fncSetJobProperties
  'fncAddMemLog "Added job (" & lCurrentJob & ")"
End Function

' This function adds a XML joblist file with several jobs in
Public Function fncAddJobList(isJobListFile As String) As Boolean
  Dim objDom As Object, objJobNode As Object
  Dim objNodeList As Object
  Dim objNode As Object
  
  On Error Resume Next
  
  Set objDom = CreateObject("Msxml2.DOMDocument.4.0")
  If objDom Is Nothing Then
    objUI.fncAddLog "MSXML not installed or correctly registered, aborting.", False
    Exit Function
  End If
  
  objDom.validateOnParse = False
  objDom.setProperty "SelectionLanguage", "XPath"
  If Not objDom.Load(isJobListFile) Then
    objUI.fncAddLog "<error from='ui' in='fncAddJobList'>Error in adding joblist: [" & objDom.parseError.Line & ":" & _
      objDom.parseError.linepos & "] " & objDom.parseError.errorCode & " " & _
      objDom.parseError.reason & "</error>", False
    Exit Function
  End If
  
  'on error Resume Next
  'select all <job> elements

  Set objJobNode = objDom.selectSingleNode("//default")
  If Not objJobNode Is Nothing Then
    Dim objDefaultSettings As New oJobItem
    fncGetJobFromFile objJobNode, objDefaultSettings
    
    With objDefaultSettings
      bolMoveBook = .bolMoveBook
      bolPreserveMeta = .bolPreserveMeta
      bolSeqRename = .bolSeqRename
      bolUseNumeric = .bolUseNumeric
      lIANACharset = .lCharsetOther
      lDtbType = .lDtbType
      lCharset = .lInputCharset
      sMetaFile = .sMetaImport
      sSavePath = .sNewFolder
      sPrefix = .sPrefix
    End With
  End If

  Set objNodeList = objDom.selectNodes("//job")
  For Each objJobNode In objNodeList
    fncGetJobFromFile objJobNode
  Next objJobNode
  
  objUI.fncUpdateJobList
  objUI.subRefreshInterface
End Function

Private Function fncGetJobFromFile(iobjJobNode As Object, _
  Optional objDefault As Variant) As Boolean
    
  Dim sNccFile As String
  
  Dim sMF As String, bMI As Boolean, sOPP As String, lDT As Long, lCS As Long, lSCS As Long
  Dim bolSR As Boolean, bolUN As Boolean, sP As String, bolMB As Boolean
  
  Dim objNode As Object, objCurrentJob As oJobItem
  
  Set objNode = iobjJobNode.selectSingleNode("nccfile")
  If (objNode Is Nothing) And (IsMissing(objDefault)) Then
    Exit Function
  ElseIf IsMissing(objDefault) Then
    sNccFile = objNode.Text
    fncInsertJob sNccFile
    Set objCurrentJob = aJobItems(lJobCount)
  Else
    Set objCurrentJob = objDefault
  End If
  
  Set objNode = iobjJobNode.selectSingleNode("metafile")
  'mg20030325: made metaimport an option
  If objNode Is Nothing Then
    sMF = sMetaFile
    bMI = False
  Else
    sMF = objNode.Text
    If objNode.Text = "" Then
      bMI = False
    Else
      bMI = True
    End If
  End If
  Set objNode = iobjJobNode.selectSingleNode("outputpath")
  If objNode Is Nothing Then sOPP = sSavePath Else sOPP = objNode.Text
  Set objNode = iobjJobNode.selectSingleNode("dtbtype")
  If objNode Is Nothing Then lDT = lDtbType Else lDT = CLng(objNode.Text)
  Set objNode = iobjJobNode.selectSingleNode("charset")
  If objNode Is Nothing Then lCS = lCharset Else lCS = CLng(objNode.Text)
  Set objNode = iobjJobNode.selectSingleNode("subcharset")
  If objNode Is Nothing Then lSCS = lIANACharset Else lSCS = CLng(objNode.Text)
  Set objNode = iobjJobNode.selectSingleNode("prefix")
  If objNode Is Nothing Then sP = sPrefix Else sP = objNode.Text
  Set objNode = iobjJobNode.selectSingleNode("seqrename")
  If objNode Is Nothing Then bolSR = bolSeqRename Else bolSR = CBool(objNode.Text)
  Set objNode = iobjJobNode.selectSingleNode("usenumeric")
  If objNode Is Nothing Then bolUN = bolUseNumeric Else bolUN = CBool(objNode.Text)
  Set objNode = iobjJobNode.selectSingleNode("movebook")
  If objNode Is Nothing Then bolMB = bolMoveBook Else bolMB = CBool(objNode.Text)
  
  With objCurrentJob
    .bolPreserveMeta = Not bMI
    .sMetaImport = sMF
    .sNewFolder = sOPP
    .lDtbType = lDT
    .lInputCharset = lCS
    .lCharsetOther = lSCS
    .sPrefix = sP
    .bolSeqRename = bolSR
    .bolUseNumeric = bolUN
    .bolMoveBook = bolMB
  End With
  
  Set objCurrentJob = Nothing
  fncGetJobFromFile = True
End Function

Public Function fncRunJobQueue()
Dim objRegenerator As Object

  'Show the UI that we're working so it can disable parts of the interface etc.
  bolRegenerating = True
  
  For lCurrentJob = 1 To lJobCount
    fncSetLogFile
    objUI.subRefreshInterface
    objUI.fncUpdateJobList
    With aJobItems(lCurrentJob)
      
      ' Create a new instance of the oRegenerator object
      Set objRegenerator = CreateObject("Regenerator.oregenerator")
      If objRegenerator Is Nothing Then
        objUI.fncAddLog "<error from='ui' in='fncRunJobQueue'>Regenerator dll not registered, aborting.</error>", True
        Exit Function
      End If
      
      ' Set the current UI to respond to events from the new oRegenerator object
      Set objRegenerator.objowner = objUI
      
      objUI.fncAddLog "<job>", True
      
      If fncInitJob(aJobItems(lCurrentJob), lCurrentJob) Then
        'debug section
         'objUI.fncAddLog ".sPath: " & .sPath, False
         'objUI.fncAddLog ".lInputCharset: " & .lInputCharset, False
         'objUI.fncAddLog "frmMain.objComboIanaCS.Text: " & frmMain.objComboIanaCS.Text, False
         'objUI.fncAddLog ".lDtbType: " & .lDtbType, False
         'objUI.fncAddLog "False: ", False
         'objUI.fncAddLog ".bolPreserveMeta: " & .bolPreserveMeta, False
         'objUI.fncAddLog ".sMetaImport: " & .sMetaImport, False
         'objUI.fncAddLog ".bolSeqRename: " & .bolSeqRename, False
         'objUI.fncAddLog ".bolUseNumeric: " & .bolUseNumeric, False
         'objUI.fncAddLog ".sPrefix: " & .sPrefix, False
         'objUI.fncAddLog "bolRebuildLinkStructure: " & bolRebuildLinkStructure, False
         'objUI.fncAddLog "bolFixPar: " & bolFixPar, False
         'objUI.fncAddLog "bolMergeShortPhrases: " & bolMergeShortPhrases, False
         'objUI.fncAddLog "lClipSpan: " & lClipSpan, False
         'objUI.fncAddLog "lClipLessThan: " & lClipLessThan, False
         'objUI.fncAddLog "lFirstClipLessThan: " & lFirstClipLessThan, False
         'objUI.fncAddLog "lNextClipLessThan: " & lNextClipLessThan, False
         'objUI.fncAddLog "bolDisableBrokenXhtmlLinks: " & bolDisableBrokenXhtmlLinks, False
         'objUI.fncAddLog "bolEstimateBrokenXhtmlLinks: " & bolEstimateBrokenXhtmlLinks, False
         'objUI.fncAddLog "bolPb2kLayoutFix: " & bolPb2kLayoutFix, False
         'objUI.fncAddLog "bolDoVerboseLog: " & bolDoVerboseLog, False
         'objUI.fncAddLog "bolAddCss: " & bolAddCss, False
         'objUI.fncAddLog "bolMakeTrueNccOnly: " & bolMakeTrueNccOnly, False
         'objUI.fncAddLog "lSmilTarget: " & lSmilTarget, False
         'Stop
        'end debug section
      
        If Not objRegenerator.fncRunTransform(.sPath, .lInputCharset, frmMain.objComboIanaCS.Text, .lDtbType, _
          False, .bolPreserveMeta, .sMetaImport, .bolSeqRename, .bolUseNumeric, .sPrefix, _
          bolRebuildLinkStructure, bolFixPar, bolMergeShortPhrases, lClipSpan, lClipLessThan, _
          lFirstClipLessThan, lNextClipLessThan, bolDisableBrokenXhtmlLinks, bolEstimateBrokenXhtmlLinks, bolPb2kLayoutFix, _
          bolDoVerboseLog, bolAddCss, bolMakeTrueNccOnly, lSmilTarget _
          ) Then
          .bolRegResult = False
          .bolRegRun = True
          If bolAbort Then Exit For
          objUI.fncUpdateJobList
        Else
          .bolRegResult = True
          .bolRegRun = True
          If bolAbort Then Exit For
          objUI.fncUpdateJobList
          
          'regeneration didnt abort so proceed with rendering
          'set the .sID property using a prop in oRegen, only avail after fncRunTransform
          Dim sDTBID As String
          If fncConvertToUri(objRegenerator.fncGetDcIdentifier(), sDTBID) Then
            If sDTBID = "" Then sDTBID = "unknown"
            aJobItems(lCurrentJob).sID = sDTBID
          End If
          
          'set .sNewFolder
          Dim sTemp As String
           If Not .bolSaveSame Then
             fncParsePathWithConstants .sNewFolder, sTemp, lCurrentJob
             If Not fncValidatePath(sTemp) Then
               objUI.fncAddLog "<error from='ui' in='subRunJobQueue'>Cannot resolve path " & sTemp & "</error>", True
             Exit Function
             Else
              .sNewFolder = sTemp
             End If
          End If ' not .bolSaveSame
             
          Dim lTemp As Long
          If .bolSaveSame Then lTemp = RENDER_REPLACE Else lTemp = RENDER_NEWFOLDER
          'objUI.fncAddLog "rendering files...", False
          If Not objRegenerator.fncRenderFiles(lTemp, .sNewFolder) Then
            .bolRendered = False
            objUI.fncAddLog "rendering aborted.", False
            If Not objRegenerator.fncterminateobject Then
              objUI.fncAddLog "Couldn't terminate object (objRegenerator.fncterminateobject)", False
            End If
            Set objRegenerator = Nothing
          Else
            .bolRendered = True
            'objUI.fncAddLog "rendering done.", False
            If Not objRegenerator.fncterminateobject Then
              objUI.fncAddLog "Couldn't terminate object (objRegenerator.fncterminateobject)", False
            End If
            Set objRegenerator = Nothing
                        
            If .bolSaveSame Then
              .sRenderedTo = fncGetPathName(.sPath)
            Else
              .sRenderedTo = .sNewFolder
            End If
            If Not Right$(.sRenderedTo, 1) = "\" Then .sRenderedTo = .sRenderedTo & "\"
            
            If bolAbort Then Exit For 'goto cleanup in orig
        
            'rendering didnt abort so proceed with validation
            If .bolRegResult And .bolRendered And bolUseValidator And bolValidatorExists Then
              objUI.fncAddLog "<validating>", True
              If fncInitValJob Then
                objUI.fncAddLog "Validating...", False
                DoEvents
                If fncValidate(.sRenderedTo) Then
                  'objUI.fncAddLog "Validation succeeded."
                Else
                  objUI.fncAddLog "<processFailure in='validation'>Validation process failed</processFailure>", True
                End If
              End If 'fncInitValJob
              objUI.fncAddLog "</validating>", True
            End If '.bolRegResult And
            
            
            
            'regardless of validation, proceed with move book
            If .bolMoveBook Then
              If aJobItems(lCurrentJob).sID = "" Then
                fncGetDtbId aJobItems(lCurrentJob)
              End If
              fncMoveBook
            End If
                               
          End If 'Not objRegenerator.fncRenderFiles
        End If 'Not objRegenerator.fncRunTransform
      End If 'fncinitjob
      
      objUI.fncAddLog "</job>", True
    
      If bolAbort Then Exit For
      If (.bolRegResult = False And bolHalt = True) Then Exit For
    End With
  Next lCurrentJob
  
  bolRegenerating = False
  bolAbort = False
  Set objCurrentLog = objJobsLog
  If Not objLogFile Is Nothing Then objLogFile.Close
  Set objLogFile = Nothing
  lCurrentJob = 1
  objUI.fncUpdateJobList
  objUI.subRefreshInterface
  
End Function

Private Function fncInitValJob() As Boolean
 Dim sTemp As String
  
  fncInitValJob = False

  objUI.fncAddLog "<dllProperties>", True
  objUI.fncAddLog "<propLightMode value='" & CStr(objValidatorUserControl.propLightMode()) & "'/>", True
  objUI.fncAddLog "<propTimeSpan value='" & CStr(objValidatorUserControl.propTimeSpan()) & "'/>", True
  objUI.fncAddLog "<propTempPath value='" & objValidatorUserControl.propTempPath() & "'/>", True
  objUI.fncAddLog "<propDTDPath value='" & objValidatorUserControl.propDTDPath() & "'/>", True
  objUI.fncAddLog "<propAdtdPath value='" & objValidatorUserControl.propAdtdPath() & "'/>", True
  objUI.fncAddLog "<propVtmPath value='" & objValidatorUserControl.propVtmPath() & "'/>", True
  objUI.fncAddLog "</dllProperties>", True
  
  If objValidatorUserControl.propLightMode() Then objUI.fncAddLog "Note: light validation mode enabled", False

  If bolCreateStandalone Then
    fncParsePathWithConstants sStandalonePath, sTemp, lCurrentJob
    If Not fncValidatePath(sTemp) Then
      objUI.fncAddLog "<error from='ui' in='subRunJobQueue'>Cannot resolve path " & sTemp & "</error>", True
      Exit Function
    Else
      ''?
    End If
  End If

  fncInitValJob = True
End Function

Private Function fncInitJob(ByRef JobItem As oJobItem, lCurrentJob As Long) As Boolean
Dim sTemp As String, sTemp2 As String
  
  fncInitJob = False
  With JobItem
    ' Set the current jobs log object to collect data
    Set objCurrentLog = .objLog
    ' default those JobItem values that will be set by runjobqueue func
    .bolRegResult = False
    .bolRendered = False
    .bolRegRun = False
    .bolValResult = False
    .sErrorClass = ""
    .sErrorType = ""
    .sRenderedTo = ""

    ' Parse the paths given; they may contain path variables to be absoluted,
    ' ignore any validation result variables,
    ' since the data hasn't been validated yet.
    ' Also validate all paths before starting the regeneration
        
' doesnt .sPath have to be absolute??
'      fncParsePathWithConstants .sPath, sTemp, lCurrentJob
'      If Not fncValidatePath(sTemp) Then
'        objUI.fncAddLog "<error from='ui' in='subRunJobQueue'>Cannot resolve path " & sTemp & "</error>", True
'        'Stop
'         GoTo Skip2
'      End If
        
                           
    If Not .bolPreserveMeta Then
      fncParsePathWithConstants .sMetaImport, sTemp2, lCurrentJob
      ' If .sMetaImport is relative we'll relate it to the dtbpath
      sTemp2 = fncStripIdAddPath(sTemp2, .sPath)
      If Not fncValidatePath(sTemp2) Then
        objUI.fncAddLog "<error from='ui' in='subRunJobQueue'>Cannot resolve path " & sTemp2 & "</error>", True
      Else
        .sMetaImport = sTemp2
      End If
    End If
    
  End With
  fncInitJob = True

End Function


' ***** Misc functions *****

' This function loads the charsets.xml file and puts the charsets and their aliases
' in the aIanaCharset array
Private Function fncLoadCharsetList()
  Dim objDom As Object, objNodeList As Object
  Dim objNode As Object, lCounter As Long

  On Error Resume Next

  Set objDom = CreateObject("Msxml2.DOMDocument.4.0")
  If objDom Is Nothing Then
    objUI.fncAddLog "MSXML not installed or correctly registered, aborting.", False
    Exit Function
  End If
  
  objDom.validateOnParse = False
  objDom.setProperty "SelectionLanguage", "XPath"
  
  objDom.Load sAppPath & "resources\charsets.xml"
  
  Set objNodeList = objDom.selectNodes("//name | //alias")
  For Each objNode In objNodeList
    ReDim Preserve aIanaCharset(lCharsetCount)
    aIanaCharset(lCharsetCount) = objNode.Text
    lCharsetCount = lCharsetCount + 1
  Next
End Function

Public Function fncSetLogFile()
Dim sLogFile As String, lCounter As Long
  
' This function creates or opens a logfile that is going to be used for logging the
' current session
      
  If Not fncParsePathWithConstants(sLogPath, sLogFile, lCurrentJob) Then
    sLogFile = sAppPath
  End If
  
  fncCreateDirectoryChain (sLogFile)
  
  sLogFile = sLogFile & "regenerator.log"
  
  If Not objLogFile Is Nothing Then
    objLogFile.Close
    Set objLogFile = Nothing
  End If
  
  Set objLogFile = oFSO.opentextfile(sLogFile, 8, True)
  
End Function

' This function moves the rendered book to it's final destination
Public Function fncMoveBook()
  Dim lCounter As Long, sSrc As String, sDest As String
  Dim oFolder As Object, sTemp As String
  
  objUI.fncAddLog "Moving data...", False
  
' This function gets the Destination path for the book
  fncParsePathWithConstants aJobItems(lCurrentJob).sNewFolder, sDest, lCurrentJob
  If Not Right$(sDest, 1) = "\" Then sDest = sDest & "\"
  
' If path is relative we'll relate it to the dtbpath
  sDest = fncStripIdAddPath(sDest, aJobItems(lCurrentJob).sPath)

' This function gets the Source path for the book (the source path is always the
' same path as the destination with the exception that the validation variables are
' ignored, since the validation is being done on the rendered data)
  
  sSrc = aJobItems(lCurrentJob).sRenderedTo

  lCounter = InStrRev(sDest, "\", Len(sDest) - 1, vbBinaryCompare)
  sTemp = Mid$(sDest, lCounter + 1, Len(sDest) - lCounter - 1)

' Create the whole chain of directories needed
  sDest = Left$(sDest, lCounter)
  fncCreateDirectoryChain sDest

' Rename the folder to the last directory name in the Destination path
' Example: 'd:\dtbs\error_critical\book1\' = oFolder.Name = 'book1'
  Set oFolder = oFSO.GetFolder(sSrc)
  
  If Right$(sSrc, 1) = "\" Then sSrc = Left$(sSrc, Len(sSrc) - 1)
  
' If the destination folder already exists, add a number with format "0000" to the
' end of it
  Dim sBackup As String
  sBackup = sTemp: lCounter = 0
  Do Until Not (oFSO.folderexists(sDest + sTemp + "\") Or _
    oFSO.fileexists(sDest + sTemp))
    lCounter = Int(Rnd * 1000)
    sTemp = sBackup & "_" & Format(lCounter, "0000")
  Loop
  If Not sBackup = sTemp Then sSrc = sSrc & "_" & Format(lCounter, "0000")

' Change the name of the source folder
  If Not oFolder.Name = sTemp Then oFolder.Name = sTemp

' Set oFolder = Nothing

' Move the source folder to the destination folder
' example: moveFolder 'd:\dtbs\book1\', 'd:\dtbs\error_critical\'
' result: 'd:\dtbs\error_critical\book1\'
  
  'added 20030223:
  fncCreateDirectoryChain (sDest)
  Dim sSource As String
  'Stop
  sSource = oFolder.Path
  Set oFolder = Nothing
  'oFSO.movefolder sSource, sDest
  
  oFSO.CopyFolder sSource, sDest
  oFSO.DeleteFolder sSource
        
  objUI.fncAddLog "Moving book done.", False
  
' Update the sRenderedTo string
End Function

'Public Function fncLaunchProgram()
'Dim sPath As String, dTaskID As Double, lResult As Long
'
' This function launches a custom program with the given flags after each successfully
' regenerated job
' **** NOT IMPLEMENTED
'
'  fncParsePathWithConstants sLaunchProgram, sPath, lCurrentJob
'  dTaskID = Shell(sPath, vbNormalFocus)
'
'End Function



' This function replaces path variables within expressions

Public Function fncParsePathWithConstants( _
    ByVal isPath As String, _
    ByRef isOutput As String, _
    ByVal lDTBIndex As Long _
    ) As Boolean
  
Dim sFullDTBPath As String
Dim sDTBDirectory As String
Dim sTemp As String
Dim lTemp As Long
Dim sValResult As String
Dim oFolder As Object
Dim sDTBID As String

  On Error GoTo ErrorH
  
  sTemp = aJobItems(lDTBIndex).sPath
  
  isPath = LCase$(isPath)
  
  sFullDTBPath = oFSO.getparentfoldername(sTemp)
  If Right$(sFullDTBPath, 1) = "\" Then sFullDTBPath = Left$(sFullDTBPath, Len(sFullDTBPath) - 1)
  
  Set oFolder = oFSO.GetFolder(sFullDTBPath)
  If oFolder Is Nothing Then Exit Function
  sDTBDirectory = oFolder.Name
  
  With aJobItems(lDTBIndex)
    If .bolValResult = True Then
      sValResult = "pass"
    ElseIf .sErrorType = "" Then
      sValResult = "not_tested"
    Else
      sValResult = .sErrorType
      If Not .sErrorClass = "" Then sValResult = sValResult & "_" & .sErrorClass
    End If
    If sValResult = "" Then sValResult = "Not_tested"
  End With

' Derive DTBID flag

  isPath = Replace(isPath, "*fulldtbpath*", sFullDTBPath)
  isPath = Replace(isPath, "*dtbdirectory*", sDTBDirectory)
  isPath = Replace(isPath, "*validationresult*", sValResult)
  isPath = Replace(isPath, "*dtbid*", aJobItems(lCurrentJob).sID)
  
  'mg20030325
  isPath = Replace(isPath, "*destinationdirectory*", aJobItems(lCurrentJob).sRenderedTo)
  
  isOutput = isPath
  
  fncParsePathWithConstants = True
ErrorH:
  Set oFolder = Nothing
End Function

' This function gets the ID (dc:identifier or ncc:identifier) from the DTB

Public Function fncGetDtbId(ByRef JobItem As oJobItem)
Dim objDom As Object, objDomNode As Object, sDTBID As String

  On Error GoTo errH
  Set objDom = CreateObject("Msxml2.DOMDocument.4.0")
  If objDom Is Nothing Then
    objUI.fncAddLog "MSXML not installed or correctly registered, aborting.", False
    Exit Function
  End If
  objDom.validateOnParse = False
  objDom.resolveExternals = False
  objDom.setProperty "SelectionLanguage", "XPath"
  objDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
  objDom.setProperty "NewParser", True

  ' Load the file
  If objDom.Load(aJobItems(lCurrentJob).sRenderedTo & "ncc.html") Then
    ' get the meta name="dc:identifier"
    Set objDomNode = objDom.selectSingleNode("//xht:meta[@name = 'dc:identifier']/@content")
    If objDomNode Is Nothing Then
      ' get the meta name="ncc:identifier" if "dc:identifier" doesn't exist
      Set objDomNode = objDom.selectSingleNode("//xht:meta[@name='ncc:identifier']/@content")
      If objDomNode Is Nothing Then
        aJobItems(lCurrentJob).sID = "unknown"
        Exit Function
      End If
    End If
    
    ' Get the node value
    If objDomNode.nodeValue = "" Then
      aJobItems(lCurrentJob).sID = "unknown"
      Exit Function
    Else
      ' convert the value to valid URI characters and set the value in the job properties
      If fncConvertToUri(Trim$(objDomNode.nodeValue), sDTBID) Then
        aJobItems(lCurrentJob).sID = sDTBID
      End If
    End If 'objDomNode.nodeValue = ""
  Else
    aJobItems(lCurrentJob).sID = "unknown"
  End If 'objDom.load
    
  Exit Function
errH:
  objUI.fncAddLog "Error in fncGetDTBID [" & Err.Number & ", " & Err.Description & "]", False
End Function

Private Function fncLoadRegistrySettingsOld()
  Dim lValue As Long, sValue As String, bolValue As Boolean
  
  sBaseKey = "Software\DaisyWare\Regenerator\Last properties"
  
  'on error Resume Next
  
  fncLoadRegistryData "DTB Type", lDtbType, , , 1
  fncLoadRegistryData "Input charset", lCharset, , , 0
  fncLoadRegistryData "IANA charset", lIANACharset, , , 0
  fncLoadRegistryData "Preserve meta", bolPreserveMeta, , , True
  fncLoadRegistryData "Sequential rename", bolSeqRename, , , True
  fncLoadRegistryData "Use numeric portion", bolUseNumeric, , , False
  fncLoadRegistryData "Same folder", bolSameFolder, , , False
  fncLoadRegistryData "Move book", bolMoveBook, , , True
  
  sBaseKey = "Software\DaisyWare\Regenerator\General"
  
  fncLoadRegistryData "Default save path", sDefaultSavePath, , , "*fulldtbpath*\regenerated\"
  fncLoadRegistryData "Default meta path", sDefaultMetaPath, , , "*fulldtbpath*\"
  fncLoadRegistryData "Halt 'on error", bolHalt, , , False
  fncLoadRegistryData "Fix par", bolFixPar, , , True
  fncLoadRegistryData "Mangle links", bolRebuildLinkStructure, , , True
  fncLoadRegistryData "Disable broken xhtml links", bolDisableBrokenXhtmlLinks, , , False
  fncLoadRegistryData "Pb2k Layout fix", bolPb2kLayoutFix, , , True
  fncLoadRegistryData "Merge phrases", bolMergeShortPhrases, , , True
  fncLoadRegistryData "Merge short", lClipLessThan, , , 400
  fncLoadRegistryData "Merge short 2", lFirstClipLessThan, , , 1000
  fncLoadRegistryData "Merge next less than", lNextClipLessThan, , , 5000
  fncLoadRegistryData "Merge maximum span", lClipSpan, , , 100
  'new 20030917
  fncLoadRegistryData "Estimate broken xhtml links", bolEstimateBrokenXhtmlLinks, , , False
  fncLoadRegistryData "Verbose log", bolDoVerboseLog, , , True
  fncLoadRegistryData "Add default css", bolAddCss, , , False
  fncLoadRegistryData "Make true NCC-Only", bolMakeTrueNccOnly, , , False
  'fncLoadRegistryData "Redirect smil targets to par", bolPointTargetsToPar, , , False
  fncLoadRegistryData "smil targets", lSmilTarget, , , SMILTARGET_NOCHANGE
  'end new 20030917
    
  sBaseKey = "Software\DaisyWare\Regenerator\Validation"
  
  fncLoadRegistryData "Validate job", bolUseValidator, , , True
  fncLoadRegistryData "Include warnings", bolIncludeWarnings, , , False
  fncLoadRegistryData "Include NC errors", bolIncludeNCErrors, , , True
  fncLoadRegistryData "Include advanced ADTD information", bolIncludeADVADTD, , , _
    False
  fncLoadRegistryData "Create standalone report", bolCreateStandalone, , , False
  fncLoadRegistryData "Save standalone report", sStandalonePath, , , _
    "*fulldtbpath*\regenerated\report.html"
  
  fncLoadRegistryData "Synchronize settings with validator", bolSyncWValidator, , _
    , True
  
  fncLoadRegistryData "Validator Light Mode", bolValidatorLightMode, , _
    , False
  
  Dim sValDTDPath As String, sValAppPath As String, sValTempPath As String
  Dim lValTimeFluct As Long, sValVTMPath As String
  
' If there are no registry settings try to load the Daisy Validators registry keys

  sBaseKey = "Software\DaisyWare\Validator\Misc"
  fncLoadRegistryData "Dtd_AdtdPath", sValDTDPath, , , ""
  fncLoadRegistryData "AppPath", sValVTMPath, , , ""
  fncLoadRegistryData "TempPath", sValTempPath, , , sValue
  sBaseKey = "Software\DaisyWare\Validator\Settings"
  fncLoadRegistryData "TimeFluctuation", lValTimeFluct, , , 0
  If bolSyncWValidator Then
    sExtPath = sValDTDPath
    sVTMPath = sValVTMPath
    sTempPath = sValTempPath
    lTimeFluctuation = lValTimeFluct
  Else
    sBaseKey = "Software\DaisyWare\Regenerator\Validation"
    fncLoadRegistryData "Externals path", sExtPath, , , sValDTDPath
    fncLoadRegistryData "VTM path", sVTMPath, , , sValVTMPath
    fncLoadRegistryData "Temp path", sTempPath, , , sValTempPath
    fncLoadRegistryData "Time fluctuation", lTimeFluctuation, , , lValTimeFluct
  End If
  
  'if the validator is not installed or something is wrong
  If (sValDTDPath = "" Or sValVTMPath = "" Or sExtPath = "") Then _
    bolUseValidator = False

  sBaseKey = "Software\DaisyWare\Regenerator\Log"
  fncLoadRegistryData "Save log", sLogPath, , , sAppPath & "log\"
  
  objRegeneratorUserControl.fncSetAudioTemplatesPath sAppPath & "audiotemplates\"
  objRegeneratorUserControl.fncSetResourcePath sAppPath & "resources\"
  objRegeneratorUserControl.fncSetTidyLibPath sAppPath & "tidylib\"
  
End Function

Public Function fncSaveRegistrySettingsOld()
  sBaseKey = "Software\DaisyWare\Regenerator\Last properties"
  
  fncSaveRegistryData "DTB Type", lDtbType
  fncSaveRegistryData "Input charset", lCharset
  fncSaveRegistryData "IANA charset", lIANACharset
  fncSaveRegistryData "Preserve meta", bolPreserveMeta
  fncSaveRegistryData "Sequential rename", bolSeqRename
  fncSaveRegistryData "Use numeric portion", bolUseNumeric
  fncSaveRegistryData "Same folder", bolSameFolder
  fncSaveRegistryData "Move book", bolMoveBook

  sBaseKey = "Software\DaisyWare\Regenerator\General"
  
  fncSaveRegistryData "Default save path", sDefaultSavePath
  fncSaveRegistryData "Default meta path", sDefaultMetaPath
  fncSaveRegistryData "Halt 'on error", bolHalt
  fncSaveRegistryData "Fix par", bolFixPar
  fncSaveRegistryData "Pb2k Layout fix", bolPb2kLayoutFix
  fncSaveRegistryData "Mangle links", bolRebuildLinkStructure
  fncSaveRegistryData "Disable broken xhtml links", bolDisableBrokenXhtmlLinks
  fncSaveRegistryData "Merge phrases", bolMergeShortPhrases
  fncSaveRegistryData "Merge short", lClipLessThan
  fncSaveRegistryData "Merge short 2", lFirstClipLessThan
  fncSaveRegistryData "Merge next less than", lNextClipLessThan
  fncSaveRegistryData "Merge maximum span", lClipSpan
  
  'new 20030917
  fncSaveRegistryData "Estimate broken xhtml links", bolEstimateBrokenXhtmlLinks
  fncSaveRegistryData "Verbose log", bolDoVerboseLog
  fncSaveRegistryData "Add default css", bolAddCss
  fncSaveRegistryData "Make true NCC-Only", bolMakeTrueNccOnly
  'fncSaveRegistryData "Redirect smil targets to par", bolPointTargetsToPar
  fncSaveRegistryData "smil targets", lSmilTarget
  'end new 20030917

  sBaseKey = "Software\DaisyWare\Regenerator\Validation"

  fncSaveRegistryData "Validate job", bolUseValidator
  fncSaveRegistryData "Include warnings", bolIncludeWarnings
  fncSaveRegistryData "Include NC errors", bolIncludeNCErrors
  fncSaveRegistryData "Include advanced ADTD information", bolIncludeADVADTD
  fncSaveRegistryData "Create standalone report", bolCreateStandalone
  fncSaveRegistryData "Save standalone report", sStandalonePath
  
  fncSaveRegistryData "Synchronize settings with validator", bolSyncWValidator
  fncSaveRegistryData "Validator Light Mode", bolValidatorLightMode
  
  fncSaveRegistryData "Externals path", sExtPath
  fncSaveRegistryData "VTM path", sVTMPath
  fncSaveRegistryData "Temp path", sTempPath
  fncSaveRegistryData "Time fluctuation", lTimeFluctuation

  sBaseKey = "Software\DaisyWare\Regenerator\Log"
  fncSaveRegistryData "Save log", sLogPath
End Function

