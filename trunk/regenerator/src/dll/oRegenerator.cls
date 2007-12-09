VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oRegenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Daisy 2.02 Regenerator DLL
' Copyright (C) 2003 Daisy Consortium
'
'    This file is part of Daisy 2.02 Regenerator.
'
'    Daisy 2.02 Regenerator is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Daisy 2.02 Regenerator is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Daisy 2.02 Regenerator; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


Option Explicit

Private sprivDtbFolderPath As String 'the input ncc path minus filename
Private sprivBackupPath As String 'folder path, ends with "\", no filename
Private sprivUnrefPath As String 'folder path, ends with "\", no filename
Private sprivOutCharSetName As String 'carries charset name for output files
Private sprivDcIdentifier As String 'carries dc:identifier for ui access
Private lprivTotalTimeMs As Long
Private sprivEmptyMp3FileName As String
Private bolprivVerboseLog As Boolean
'added 20041005:
Private bolprivPreserveBibliometa As Boolean

Private objprivFileSetHandler As oFileSetHandler
Private objprivCommonMeta As oCommonMeta
Private objprivXhtmlEntities As oXhtmlEntities
Private objprivLangCodes As oLangCodes

Private aLogItems() As String
Private lLogItems As Long

Public Event evLogItem(sLogItem As String)
Public objOwner As Object

' **** Friend methods / properties
Friend Property Let sDtbFolderPath(isDtbFolderPath As String)
  sprivDtbFolderPath = isDtbFolderPath
End Property

Friend Property Get sDtbFolderPath() As String
  sDtbFolderPath = sprivDtbFolderPath
End Property

Friend Property Let sBackupPath(isBackupPath As String)
  sprivBackupPath = isBackupPath
End Property

Friend Property Get sBackupPath() As String
  sBackupPath = sprivBackupPath
End Property

Friend Property Let sUnrefPath(isUnrefPath As String)
  sprivUnrefPath = isUnrefPath
End Property

Friend Property Get sUnrefPath() As String
  sUnrefPath = sprivUnrefPath
End Property

Friend Property Let sOutCharsetName(isOutCharsetName As String)
  sprivOutCharSetName = isOutCharsetName
End Property

Friend Property Get sOutCharsetName() As String
  sOutCharsetName = sprivOutCharSetName
End Property

Friend Property Get bolVerboseLog() As Boolean
 bolVerboseLog = bolprivVerboseLog
End Property

Friend Property Let sEmptyMp3Filename(isEmptyMp3Filename As String)
  sprivEmptyMp3FileName = isEmptyMp3Filename
End Property

Friend Property Get sEmptyMp3Filename() As String
  sEmptyMp3Filename = sprivEmptyMp3FileName
End Property

Friend Property Let lTotalTimeMs(ilTotalTimeMs As Long)
  lprivTotalTimeMs = ilTotalTimeMs
End Property

Friend Property Get lTotalTimeMs() As Long
  lTotalTimeMs = lprivTotalTimeMs
End Property

Friend Property Get objCommonMeta() As oCommonMeta
  Set objCommonMeta = objprivCommonMeta
End Property

Friend Property Get objXhtmlEntities() As oXhtmlEntities
  Set objXhtmlEntities = objprivXhtmlEntities
End Property

Friend Property Get objLangCodes() As oLangCodes
  Set objLangCodes = objprivLangCodes
End Property

Friend Function fncSetDcIdentifier(sDcId As String)
  sprivDcIdentifier = sDcId
End Function

Public Function fncGetDcIdentifier() As String
  'used to get dc:identifier from the outside
  fncGetDcIdentifier = sprivDcIdentifier
End Function

Public Function bPreserveBiblioMeta() As Boolean
  'added 20041005 when fixing a malformed meta before tidy run on ncc or content
  bPreserveBiblioMeta = bolprivPreserveBibliometa
End Function

Friend Sub addlog(isLogItem As String)
Dim bolInNonVerbose As Boolean
  
  If Not (Me.bolVerboseLog) Then
    If (InStr(1, isLogItem, "warning", vbTextCompare) > 1) Then bolInNonVerbose = True
    If (InStr(1, isLogItem, "errH", vbTextCompare) > 1) Then bolInNonVerbose = True
    If (InStr(1, isLogItem, "processFailure", vbTextCompare) > 1) Then bolInNonVerbose = True
    If (InStr(1, isLogItem, "error", vbTextCompare) > 1) Then bolInNonVerbose = True
    If (InStr(1, isLogItem, "identifiers", vbTextCompare) > 1) Then bolInNonVerbose = True
    If (InStr(1, isLogItem, "dc:identifier", vbTextCompare) > 1) Then bolInNonVerbose = True
    If (InStr(1, isLogItem, "dc:title", vbTextCompare) > 1) Then bolInNonVerbose = True
    If (InStr(1, isLogItem, "date>", vbTextCompare) > 1) Then bolInNonVerbose = True
    If (InStr(1, isLogItem, "time>", vbTextCompare) > 1) Then bolInNonVerbose = True
    If Not (bolInNonVerbose) Then Exit Sub
  End If
    
  lLogItems = lLogItems + 1
  ReDim Preserve aLogItems(lLogItems)
  aLogItems(lLogItems) = isLogItem
  Debug.Print isLogItem
  RaiseEvent evLogItem(isLogItem)
  subRaiseVBFriendlyEvent "Regenerator.AddLog", isLogItem, ""
  DoEvents
End Sub

Private Sub Class_Initialize()
  sAppVersion = App.Major & "." & App.Minor & "." & App.Revision
  Set objprivFileSetHandler = New oFileSetHandler
  Set objprivFileSetHandler.objOwner = Me
  Set objprivCommonMeta = New oCommonMeta
  Set objprivXhtmlEntities = New oXhtmlEntities
  If Not objprivXhtmlEntities.fncInitialize(Me) Then addlog "<warning>warning: xhtml entity object not intitialized</warning>"
  Set objprivLangCodes = New oLangCodes
  If Not objprivLangCodes.fncInitialize(Me) Then addlog "<warning>warning: lang code object not intitialized</warning>"
End Sub

Public Property Get lLogItemCount() As Long
  lLogItemCount = lLogItems
End Property

Public Function fncRetrieveLogItem(lItem As Long) As String
  If lItem < 1 Or lItem > lLogItems Then Exit Function
  fncRetrieveLogItem = aLogItems(lItem)
End Function

Public Function fncTerminateObject() As Boolean
  Dim bolResult As Boolean
  bolResult = True
  If Not objprivCommonMeta.fncTerminateObject Then bolResult = False
  If Not objprivFileSetHandler.fncTerminateObject Then bolResult = False
  If Not objprivXhtmlEntities.fncTerminateObject Then bolResult = False
  If Not objprivLangCodes.fncTerminateObject Then bolResult = False
  Set objprivCommonMeta = Nothing
  Set objprivFileSetHandler = Nothing
  Set objprivXhtmlEntities = Nothing
  Set objprivLangCodes = Nothing
  fncTerminateObject = bolResult
End Function

Private Sub Class_Terminate()
  Set objprivCommonMeta = Nothing
  Set objprivFileSetHandler = Nothing
  Set objprivXhtmlEntities = Nothing
  Set objprivLangCodes = Nothing
End Sub

Private Sub subRaiseVBFriendlyEvent(isEvent As String, vParam1 As Variant, _
  vParam2 As Variant)
  If objOwner Is Nothing Then Exit Sub
  objOwner.subVBFriendlyEvent isEvent, vParam1, vParam2
End Sub

Public Property Get objFileSetHandler() As oFileSetHandler
  Set objFileSetHandler = objprivFileSetHandler
End Property

Public Function fncRunTransform( _
    sOrigNccPath As String, _
    lInputEncoding As Long, sWantsThisEncoding As String, _
    lDtbType As Long, _
    bolIsMultiVolume As Boolean, _
    bolPreserveBiblioMeta As Boolean, _
    sMetaPath As String, _
    bolSeqRename As Boolean, _
    bolUseDcIdNum As Boolean, _
    sWantedPrefix As String, _
    bolRebuildLinkStructure As Boolean, _
    bolFixMultiTextInPar As Boolean, _
    bolMergeShortPhrases As Boolean, _
    lClipSpan As Long, _
    lClipLessThan As Long, _
    lFirstClipLessThan As Long, _
    lNextClipLessThan As Long, _
    bolDisableBrokenXhtmlLinks As Boolean, _
    bolEstimateBrokenXhtmlLinks As Boolean, _
    bolPb2kLayoutFix As Boolean, _
    bolDoVerboseLog As Boolean, _
    bolAddCss As Boolean, _
    bolMakeTrueNccOnly As Boolean, _
    lSmilTarget As Long _
    ) As Boolean

Dim oNccDom As New MSXML2.DOMDocument40
    oNccDom.async = False
    oNccDom.validateOnParse = False
    oNccDom.resolveExternals = False
    oNccDom.preserveWhiteSpace = False
    oNccDom.setProperty "SelectionLanguage", "XPath"
    oNccDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
    oNccDom.setProperty "NewParser", True

  On Error GoTo ErrHandler
  
  fncRunTransform = False
  lInstancesRunning = lInstancesRunning + 1
  bolprivVerboseLog = bolDoVerboseLog
  bolprivPreserveBibliometa = bolPreserveBiblioMeta
        
  addlog "<regenerating dllVersion='" & sAppVersion & "'>"
  addlog "<stamp type='begin'>"
  addlog "<time>" & Date & " " & Time & ": begin.</time>"
  addlog "</stamp>"
              
  If Not fncFileExists(sOrigNccPath, Me) Then
    addlog "<error>" & sOrigNccPath & " does not exist. Aborting.</error>"
    GoTo ErrHandler
  End If
                      
  sDtbFolderPath = fncGetPathName(sOrigNccPath)
  
  sEmptyMp3Filename = "rgn_empty_48.mp3"
    
  If Not fncCheckUtf8Encoding(lInputEncoding, sOrigNccPath, Me) Then GoTo ErrHandler
    
  If Not fncSetEncodingName(lInputEncoding, sprivOutCharSetName, sWantsThisEncoding, Me) Then GoTo ErrHandler
      
  addlog "<inParams>"
    addlog "<sOrigNccPath value='" & sOrigNccPath & "'/>"
    addlog "<encoding>"
      addlog "<lInputEncoding value='" & CStr(lInputEncoding) & "'/>"
      addlog "<sWantsThisEncoding value='" & sWantsThisEncoding & "'/>"
      addlog "<derivedOutCharset value='" & sOutCharsetName & "'/>"
    addlog "</encoding>"
    addlog "<lDtbType value='" & CStr(lDtbType) & "'/>"
    addlog "<bolPreserveBiblioMeta value='" & CStr(bolPreserveBiblioMeta) & "'/>"
    addlog "<bolSeqRename value='" & CStr(bolSeqRename) & "'/>"
    addlog "<bolUseDcIdNum value='" & CStr(bolUseDcIdNum) & "'/>"
    addlog "<sWantedPrefix value='" & sWantedPrefix & "'/>"
    addlog "<bolIsMultiVolume value='" & CStr(bolIsMultiVolume) & "'/>"
    addlog "<bolFixMultiTextInPar value='" & CStr(bolFixMultiTextInPar) & "'/>"
    addlog "<bolRebuildLinkStructure value='" & bolRebuildLinkStructure & "'/>"
    addlog "<sMetaPath value='" & sMetaPath & "'/>"
    addlog "<bolMergeShortPhrases value='" & bolMergeShortPhrases & "'/>"
    If bolMergeShortPhrases Then
      addlog "<lClipSpan value='" & CStr(lClipSpan) & "'/>"
      addlog "<lClipLessThan value='" & CStr(lClipLessThan) & "'/>"
      addlog "<lFirstClipLessThan value='" & CStr(lFirstClipLessThan) & "'/>"
      addlog "<lNextClipLessThan value='" & CStr(lNextClipLessThan) & "'/>"
    End If
    addlog "<bolDisableBrokenXhtmlLinks value='" & bolDisableBrokenXhtmlLinks & "'/>"
    addlog "<bolEstimateBrokenXhtmlLinks value='" & bolEstimateBrokenXhtmlLinks & "'/>"
    addlog "<bolPb2kLayoutFix value='" & CStr(bolPb2kLayoutFix) & "'/>"
    addlog "<bolDoVerboseLog value='" & CStr(bolDoVerboseLog) & "'/>"
    addlog "<bolAddCss value='" & CStr(bolAddCss) & "'/>"
    addlog "<bolMakeTrueNccOnly value='" & CStr(bolMakeTrueNccOnly) & "'/>"
    addlog "<lSmilTarget value='" & CStr(lSmilTarget) & "'/>"
  addlog "</inParams>"
    
  'add ncc as first member of input fileset array
  objprivFileSetHandler.fncAddObjectToInputArray TYPE_NCC, sOrigNccPath
    
  'try to tidy the ncc
  If Not fncTidyNcc(oNccDom, sOrigNccPath, lInputEncoding, sOutCharsetName, Me) Then GoTo ErrHandler
      
  'if we are here, ncc is now successfully outfilearray item (0). Create rest of fileset.
  If Not objFileSetHandler.fncCreateFileSetArrays(oNccDom, lInputEncoding, sOutCharsetName) Then GoTo ErrHandler
    
  'give the ncc correct spelling (references to this file are fixed in in mNamesAndUris)
  objprivFileSetHandler.aOutFileSet(0).sFileName = "ncc.html"
    
  'check the two arrays for relational consistency
  If Not objFileSetHandler.fncCheckArrayConsistency Then GoTo ErrHandler

  '*****************************************************************************
  '* at this point, both input array and output array are filled
  '* with everything except mastersmil which will be rendered from scratch later.
  '* All documents have been tidied and testparsed for wellformedness with msxml
  '* all files have had fileExists checked. Arrays have been checked for consistency.
  '* Now, loop through the outFileSet array
  '* and call internal fix functions as appropriate for type
  '* let oNccDom be live to and reuse same instance byref
  '***********************************************************************

  'Stop
  'Debug.Print objFileSetHandler.aInFileSet(2).sAbsPath
  'Stop

  addlog "<status>initiating fix iterations...</status>"
  Dim i As Long, lSmilNumber As Long: lSmilNumber = 0
  For i = 0 To objprivFileSetHandler.aOutFileSetMembers - 1
    Select Case objprivFileSetHandler.aOutFileSet(i).eType
      Case TYPE_NCC
        If Not fncInternalNcc(oNccDom, bolPreserveBiblioMeta, sMetaPath, sOutCharsetName, bolIsMultiVolume, lDtbType, bolAddCss, Me) Then GoTo ErrHandler
      Case TYPE_SMIL_CONTENT
        If Not fncInternalContent(oNccDom, i, bolPreserveBiblioMeta, bolAddCss, Me) Then GoTo ErrHandler
      Case TYPE_SMIL_1
        lSmilNumber = lSmilNumber + 1
        If Not fncInternalSmil(oNccDom, i, bolPreserveBiblioMeta, bolIsMultiVolume, bolFixMultiTextInPar, bolMergeShortPhrases, lClipSpan, lClipLessThan, lFirstClipLessThan, lNextClipLessThan, lSmilNumber, Me) Then GoTo ErrHandler
      Case Else
     End Select
  Next i
  'since pars may have been moved around, do timecalc and id add here, not above
  addlog "<status>timing information...</status>"
  lSmilNumber = 0
  For i = 0 To objprivFileSetHandler.aOutFileSetMembers - 1
    Select Case objprivFileSetHandler.aOutFileSet(i).eType
    Case TYPE_SMIL_1
      lSmilNumber = lSmilNumber + 1
      If Not fncInternalSmilTimeAndId(oNccDom, i, bolIsMultiVolume, lSmilNumber, Me) Then GoTo ErrHandler
    End Select
  Next i
                                                                                                                                          
'  'set lang attrs in ncc and content
'  moved this into internal ncc mg 20030913
'  If Not fncSetLangAttrs(oNccDom, Me, 0) Then GoTo ErrHandler
  
  'finalize and close the ncc dom now since fncNamesAndUris below uses array domdata updates
  '(for smil and content, this is done at end of "mInternal.." in the loops above
  objprivFileSetHandler.aOutFileSet(0).sDomData = oNccDom.xml
  Set oNccDom = Nothing
      
  '*******************************************************************************
  '* at this point, output array contains documents fixed internally, except URIs and names
  '*******************************************************************************
  'Stop
  If Not fncNamesAndUris(bolSeqRename, bolUseDcIdNum, sWantedPrefix, Me) Then GoTo ErrHandler
  
'  Debug.Print objprivFileSetHandler.aOutFileSet(10).sFileName
'  Debug.Print objprivFileSetHandler.aOutFileSet(10).sDomData
'  Stop
                                                
  If bolRebuildLinkStructure Then
    If Not fncRebuildLinkStructure(Me, bolEstimateBrokenXhtmlLinks) Then GoTo ErrHandler
  End If
                                                
  'disable ncc links that point to void
'  If bolDisableBrokenNccLinks Then
'    If Not fncDisableBrokenNccLinks(Me) Then GoTo ErrHandler
'  End If
  
  'mg 20030911: rewrite of the above
  If bolDisableBrokenXhtmlLinks Then
    If Not fncDisableBrokenXhtmlLinks(bolEstimateBrokenXhtmlLinks, Me) Then GoTo ErrHandler
  End If
                                                  
  If bolMakeTrueNccOnly Then
    If Not fncMakeTrueNccOnly(lDtbType, Me) Then GoTo ErrHandler
  End If
                                                  
  If lSmilTarget > 0 Then 'if lSmilTarget = 0 then SMILTARGET_NOCHANGE
    If Not fncAdjustSmilTargetPointers(lSmilTarget, Me) Then GoTo ErrHandler
  End If
                                                                                                    
  'mg20040216 special fix for sdar2 books
  'do it only if book is made true ncc only
  If bolMakeTrueNccOnly Then
    If Not fncSdar2Fix(Me) Then GoTo ErrHandler
  End If
                                                                                                    
  If Not fncUpdateTotalTimeAndFiles(lTotalTimeMs, Me) Then GoTo ErrHandler
                                                                                                    
  'generate mastersmil
  If Not fncGenerateMasterSmil(sOutCharsetName, Me) Then GoTo ErrHandler
    
  'pretty print
  If Not fncPrettyPrint(Me, bolPb2kLayoutFix) Then GoTo ErrHandler
    
  'reenter namespace
  If Not fncAddNamesSpaces(Me) Then GoTo ErrHandler
    
'  addlog "<stamp type='end'>"
'  addlog "<time>" & Time & "</time>"
'  addlog "</stamp>"
  
  fncRunTransform = True

ErrHandler:
  If Not fncRunTransform Then
    addlog "<errH in='fncRunTransform'>fncRunTransform ErrH (" & Err.Number & " : " & Err.Description & ")</errH>"
    addlog "<processFailure in='regeneration'>Regeneration process failed</processFailure>"
    If Me.bolVerboseLog Then
      'print the array names
      objprivFileSetHandler.fncPrintFileSet Me
    End If
  End If
  lInstancesRunning = lInstancesRunning - 1
  addlog "</regenerating>"
End Function

Public Function fncRenderFiles( _
    ByVal lRenderType As Long, _
    ByVal sRenderPath As String _
    ) As Boolean
Dim i As Long
  
  On Error GoTo ErrHandler
  fncRenderFiles = False
  addlog "<rendering>"
  addlog "<inParams>"
      addlog "<lRenderType value='" & CStr(lRenderType) & "'/>"
      addlog "<sRenderPath value='" & sRenderPath & "'/>"
  addlog "</inParams>"

  lInstancesRunning = lInstancesRunning + 1
  
  '********************************************************************
  '* renderType "replace" means: make backup of all xml, save new xml,
  '* rename audio and other, create renamelog,
  '* all in original folder
  '* renderType "new folder" means: save all xml to new folder,
  '* copy and rename audio+other there
  '*******************************************************************
    
  Select Case lRenderType
    Case RENDER_REPLACE
      
      '****************************************************
      '* create backuppath
      '* create rename log file, place in backuppath
      '* move source array xml documents to backuppath
      '* rename all files but xml documents
      '* if unreferenced files exist,
      '* create unrefd folder
      '* move unreferenced files to unrefpath
      '* save xml documents from outarray
      '****************************************************
    
      'create backuppath
      Dim sDateString As String
      sDateString = fncGetDateString
      sBackupPath = fncCreateFolder(sDtbFolderPath & "rgn~bkp_" & sDateString & "\", Me, True)
      
      'create rename log file, place in backuppath
      If Not fncCreateFileNameLog(Me) Then GoTo ErrHandler
      
      'move source array xml documents to backuppath
      addlog "<status>making xml doc backups...</status>"
      For i = 0 To objprivFileSetHandler.aInFileSetMembers - 1
        If (objprivFileSetHandler.aInFileSet(i).eType = TYPE_NCC) Or _
          (objprivFileSetHandler.aInFileSet(i).eType = TYPE_SMIL_1) Or _
          (objprivFileSetHandler.aInFileSet(i).eType = TYPE_SMIL_CONTENT) Then
            If Not fncMoveFile(objprivFileSetHandler.aInFileSet(i).sAbsPath, _
              sBackupPath, Me) Then GoTo ErrHandler
        End If
        DoEvents
      Next i
      
      '20041005 continue move: there may be a preexisting master.smil which is not in array since never referenced
      'and may be write protected
      If fncFileExists(sDtbFolderPath & "master.smil", Me) Then
        If Not fncMoveFile(sDtbFolderPath & "master.smil", _
              sBackupPath, Me) Then GoTo ErrHandler
      End If
                                                      
      ' rename all files but xml documents
      addlog "<status>renaming non-xml media members...</status>"
      
      Dim sSafetyString As String
      'add safetystring to rename destination
      'to make sure there is no conflict with existing files that are to be renamed later
      sSafetyString = Replace(Replace(Replace(Now, " ", ""), "-", ""), ":", "") & "regentemp"
            
      For i = objprivFileSetHandler.aInFileSetMembers - 1 To 0 Step -1
        If (objprivFileSetHandler.aInFileSet(i).eType = TYPE_SMIL_AUDIO) Or _
          (objprivFileSetHandler.aInFileSet(i).eType = TYPE_SMIL_IMG) Or _
          (objprivFileSetHandler.aInFileSet(i).eType = TYPE_OTHER) Then
          If Not fncRenameFilesetFile(objprivFileSetHandler.aInFileSet(i).sAbsPath, _
          fncGetPathName(objprivFileSetHandler.aInFileSet(i).sAbsPath) & objprivFileSetHandler.aOutFileSet(i).sFileName & sSafetyString, sBackupPath, Me) Then _
            GoTo ErrHandler
        End If
        DoEvents
      Next i
            
      'now remove the safetystring from each file
      If Not fncRemoveStringFromEachFileInFolder(sDtbFolderPath, sSafetyString, Me) Then GoTo ErrHandler
            
      'if unreferenced files exist, create unrefd folder, move unreferenced files to unrefpath
      addlog "<status>moving unreferenced files...</status>"
      If Not fncMoveUnref(sDateString, Me) Then GoTo ErrHandler
      
      ' save xml documents from outarray
      addlog "<status>saving xml members...</status>"
      If Not fncSaveAllXmlInOutArray(sDtbFolderPath, sOutCharsetName, Me) Then GoTo ErrHandler
    
    Case RENDER_NEWFOLDER
      '****************************************************
      '* check if inputparam folder preexists, else create it
      '* if preexists, check that its empty, else break
      '* save all xml array items to sRenderPath
      '* copy all non-xml from original, inc rename
      '****************************************************
      
      'check if inputparam folder preexists, else create it
      'if preexists, check that its empty
      
      If Right(sRenderPath, 1) <> "\" Then sRenderPath = sRenderPath & "\"
      If Not fncFolderExists(sRenderPath, Me) Then
        addlog "<status>creating folder " & sRenderPath & "</status>"
        'fncCreateFolder sRenderPath, Me
        fncCreateDirectoryChain sRenderPath, Me 'changed mg 20030219
      Else
        'If fncFileCount(sRenderPath) > 0 Then 'changed mg 20030219
        If (fncFileExists(sRenderPath & "ncc.html", Me)) Or (fncFileExists(sRenderPath & "ncc.htm", Me)) Then
          addlog "<error>" & sRenderPath & "appears to contain DTB data. Aborting.</error>"
          GoTo ErrHandler
        End If
      End If
      'save all xml array items to sRenderPath
      addlog "<status>saving xml members...</status>"
    
      If Not fncSaveAllXmlInOutArray(sRenderPath, sOutCharsetName, Me) Then GoTo ErrHandler
      
      'copy all non-xml from original, inc rename
      addlog "<status>copying/renaming non-xml members...</status>"
            
      'get base dir of input ncc for use in case else below
      Dim sBaseOrigDir As String, sCurrOrigDir As String
      sBaseOrigDir = fncGetPathName(objprivFileSetHandler.aInFileSet(0).sAbsPath)
      If Right(sBaseOrigDir, 1) <> "\" Then sBaseOrigDir = sBaseOrigDir & "\"
      
      For i = 1 To objprivFileSetHandler.aInFileSetMembers - 1
        Select Case objprivFileSetHandler.aOutFileSet(i).eType
          Case TYPE_NCC, TYPE_SMIL_1, TYPE_SMIL_CONTENT, TYPE_SMIL_MASTER, TYPE_SMIL_AUDIO_INSERTED, TYPE_DELETED
            'do nothing
          Case Else
            'copy audio, img, css etc
            'note that these may be in subfolders to the base dir ("/images/bla.jpg etc")
            'therefore do as follows
            'get the path for current file
            sCurrOrigDir = fncGetPathName(objprivFileSetHandler.aInFileSet(i).sAbsPath)
            If Right(sCurrOrigDir, 1) <> "\" Then sCurrOrigDir = sCurrOrigDir & "\"
            'compare it to base dir
            If LCase$(sBaseOrigDir) <> LCase$(sCurrOrigDir) Then
              'if different, then
              'current file to be copied to a subfolder of sBaseOrigDir
              Dim sNewSubFolder As String
              'get the name of the subfolder
              sNewSubFolder = fncGetMotherFolder(sCurrOrigDir)
              'if it doesnt exist, create it
              If Not fncFolderExists(sRenderPath & sNewSubFolder, Me) Then
                fncCreateDirectoryChain sRenderPath & sNewSubFolder, Me
              End If
              If Not Right(sNewSubFolder, 1) = "\" Then sNewSubFolder = sNewSubFolder & "\"
            Else
              'current file is in same dir as base (=ncc dir)
              'set CurrOrigDir to null so that copy below works for both cases
              sNewSubFolder = ""
            End If
                        
            If Not fncCopyFile( _
                objprivFileSetHandler.aInFileSet(i).sAbsPath, _
                sRenderPath & sNewSubFolder & objprivFileSetHandler.aOutFileSet(i).sFileName, Me _
                ) Then
              addlog "<error>renderfiles fncCopyFile error</error>"
              GoTo ErrHandler
            End If
        End Select
        DoEvents
      Next i
    Case Else
      addlog "<error>rendertype selection invalid</error>"
      GoTo ErrHandler
  End Select
   
  'regardless of savetype,
  'fncFixMultiTextInPar may have inserted a reference
  'to an audiofile that does not exist in orig fileset
  'get its name from array
  'copy this from app.path & /audioTemplates
  'into dtb destination dir
  'note that for each time this happens
  'ncc:files is upped by one - but ncc:files is already updated so dont worry here
  'mg20030914: do the same for added stylesheets
   
  Dim sPath As String
  Select Case lRenderType
    Case RENDER_REPLACE
      sPath = sDtbFolderPath
    Case RENDER_NEWFOLDER
      sPath = sRenderPath
  End Select
 
  For i = 0 To objprivFileSetHandler.aOutFileSetMembers - 1
    If (objprivFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_AUDIO_INSERTED) Then
      If Not fncCopyFile(sAudioTemplatesPath & _
        objprivFileSetHandler.aOutFileSet(i).sFileName, sPath, Me) Then
          addlog "<error>file could not be copied: " & sAudioTemplatesPath & _
            objprivFileSetHandler.aOutFileSet(i).sFileName & "</error>"
        GoTo ErrHandler
      End If
    End If
    If (objprivFileSetHandler.aOutFileSet(i).eType = TYPE_CSS_INSERTED) Then
      If Not fncCopyFile(sResourcePath & _
        objprivFileSetHandler.aOutFileSet(i).sFileName, sPath, Me) Then
          addlog "<error>file could not be copied: " & sResourcePath & _
            objprivFileSetHandler.aOutFileSet(i).sFileName & "</error>"
        GoTo ErrHandler
      End If
    End If
    
    DoEvents
  Next i
      
  addlog "<stamp type='end'>"
  addlog "<time>" & Time & ": done.</time>"
  addlog "</stamp>"
  fncRenderFiles = True

ErrHandler:
  If Not fncRenderFiles Then
   addlog "<errH in='fncRenderFiles'>fncRenderFiles ErrH (" & Err.Number & " : " & Err.Description & ")</errH>"
   addlog "<processFailure in='render'>Render process failed</processFailure>"
  End If
  addlog "</rendering>"
  lInstancesRunning = lInstancesRunning - 1
End Function

Private Function fncGetDateString() As String
  On Error GoTo ErrHandler
  fncGetDateString = Format(Now, "yyyy") & Format(Now, "mm") & Format(Now, "dd") & Format(Now, "Hh") & Format(Now, "Nn")
  Exit Function
ErrHandler:
 fncGetDateString = "unknown_date"
End Function
