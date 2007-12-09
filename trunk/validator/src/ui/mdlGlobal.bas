Attribute VB_Name = "mdlGlobal"
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

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
   ByVal lpParameters As String, ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Public lTestsSucceeded As Long  'catch event raised by dll
Public lTestsFailed As Long     'catch event raised by dll
Public lErrorLogs As Long       'catch event raised by dll
Public bolQuit As Boolean
Public sCurrentState As String 'contains current state of process; "idle", "validating", "aborted"

Public Enum enuQueueMode 'used as input parameter to the runqueue
    runall
    runselected
    runchecked
End Enum

Public Enum eFileType 'used for commondialog save as
    eventlog
    report
    filesetfile
End Enum

Public Const TYPE_SINGLEDTB = 0
Public Const TYPE_MULTIVOLUME = 1
Public Const TYPE_SINGLE_NCC = 2
Public Const TYPE_SINGLE_SMIL = 3
Public Const TYPE_SINGLE_MSMIL = 4
Public Const TYPE_SINGLE_CONTENTDOC = 5
Public Const TYPE_SINGLE_DISCINFO = 6
'Public Const TYPE_SINGLEDTB_LIGHT = 7
'Public Const TYPE_MOTHERDIR = 8


Public Type tCandidateInfo
    sAbsPath As String
    lCandidateType As Long
    objReport As oReport
    bolSelected As Boolean
    bolChecked As Boolean
End Type

Public aCandidateQueue() As tCandidateInfo 'the array where added candidates rest waiting for run
Public lCandidatesAdded As Long 'number of candidates in aCandidateQueue
Public lCurrentCandidate As Long    'used for statusbar: adctive candidate in queue

Public Type tFileHistory
  sAbsPath As String
  lBytePos As Long
  sID As String
End Type

Public aFileHistory(10) As tFileHistory, lCurrentHistory As Long
Public sDefaultReportPath As String

Public lIntProgress As Long, lIntProgMax As Long
Public lTimeFluct As Long, bolADTDAdvanced As Boolean

Public sAppPath As String, bolLightMode As Boolean, bolDisableAudioTests As Boolean

Sub Main()

 '******************** fixlog **********************
 'mg 20050330
 ' fi.hasRecommendedName added and run in lightmode
 'mg 20050324
 ' - files referenced via url() statements in CSS are now detected
 ' - reinserted test for first h1 of NCC having @class='title' - this test has mistakenly been omitted since the prior release.
 ' note: the h1/@class='title' test is run in both light- and fullmode. The absence of this attribute and its value causes problems in a certain player.
 ' - audiofile integrity (isValidAudioFile, hasValidExtension) tests now also run in lightmode
 
 '  mg20030325 fixed ncc:revision and ncc:revisionDate in ncc.adtd from #?r to #?
 '  mg20030325 fixed et loop when replace all and ex "<a" --> "<A"
 '**************************************************
  
 ' ** KE 2004-06-23 <
  Dim oTemp As Object
  Dim bMsxmlPresenceIsTested As Boolean
  
  On Error GoTo errH
  'MsgBox "1"
  bMsxmlPresenceIsTested = True
  Set oTemp = CreateObject("Msxml2.DOMDocument.4.0")
  If oTemp Is Nothing Then
    MsgBox "MSXML 4 not installed, exiting program.", vbOKOnly
    Exit Sub
  End If
  bMsxmlPresenceIsTested = False
 ' ** >
  'MsgBox "2"
  sAppPath = App.Path
  If Not Right$(sAppPath, 1) = "\" Then sAppPath = sAppPath & "\"
  'MsgBox "3"
  SetEventCounters
  sCurrentState = "idle"
  lCandidatesAdded = 0
  frmMain.Form_Load
  'MsgBox "4"
  fncLoadFromRegistry
  'MsgBox "5"
  subApplyUserLcid
  'MsgBox "6"
  frmMain.SetFocus
  'MsgBox "7"
'  Do
'    DoEvents
'  Loop Until bolQuit
'
'  frmMain.tmUiUpdate.Enabled = False
'  Unload frmAbout
'  Unload frmCandidateList
'  Unload frmDocuPaths
'  Unload frmErrorLog
'  Unload dlgSearchReplace
'  Unload frmMain

  Exit Sub
errH:
  If bMsxmlPresenceIsTested Then
    MsgBox "MSXML 4 is not installed. Exiting program.", vbOKOnly
  Else
    MsgBox "An unknown error occured. Exiting program.", vbOKOnly
  End If
End Sub

Private Function fncLoadFromRegistry()
  Dim sEExePath As String, sTempPath As String, sDTDPath As String
  Dim sExtPath As String
  
  Dim bolUP As Boolean
  
  sBaseKey = "Software\DaisyWare\Validator"
  'fncLoadRegistryData "Elcel\ExePath", sEExePath, , , propElcelExePath
  fncLoadRegistryData "Misc\TempPath", sTempPath, HKEY_CURRENT_USER, , sAppPath
  fncLoadRegistryData "Misc\Dtd_AdtdPath", sDTDPath, HKEY_CURRENT_USER, , sAppPath & "externals\"
  fncLoadRegistryData "Misc\DefRepPath", sDefaultReportPath, HKEY_CURRENT_USER, , _
    sAppPath & "reports\"
  
  fncLoadRegistryData "Settings\TimeFluctuation", lTimeFluct, HKEY_CURRENT_USER, , 0
  fncSetTimeSpan lTimeFluct
  
  fncLoadRegistryData "Settings\Show ADTD advanced information", bolADTDAdvanced, HKEY_CURRENT_USER, , _
    False
  
  'fncSetElcelPaths sEExePath, sDTDPath, sTempPath
  fncSetDTDPath sDTDPath
  fncSetAdtdPath sDTDPath
  fncSetVtmPath sAppPath
  
  fncSetAdvancedADTD bolADTDAdvanced
End Function

Public Function fncSaveToRegistry()
  sBaseKey = "Software\DaisyWare\validator"
  
  fncSaveRegistryData "Misc\TempPath", propTempPath, HKEY_CURRENT_USER
  fncSaveRegistryData "Misc\Dtd_AdtdPath", propDTDPath, HKEY_CURRENT_USER
  fncSaveRegistryData "Misc\DefRepPath", sDefaultReportPath, HKEY_CURRENT_USER
  fncSaveRegistryData "Apperance\Windowstate", CLng(frmMain.WindowState), HKEY_CURRENT_USER
  fncSaveRegistryData "Settings\TimeFluctuation", lTimeFluct, HKEY_CURRENT_USER
  fncSaveRegistryData "Misc\AppPath", sAppPath, HKEY_CURRENT_USER
  
End Function

Public Function RunQueue(ieRunThese As enuQueueMode)
Dim i As Long, mItem As ListItem
    frmMain.subSetInterfaceStatus False
    
    Select Case ieRunThese
        Case runall 'run all in array
            For i = 0 To (lCandidatesAdded - 1)
                lCurrentCandidate = i + 1 'used for statusbar: adctive candidate in queue
                If Not RunValidation(aCandidateQueue(i)) Then Exit For
                aCandidateQueue(i).bolSelected = False
            Next i
        Case runselected 'run currently highlighted item in mainform treeview
            Dim lCounter As Long
            lCounter = fncGetSelectedCandidate
            If lCounter < 0 Then GoTo Skip
            If Not RunValidation(aCandidateQueue(lCounter)) Then GoTo Skip
        Case runchecked 'run only those with checkbox checked in frmCandidatelist
            For i = 0 To (lCandidatesAdded - 1)
                If aCandidateQueue(i).bolSelected Then
                    lCurrentCandidate = i + 1 'used for statusbar: active candidate in queue
                    If Not RunValidation(aCandidateQueue(i)) Then Exit For
                End If
                aCandidateQueue(i).bolSelected = False
            Next i
    End Select
    
    fncUpdateCandidateList
    Beep
Skip:
    frmMain.subSetInterfaceStatus True
End Function

Public Function RunValidation(CandidateInfo As tCandidateInfo) As Boolean
Dim objValidate As Object

  RunValidation = False
  SetEventCounters
  sCurrentState = "validating"
  
  On Error GoTo errH
  
  Select Case CandidateInfo.lCandidateType
        Case 0 'TYPE_SINGLEDTB
            Set objValidate = New oValidateDTB
            If Not (objValidate.fncValidate(CandidateInfo.sAbsPath)) Then GoTo errH
        Case 1 'TYPE_MULTIVOLUME
            Set objValidate = New oValidateMultivolume
            Dim sAbsPath() As String, lAbsPathCount As Long, lValue1 As Long
            Dim lValue2 As Long
            
            lValue1 = 1
            lValue2 = InStr(1, CandidateInfo.sAbsPath, ", ", vbBinaryCompare)
            Do Until lValue2 = 0
              ReDim Preserve sAbsPath(lAbsPathCount)
              sAbsPath(lAbsPathCount) = Mid$(CandidateInfo.sAbsPath, lValue1, lValue2 - lValue1)
              
              lValue1 = lValue2 + 2
              lValue2 = InStr(lValue1, CandidateInfo.sAbsPath, ", ", vbBinaryCompare)
              lAbsPathCount = lAbsPathCount + 1
            Loop
            
            If Not (objValidate.fncValidate(sAbsPath)) Then GoTo errH
        Case 2 'TYPE_SINGLE_NCC
            Set objValidate = New oValidateNcc
            If Not (objValidate.fncValidate(CandidateInfo.sAbsPath)) Then GoTo errH
        Case 3 'TYPE_SINGLE_SMIL
            Set objValidate = New oValidateSmil
            If Not (objValidate.fncValidate(CandidateInfo.sAbsPath)) Then GoTo errH
        Case 4 'TYPE_SINGLE_MSMIL
            Set objValidate = New oValidateMasterSmil
            If Not (objValidate.fncValidate(CandidateInfo.sAbsPath)) Then GoTo errH
        Case 5 'TYPE_SINGLE_CONTENTDOC
            Set objValidate = New oValidateContent
            If Not (objValidate.fncValidate(CandidateInfo.sAbsPath)) Then GoTo errH
        Case 6 'TYPE_SINGLE_DISCINFO
            Set objValidate = New oValidateDiscinfo
            If Not (objValidate.fncValidate(CandidateInfo.sAbsPath)) Then GoTo errH
'        Case 7 'TYPE_SINGLEDTB light
'            Set objValidate = New ozValidateDTBLight
'            If Not (objValidate.fncValidate(CandidateInfo.sAbsPath)) Then GoTo Errh
  End Select
  
'  Dim oRA As oReport, repIt As oReportItem, lCounter As Long
'  For lCounter = 0 To objValidate.objReport.lCount
'    objValidate.objReport.fncRetrieveFailedTestItem 10, repIt
'    MsgBox (repIt.sTestId)
'  Next lCounter
  
  If fncGetCancelFlag Then fncClearCancelFlag
  
  RunValidation = True
  
errH:
  If Not RunValidation Then MsgBox ("Internal error while validating!")
  sCurrentState = "idle"
  Set CandidateInfo.objReport = objValidate.objReport
  Set objValidate = Nothing
  fncDeinitializeValidator
  
  'save report
  
  'add report to treeview
  fncAddCandidateToTree CandidateInfo
  frmMain.tmUiUpdate_Timer
End Function

Private Sub SetEventCounters()
    lTestsSucceeded = 0
    lTestsFailed = 0
    lErrorLogs = 0
End Sub

Public Function fncAddCandidateToArray(ieCandidateType As Long, _
  isAbsPath As String, Optional ibolNoUpdate As Variant) As Boolean
        
        On Error GoTo errH
        
        Dim bolNoUpdate As Boolean
        bolNoUpdate = False
        If Not IsMissing(ibolNoUpdate) Then bolNoUpdate = ibolNoUpdate
        
        fncAddCandidateToArray = False
        
            ReDim Preserve aCandidateQueue(lCandidatesAdded)
            aCandidateQueue(lCandidatesAdded).lCandidateType = ieCandidateType
            aCandidateQueue(lCandidatesAdded).sAbsPath = isAbsPath
            aCandidateQueue(lCandidatesAdded).bolSelected = True
            Set aCandidateQueue(lCandidatesAdded).objReport = Nothing
            lCandidatesAdded = lCandidatesAdded + 1
            fncAddCandidateToArray = True

errH:
        If (Not fncAddCandidateToArray) Then
            MsgBox ("error adding candidate")
            frmMain.objDllEvent_evErrorLog ("error adding " & isAbsPath & "to candidate list")
        End If
        
        If (Not bolNoUpdate) Then fncUpdateCandidateList
End Function

Public Function fncRemoveCandidateFromArray(lItem As Long) As Boolean
Dim i As Long
'    On Error GoTo Errh
    fncRemoveCandidateFromArray = False

    If Not lItem > lCandidatesAdded Then
        
        For i = lItem To lCandidatesAdded - 2
          aCandidateQueue(i).lCandidateType = aCandidateQueue(i + 1).lCandidateType
          aCandidateQueue(i).sAbsPath = aCandidateQueue(i + 1).sAbsPath
          'Set aCandidateQueue(i + 1).objReport = Nothing
          Set aCandidateQueue(i).objReport = aCandidateQueue(i + 1).objReport
          Set aCandidateQueue(i + 1).objReport = Nothing
        Next i
        
        With aCandidateQueue(lCandidatesAdded - 1)
          Set .objReport = Nothing
          .lCandidateType = -1
          .sAbsPath = ""
        End With
        
        lCandidatesAdded = lCandidatesAdded - 1
        If lCandidatesAdded > 0 Then
          ReDim Preserve aCandidateQueue(lCandidatesAdded - 1)
        End If
    End If
    
    
    fncRemoveCandidateFromArray = True
    
errH:
    If Not fncRemoveCandidateFromArray Then
      MsgBox ("error removing candidate" & lItem)
      frmMain.objDllEvent_evErrorLog ("error removing candidate " & lItem)
    End If
    
    If lCandidatesAdded = 0 Then
      Set aCandidateQueue(0).objReport = Nothing
    End If
    
    fncUpdateCandidateList
End Function

Public Sub AddSingleContentDoc()
  Dim sTemp As String, aPaths() As String, i As Long
    
    sTemp = fncOpenFile(TYPE_SINGLE_CONTENTDOC)
    
    If sTemp <> "" Then
        aPaths = Split(sTemp, Chr(0))
        If UBound(aPaths) = 0 Then 'if only one file was selected
            fncAddCandidateToArray TYPE_SINGLE_CONTENTDOC, aPaths(0)
        Else 'multiselect
            For i = 1 To UBound(aPaths)
                fncAddCandidateToArray TYPE_SINGLE_CONTENTDOC, aPaths(0) & "\" & _
                  aPaths(i)
            Next i
        End If
    End If

End Sub
Public Sub AddSingleDiscInfo()
    Dim sTemp As String
    sTemp = fncOpenFile(TYPE_SINGLE_DISCINFO)
    If sTemp <> "" Then fncAddCandidateToArray TYPE_SINGLE_DISCINFO, sTemp
End Sub

Public Sub AddSingleMasterSmil()
    Dim sTemp As String
    sTemp = fncOpenFile(TYPE_SINGLE_MSMIL)
    If sTemp <> "" Then fncAddCandidateToArray TYPE_SINGLE_MSMIL, sTemp
End Sub

Public Sub AddSingleSmil()
  Dim sTemp As String, aPaths() As String, i As Long
    Dim sBasePath As String
    sTemp = fncOpenFile(TYPE_SINGLE_SMIL)
    
    If sTemp <> "" Then
        aPaths = Split(sTemp, Chr(0))
        If UBound(aPaths) = 0 Then 'if only one file was selected
            fncAddCandidateToArray TYPE_SINGLE_SMIL, aPaths(0)
        Else 'multiselect
            For i = 1 To UBound(aPaths)
                fncAddCandidateToArray TYPE_SINGLE_SMIL, aPaths(0) & "\" & _
                  aPaths(i)
            Next i
        End If
    End If

End Sub

Public Sub AddSingleDtb()
    Dim sTemp As String
    sTemp = fncGetPathName(fncOpenFile(TYPE_SINGLEDTB))
    If sTemp <> "" Then fncAddCandidateToArray TYPE_SINGLEDTB, sTemp
End Sub

'Public Sub AddSingleDtbLight()
'    Dim sTemp As String
'    sTemp = fncGetPathName(fncOpenFile(TYPE_SINGLEDTB))
'    If sTemp <> "" Then fncAddCandidateToArray TYPE_TYPE_SINGLEDTB_LIGHT, sTemp
'End Sub

Public Sub AddMultiVolume()
    Dim sTemp As String, sVolumes As String
    Do
      sTemp = fncOpenFile(TYPE_MULTIVOLUME)
      If Not sTemp = "" Then sVolumes = sVolumes & fncGetPathName(sTemp) & ", "
    Loop Until sTemp = ""
    
    If sVolumes = "" Then Exit Sub
    fncAddCandidateToArray TYPE_MULTIVOLUME, sVolumes
End Sub


Public Sub AddSingleNcc()
Dim sTemp As String
    sTemp = fncOpenFile(TYPE_SINGLE_NCC)
    If sTemp <> "" Then fncAddCandidateToArray TYPE_SINGLE_NCC, sTemp
End Sub

Public Function doRtfSetSelection(ByVal lngSelStart As Long, ByVal lngSelLength As Long) As Boolean
    doRtfSetSelection = False
    With frmMain.rtfEdit
        If lngSelStart < Len(.Text) And lngSelLength > 0 Then
            .SelStart = lngSelStart - 1
            .SelLength = lngSelLength
        Else
            Exit Function
        End If
    End With
    doRtfSetSelection = True
End Function

Public Sub subFollowLink()
  Dim sSearchstring As String, lFound1 As Long, lFound2 As Long
  Dim sD As String, sP As String, sF As String, sFile2Open As String
  
  If frmMain.rtfEdit.SelLength > 0 Then
    sSearchstring = Mid$(frmMain.rtfEdit.Text, frmMain.rtfEdit.SelStart, _
      frmMain.rtfEdit.SelLength + 1)
  ElseIf frmMain.rtfEdit.SelStart > 0 Then
    sSearchstring = Mid$(frmMain.rtfEdit.Text, frmMain.rtfEdit.SelStart, _
      Len(frmMain.rtfEdit) - frmMain.rtfEdit.SelStart)
  Else
    Exit Sub
  End If
  
  lFound1 = InStr(1, sSearchstring, "href", vbTextCompare)
  lFound2 = InStr(1, sSearchstring, "src", vbTextCompare)
  
  If (lFound1 = 0) Then lFound1 = lFound2
  If (lFound1 > lFound2) And (lFound2 > 0) Then lFound1 = lFound2
  If lFound1 < 1 Then Exit Sub
  
  lFound1 = InStr(lFound1, sSearchstring, Chr(34), vbBinaryCompare)
  lFound2 = InStr(lFound1 + 1, sSearchstring, Chr(34), vbBinaryCompare)
  
  If lFound1 < 1 Or lFound1 > Len(sSearchstring) Or lFound2 < 1 Or _
    lFound2 > Len(sSearchstring) Then Exit Sub
    
  sFile2Open = Mid$(sSearchstring, lFound1 + 1, lFound2 - lFound1 - 1)
  fncShowFile sFile2Open, False, True
End Sub

Public Function fncShowFile( _
  isAbsPath As String, bolNew As Boolean, bolForth As Boolean)
  
  'mg 20050330:
  If (InStr(1, LCase$(isAbsPath), ".mp3", vbTextCompare) > 0) Or _
    InStr(1, LCase$(isAbsPath), ".wav", vbTextCompare) > 0 Then Exit Function
  
  If bolNew Then
    lCurrentHistory = 1
    aFileHistory(lCurrentHistory).sAbsPath = isAbsPath
    aFileHistory(lCurrentHistory).lBytePos = frmMain.rtfEdit.SelStart
    aFileHistory(lCurrentHistory).sID = ""
  ElseIf bolForth Then
    aFileHistory(lCurrentHistory).lBytePos = frmMain.rtfEdit.SelStart
    lCurrentHistory = lCurrentHistory + 1
    
    If lCurrentHistory = 11 Then
      For lCurrentHistory = 1 To 9
        aFileHistory(lCurrentHistory) = aFileHistory(lCurrentHistory + 1)
      Next lCurrentHistory
      lCurrentHistory = 10
    End If
    
    Dim sD As String, sP As String, sF As String, sID As String
    fncParseURI isAbsPath, sD, sP, sF, sID, _
      aFileHistory(lCurrentHistory - 1).sAbsPath
    
    aFileHistory(lCurrentHistory).sAbsPath = sD & sP & sF
    aFileHistory(lCurrentHistory).sID = sID
  Else
    If lCurrentHistory < 2 Then Exit Function
    lCurrentHistory = lCurrentHistory - 1
  End If
  
  On Error GoTo Skip
  
  Dim oFSO As Object, oFile As Object
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFile = oFSO.opentextfile(aFileHistory(lCurrentHistory).sAbsPath)
  frmMain.rtfEdit.Text = oFile.readall
  
  If sID = "" Then
    frmMain.rtfEdit.SelStart = aFileHistory(lCurrentHistory).lBytePos
  Else
    Dim lCounter As Long, lIteration As Long, lCounter2 As Long
    Do
      lCounter = InStr(lCounter + 1, frmMain.rtfEdit.Text, _
        aFileHistory(lCurrentHistory).sID)
      
      If lCounter = 0 Then Exit Do
      
      lIteration = 0
      lCounter2 = lCounter
      Do
        lCounter2 = lCounter2 - 1
        Select Case Mid$(frmMain.rtfEdit.Text, lCounter2, 1)
          Case "="
            If lIteration = 0 Then lIteration = 1 Else Exit Do
          Case " ", vbCr, vbLf, Chr(34), Chr(39)
            
          Case "d", "D"
            If lIteration = 1 Then lIteration = 2 Else Exit Do
          
          Case "i", "I"
            If lIteration = 2 Then
              frmMain.rtfEdit.SelStart = lCounter2 - 1
              GoTo Skip
            Else
              Exit Do
            End If
          Case Else
            Exit Do
        End Select
      Loop Until lCounter2 = 0
    Loop
Skip:
  End If
  Err.Clear
  
  frmMain.rtfEdit.SetFocus
  If Not oFile Is Nothing Then oFile.Close
End Function

Public Function fncClearFileHistory()
  lCurrentHistory = 0
  frmMain.rtfEdit = ""
  frmMain.tmUiUpdate_Timer
  frmMain.txErrView = ""
End Function

Public Function fncAddCandidateToTree(itypCand As tCandidateInfo)

  Dim objReportItem As oReportItem, lCounter1 As Long, objParent As Node
  Dim objBaseNode As Node, objErrNode As Node, objWarNode As Node
  Dim objParNode As Node, lImage As Long
  
'  Dim objErrCritical As Node, objErrNonCritical As Node
  
  For lCounter1 = 1 To frmMain.treeErrorView.Nodes.Count
    If frmMain.treeErrorView.Nodes.Item(lCounter1).Key = itypCand.sAbsPath Then
      frmMain.treeErrorView.Nodes.Remove (lCounter1)
      Exit For
    End If
  Next lCounter1
      
  If itypCand.objReport Is Nothing Then
    Set objBaseNode = frmMain.treeErrorView.Nodes.Add(, , itypCand.sAbsPath, _
      itypCand.sAbsPath, itypCand.lCandidateType + 11)
    objBaseNode.Sorted = True
    
    frmMain.treeErrorView.Nodes.Add objBaseNode, tvwChild, _
      itypCand.sAbsPath & " not validated", "not validated"
  ElseIf itypCand.objReport.lFailedTestCount = 0 Then
    Set objBaseNode = frmMain.treeErrorView.Nodes.Add(, , itypCand.sAbsPath, _
      itypCand.sAbsPath, itypCand.lCandidateType + 3)
    objBaseNode.Sorted = True
    
    frmMain.treeErrorView.Nodes.Add objBaseNode, tvwChild, itypCand.sAbsPath & "noerr", "no errors", _
      18
  Else
    Set objBaseNode = frmMain.treeErrorView.Nodes.Add(, , itypCand.sAbsPath, _
      itypCand.sAbsPath, itypCand.lCandidateType + 3)
    objBaseNode.Sorted = True
    
    For lCounter1 = 0 To itypCand.objReport.lFailedTestCount - 1
      itypCand.objReport.fncRetrieveFailedTestItem lCounter1, objReportItem
    
      If LCase$(objReportItem.sFailType) = "warning" Then
        If objWarNode Is Nothing Then
          Set objWarNode = frmMain.treeErrorView.Nodes.Add(objBaseNode, _
            tvwChild, itypCand.sAbsPath & " warning", "warnings", 2)
        End If
        
        Set objParNode = objWarNode: lImage = 2
      Else
        
        
        If objErrNode Is Nothing Then
          Set objErrNode = frmMain.treeErrorView.Nodes.Add(objBaseNode, _
            tvwChild, itypCand.sAbsPath & " error", "errors", 1)
        End If

        Set objParNode = objErrNode: lImage = 1
        
'        If LCase$(objReportItem.sFailClass) = "critical" Then
'          If objErrCritical Is Nothing Then
'            Set objErrCritical = frmMain.treeErrorView.Nodes.Add(objBaseNode, _
'              tvwChild, itypCand.sAbsPath & " error [critical]", _
'              "errors [critical]", 1)
'          End If
'          Set objParNode = objErrCritical: lImage = 1
'        Else
'          If objErrNonCritical Is Nothing Then
'            Set objErrNonCritical = frmMain.treeErrorView.Nodes.Add(objBaseNode, _
'              tvwChild, itypCand.sAbsPath & " error [non-critical]", _
'              "errors [non-critical]", 1)
'          End If
'          Set objParNode = objErrNonCritical: lImage = 1
'        End If
      End If
    
      frmMain.treeErrorView.Nodes.Add objParNode, tvwChild, itypCand.sAbsPath & lCounter1, _
        objReportItem.sShortDesc & " @ " & fncGetFileName(objReportItem.sAbsPath) & _
        " [" & objReportItem.lLine & ":" & objReportItem.lColumn & "]", lImage
    Next lCounter1
  End If
End Function

Public Function fncUpdateCandidateList()
Dim colItem As ColumnHeader
Dim mItem As ListItem
Dim i As Long, aChecked() As CheckBoxConstants, lCheckCount As Long
    
    With frmCandidateList.lstCandidates
       .View = lvwReport
       .ColumnHeaders.Clear
       .ListItems.Clear
       .LabelEdit = lvwManual
       
        Set colItem = .ColumnHeaders.Add()
            colItem.Text = "Candidate"
            colItem.Width = .Width * 0.8
        Set colItem = .ColumnHeaders.Add()
            colItem.Text = "Type"
            colItem.Width = .Width * 0.2

    Dim objNode As Node
    
    Do Until frmMain.treeErrorView.Nodes.Count = 0
      lCheckCount = lCheckCount + 1
      frmMain.treeErrorView.Nodes.Remove (1)
    Loop
    
    For i = 0 To (lCandidatesAdded - 1)
        fncAddCandidateToTree aCandidateQueue(i)
            
        Set mItem = .ListItems.Add()
        mItem.Text = aCandidateQueue(i).sAbsPath
        'mItem.Checked = aCandidateQueue(i).bolChecked
        If aCandidateQueue(i).bolChecked Then mItem.Checked = True Else _
          mItem.Checked = False
        Select Case aCandidateQueue(i).lCandidateType
            Case 0
                mItem.SubItems(1) = "single dtb"
            Case 1
                mItem.SubItems(1) = "multi volume"
            Case 2
                mItem.SubItems(1) = "ncc"
            Case 3
                mItem.SubItems(1) = "smil"
            Case 4
                mItem.SubItems(1) = "master smil"
            Case 5
                mItem.SubItems(1) = "content"
            Case 6
                mItem.SubItems(1) = "discinfo"
        End Select
    Next i
    
    End With
End Function

Public Function fncGetSelectedCandidate() As Long
  fncGetSelectedCandidate = -1
  
  If frmMain.treeErrorView.SelectedItem Is Nothing Then Exit Function
  
  Dim lCounter As Long, lCounter2 As Long, oNode As Node
  
  Set oNode = frmMain.treeErrorView.SelectedItem
  Do Until oNode.Parent Is Nothing
    Set oNode = oNode.Parent
  Loop
          
  For lCounter = 0 To lCandidatesAdded - 1
    If oNode.Text = _
      aCandidateQueue(lCounter).sAbsPath Then
      Exit For
    End If
  Next lCounter
  
  If lCounter > lCandidatesAdded - 1 Then Exit Function
  
  fncGetSelectedCandidate = lCounter
End Function

Public Function fncGetSelectedCanRepItem() As oReportItem
'  fncGetSelectedCanRepItem = -1
  
  If frmMain.treeErrorView.Parent Is Nothing Then Exit Function
  
  Dim lCounter As Long, lCounter2 As Long
  lCounter = fncGetSelectedCandidate
  If lCounter = -1 Then Exit Function
  
  Dim oNode As Node
  Set oNode = frmMain.treeErrorView.SelectedItem
  If Len(oNode.Key) <= Len(aCandidateQueue(lCounter).sAbsPath) Then Exit Function
  
  On Error GoTo ErrorH
  lCounter2 = CLng(Right$(oNode.Key, Len(oNode.Key) - _
    Len(aCandidateQueue(lCounter).sAbsPath)))
    
  If (lCounter2 < 0) Or (lCounter2 > _
    aCandidateQueue(lCounter).objReport.lFailedTestCount) Then Exit Function
  
  aCandidateQueue(lCounter).objReport.fncRetrieveFailedTestItem _
    lCounter2, fncGetSelectedCanRepItem
ErrorH:
End Function
