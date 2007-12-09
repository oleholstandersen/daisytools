Attribute VB_Name = "IteratorMdlGlobal"
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

Public sAppPath As String
Public sMotherDir As String
Public sReportDir As String
Public sNccPathArray() As String
Public sNccPathArrayItems As Long
Public bolCancel As Boolean
Public bolLightMode As Boolean
Public bolDisableAudioTests As Boolean
Public lTimeFluct As Long
Public bolMoveDtbAfter As Boolean
Public sDtbMoveDestination As String
Public oFSO As Object
Public bolValidating As Boolean
Public lNumberOfFailedBooks As Long, lNumberOfSkippedbooks As Long
Public sStartTime As String, sEndTime As String

Public Type tCandidateInfo
    sAbsPath As String
    lCandidateType As Long
    objReport As oReport
    bolSelected As Boolean
    bolChecked As Boolean
End Type

Public Candidate As tCandidateInfo

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
   ByVal lpParameters As String, ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long


Sub Main()
  Dim oTemp As Object
  Dim bMsxmlPresenceIsTested As Boolean
  
  On Error GoTo errH
  
  bMsxmlPresenceIsTested = True
  Set oTemp = CreateObject("Msxml2.DOMDocument.4.0")
  If oTemp Is Nothing Then
    MsgBox "MSXML 4 not installed, exiting program.", vbOKOnly
    Exit Sub
  End If
  bMsxmlPresenceIsTested = False

  sAppPath = App.Path: If Not Right$(sAppPath, 1) = "\" Then sAppPath = sAppPath & "\"
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  fncLoadFromRegistry
  subApplyUserLcid
  IteratorFrmMain.Show
  IteratorFrmMain.fncUpdateUi
  IteratorFrmMain.SetFocus
  
    Exit Sub
errH:
  If bMsxmlPresenceIsTested Then
    MsgBox "MSXML 4 is not installed. Exiting program.", vbOKOnly
  Else
    MsgBox "An unknown error occured. Exiting program.", vbOKOnly
  End If

End Sub

Public Function fncRunQueue() As Boolean
Dim i As Long, mItem As ListItem
Dim objValidate As Object, sNccFolder As String
Dim bolIsValid As Boolean
Dim sBatchReport As String
  
  'check that reportDir exists
  If Not fncFolderExists(sReportDir) Then
    sReportDir = sAppPath & "iterator_reports\"
    If Not fncFolderExists(sReportDir) Then
      fncCreateDirectoryChain (sReportDir)
    End If
  End If

  sStartTime = Format(Date, "yyyy-mm-dd") & " " & Time
  IteratorFrmMain.txtStatus.Text = IteratorFrmMain.txtStatus.Text & vbCrLf & "start: " & sStartTime

  If sNccPathArrayItems > 0 Then
  For i = 0 To sNccPathArrayItems - 1
    If IteratorFrmMain.lstCandidates.ListItems.Item(i + 1).Checked Then
      bolValidating = True
      IteratorFrmMain.fncUpdateUi
      Set objValidate = New oValidateDTB
      fncSetLightMode bolLightMode
      fncSetDisableAudioTests bolDisableAudioTests
      fncSetTimeSpan lTimeFluct
      IteratorFrmMain.lstCandidates.ListItems.Item(i + 1).Bold = True
      sNccFolder = oFSO.GetParentFolderName(sNccPathArray(i))
      If Not Right$(sNccFolder, 1) = "\" Then sNccFolder = sNccFolder & "\"
      If Not (objValidate.fncValidate(sNccFolder)) Then GoTo errH
      Candidate.sAbsPath = sNccPathArray(i)
      Candidate.lCandidateType = 0
      Set Candidate.objReport = objValidate.objReport
      With IteratorFrmMain.lstCandidates.ListItems.Item(i + 1)
        If Candidate.objReport.lFailedTestCount = 0 Then
          .SubItems(1) = "pass"
          'mg: send fail or pass into save report in order to sort in subfolders
          bolIsValid = True
        Else
          .SubItems(1) = "fail"
          lNumberOfFailedBooks = lNumberOfFailedBooks + 1
          bolIsValid = False
        End If
      End With
                  
      If Not fncSaveReport(Candidate, bolIsValid) Then GoTo errH
      
'      If bolMoveDtbAfter Then
'          Dim sSub As String: If Candidate.objReport.lFailedTestCount = 0 Then sSub = "pass\" Else sSub = "fail"
'          Dim sTmp As String: sTmp = sDtbMoveDestination & sSub
'          'karl: här
'         'oFSO.movefolder sNccFolder & "*", sTmp
'      End If


errH:
      fncDeinitializeValidator
      Set objValidate = Nothing
      If bolCancel Then Exit For
    Else
     lNumberOfSkippedbooks = lNumberOfSkippedbooks + 1
     IteratorFrmMain.lstCandidates.ListItems.Item(i + 1).SubItems(1) = "skipped"
    End If 'if checked
  Next i
  
    bolValidating = False
    sEndTime = Format(Date, "yyyy-mm-dd") & " " & Time
    
    IteratorFrmMain.txtStatus.Text = IteratorFrmMain.txtStatus.Text & vbCrLf & lNumberOfFailedBooks & " out of " & sNccPathArrayItems & " books were invalid."
    If lNumberOfSkippedbooks > 0 Then
      IteratorFrmMain.txtStatus.Text = IteratorFrmMain.txtStatus.Text & vbCrLf & lNumberOfSkippedbooks & " out of " & sNccPathArrayItems & " books were skipped."
    End If
    IteratorFrmMain.txtStatus.Text = IteratorFrmMain.txtStatus.Text & vbCrLf & "end: " & sEndTime
    IteratorFrmMain.txtStatus.Text = IteratorFrmMain.txtStatus.Text & vbCrLf & "Done."
        
    'save a batch report to file
    sBatchReport = "[validator iterator batch report]" & vbCrLf
    sBatchReport = sBatchReport & "Started " & sStartTime & vbCrLf
    sBatchReport = sBatchReport & lNumberOfFailedBooks & " out of " & sNccPathArrayItems & " books were invalid." & vbCrLf
    If lNumberOfSkippedbooks > 0 Then
      sBatchReport = sBatchReport & lNumberOfSkippedbooks & " out of " & sNccPathArrayItems & " books were skipped." & vbCrLf
    End If
    sBatchReport = sBatchReport & "Finished " & sEndTime & vbCrLf
    
    fncSaveFile sReportDir & "batchreport_" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".txt", sBatchReport
    sBatchReport = ""
    
    IteratorFrmMain.fncUpdateUi
    Beep
  End If 'sNccPathArrayItems > 0 Then
End Function

Private Function fncLoadFromRegistry()
  Dim sTempPath As String, sDTDPath As String
   
  sBaseKey = "Software\DaisyWare\Validator"
  fncLoadRegistryData "Misc\TempPath", sTempPath, , , sAppPath
  fncLoadRegistryData "Misc\Dtd_AdtdPath", sDTDPath, , , sAppPath & "externals\"
  fncLoadRegistryData "Misc\IteratorMotherPath", sMotherDir, , , ""
  fncLoadRegistryData "Misc\IteratorReportPath", sReportDir, , , ""
  fncLoadRegistryData "Misc\IteratorLightMode", bolLightMode, , , False
  fncLoadRegistryData "Misc\IteratorDisableAudioTests", bolDisableAudioTests, , , False
  fncLoadRegistryData "Misc\IteratorTimeFluct", lTimeFluct, , , 1000
'  fncLoadRegistryData "Misc\IteratorMoveDtbAfter", bolMoveDtbAfter, , , False
'  fncLoadRegistryData "Misc\IteratorMoveDestination", sDtbMoveDestination, , , ""
      
  fncSetDTDPath sDTDPath
  fncSetAdtdPath sDTDPath
  fncSetVtmPath sAppPath
  fncSetAdvancedADTD False
End Function

Public Function fncSaveToRegistry()
  sBaseKey = "Software\DaisyWare\validator"
  fncSaveRegistryData "Misc\IteratorMotherPath", sMotherDir
  fncSaveRegistryData "Misc\IteratorReportPath", sReportDir
  fncSaveRegistryData "Misc\IteratorLightMode", bolLightMode
  fncSaveRegistryData "Misc\IteratorDisableAudioTests", bolDisableAudioTests
  fncSaveRegistryData "Misc\IteratorTimeFluct", lTimeFluct
'  fncSaveRegistryData "Misc\IteratorMoveDtbAfter", bolMoveDtbAfter
'  fncSaveRegistryData "Misc\IteratorMoveDestination", sDtbMoveDestination
End Function

Public Function fncFindFiles(sMotherDir As String, sFileName As String) As Boolean
Dim oFile As Object, oFiles As Object, oFolder As Object, oSubFolder As Object, oFolders As Object

  On Error GoTo Errhandler
  fncFindFiles = False
  If bolCancel Then fncFindFiles = True: Exit Function
  Set oFolder = oFSO.GetFolder(sMotherDir)
  If Not oFolder Is Nothing Then
    Set oFiles = oFolder.Files
    For Each oFile In oFiles
      If oFile.Name = "ncc.html" Then
        ReDim Preserve sNccPathArray(sNccPathArrayItems)
        sNccPathArray(sNccPathArrayItems) = oFile.Path
        sNccPathArrayItems = sNccPathArrayItems + 1
        GoTo found
      End If
    Next
found:
    Set oFolders = oFolder.subfolders
    For Each oSubFolder In oFolders
      If Not fncFindFiles(oSubFolder.Path, "ncc.html") Then GoTo Errhandler
    Next
  Else
    GoTo Errhandler
  End If 'Not oFolder Is Nothing
  IteratorFrmMain.fncUpdateCandidateList
  
  IteratorFrmMain.txtStatus.Text = "Found " & CStr(sNccPathArrayItems) & " books."
  fncFindFiles = True
  
Errhandler:
End Function

Public Function fncRemoveCandidateFromArray(lItem As Long) As Boolean
Dim i As Long
    fncRemoveCandidateFromArray = False
    If Not lItem > sNccPathArrayItems Then
      For i = lItem To sNccPathArrayItems - 2
          sNccPathArray(i) = sNccPathArray(i + 1)
      Next i
      sNccPathArray(sNccPathArrayItems - 1) = ""
      
    End If
    sNccPathArrayItems = sNccPathArrayItems - 1
    If sNccPathArrayItems > 0 Then
      ReDim Preserve sNccPathArray(sNccPathArrayItems)
    End If
    fncRemoveCandidateFromArray = True
End Function

Public Function fncReadFile(sAbsPath As String) As String
Dim oFSO As Object
Dim oStream As Object
    On Error GoTo Errhandler
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.opentextfile(sAbsPath)
    fncReadFile = oStream.ReadAll
Errhandler:
    oStream.Close
    Set oFSO = Nothing
End Function

Public Function fncTruncToValidUriChars(ByVal sString As String) As String

    Dim i As Long
    Dim sErr As String

    'do some pretty from windows-1252:
    ' space to underscore
    sString = Replace$(sString, Chr(32), Chr(95), , , vbBinaryCompare)
    ' german double s to ss
    sString = Replace$(sString, Chr(223), Chr(115) & Chr(115), , , vbTextCompare)
    '224-227 to a
    sString = Replace$(sString, Chr(224), Chr(97), , , vbTextCompare)
    sString = Replace$(sString, Chr(225), Chr(97), , , vbTextCompare)
    sString = Replace$(sString, Chr(226), Chr(97), , , vbTextCompare)
    sString = Replace$(sString, Chr(227), Chr(97), , , vbTextCompare)
    ' ä/Ä to ae
    sString = Replace$(sString, Chr(228), Chr(97) & Chr(101), , , vbTextCompare)
    ' å/Å to aa
    'sString = Replace$(sString, Chr(229), Chr(97) & Chr(97), , , vbTextCompare)
    '230: danish ae to ae
    sString = Replace$(sString, Chr(230), Chr(97) & Chr(101), , , vbTextCompare)
    '231: ccedil to c
    sString = Replace$(sString, Chr(231), Chr(99), , , vbTextCompare)
    '232-235 to e
    sString = Replace$(sString, Chr(232), Chr(101), , , vbTextCompare)
    sString = Replace$(sString, Chr(233), Chr(101), , , vbTextCompare)
    sString = Replace$(sString, Chr(234), Chr(101), , , vbTextCompare)
    sString = Replace$(sString, Chr(235), Chr(101), , , vbTextCompare)
    '236-239 to i
    sString = Replace$(sString, Chr(236), Chr(105), , , vbTextCompare)
    sString = Replace$(sString, Chr(237), Chr(105), , , vbTextCompare)
    sString = Replace$(sString, Chr(238), Chr(105), , , vbTextCompare)
    sString = Replace$(sString, Chr(239), Chr(105), , , vbTextCompare)
    '240 (icelandic d) to d
    sString = Replace$(sString, Chr(240), Chr(100), , , vbTextCompare)
    '241(ana) to n
    sString = Replace$(sString, Chr(241), Chr(110), , , vbTextCompare)
    '242-245 to o
    sString = Replace$(sString, Chr(242), Chr(111), , , vbTextCompare)
    sString = Replace$(sString, Chr(243), Chr(111), , , vbTextCompare)
    sString = Replace$(sString, Chr(244), Chr(111), , , vbTextCompare)
    sString = Replace$(sString, Chr(245), Chr(111), , , vbTextCompare)
    ' ö/Ö to oe
    sString = Replace$(sString, Chr(246), Chr(111) & Chr(101), , , vbTextCompare)
    ' danish ö to oe
    sString = Replace$(sString, Chr(248), Chr(111) & Chr(101), , , vbTextCompare)
    '249-252 to u
    sString = Replace$(sString, Chr(249), Chr(117), , , vbTextCompare)
    sString = Replace$(sString, Chr(250), Chr(117), , , vbTextCompare)
    sString = Replace$(sString, Chr(251), Chr(117), , , vbTextCompare)
    sString = Replace$(sString, Chr(252), Chr(117), , , vbTextCompare)
    '253 and 255 to y
    sString = Replace$(sString, Chr(253), Chr(121), , , vbTextCompare)
    sString = Replace$(sString, Chr(255), Chr(121), , , vbTextCompare)
        
    'now loop and do the rest ugly (trunc char)
    sErr = sString

    For i = 1 To Len(sString)
        If Not fncIsValidUriChar(Asc(Mid$(sString, i, 1))) Then
            sErr = Replace$(sErr, Mid$(sString, i, 1), "", , , vbTextCompare)
        End If
    Next i

    If sErr = "" Then sErr = "dtb_"
     
    sString = sErr
                
    fncTruncToValidUriChars = sString

End Function

Private Function fncIsValidUriChar(sChar As String) As Boolean

    fncIsValidUriChar = False
    
    If (CInt(sChar) > 44 And CInt(sChar) < 58) Or _
       (CInt(sChar) > 63 And CInt(sChar) < 91) Or _
       (CInt(sChar) > 94 And CInt(sChar) < 123) _
       Then
        'char is ok
    Else
        Exit Function
    End If
    
    fncIsValidUriChar = True

End Function

