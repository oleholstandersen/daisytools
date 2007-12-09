Attribute VB_Name = "mdlFileHandler"
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

Public Function fncGetParentFolderName(sAbsPath As String) As String
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    fncGetParentFolderName = oFSO.GetParentFolderName(sAbsPath)
End Function

Public Function fncFileExists(sCandidate As String) As Boolean
    Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    fncFileExists = False
    If oFSO.fileExists(sCandidate) Then
        fncFileExists = True
    Else
        
    End If
    Set oFSO = Nothing
End Function

Public Function fncSaveFile(isContent As String, filetype As eFileType, _
  Optional sAbsPath As Variant) As Boolean
  
  Dim fso As Object, stream As Object
  Dim sFileName As String
  
    Set fso = CreateObject("Scripting.FileSystemObject")
    fncSaveFile = False
    
    
    With frmMain
        .CommonDialog1.Flags = cdlOFNOverwritePrompt Or _
                cdlOFNCreatePrompt Or _
                cdlOFNNoChangeDir
        .CommonDialog1.CancelError = True
        
        On Error GoTo Errhandler
    
        .CommonDialog1.DialogTitle = "Save File"
        If Not IsMissing(sAbsPath) Then .CommonDialog1.FileName = sAbsPath
        Select Case filetype
            Case 0      'eventlog
                .CommonDialog1.Filter = "Log files (*.log)|*.log|"
                .CommonDialog1.FileName = "error"
            Case 1      'report
                .CommonDialog1.Filter = "Report file (*.html)|*.html"
                .CommonDialog1.FileName = "report"
            Case 2      'filesetfile
            
            Case Else
                .CommonDialog1.Filter = "Text files (*.txt)|*.txt|"
                .CommonDialog1.FileName = ""
        End Select
Choose:
        If Not filetype = 2 Then
          .CommonDialog1.ShowSave
          sFileName = frmMain.CommonDialog1.FileName
        End If
        '
        If Not InStr(1, .CommonDialog1.FileName, ",", vbBinaryCompare) = 0 Then _
          MsgBox "Cannot have ',' in filename!", vbOKOnly: GoTo Choose
            
        Dim msgbResult As VbMsgBoxResult
            
        Set fso = CreateObject("Scripting.FileSystemObject")
        
'mg 20030910; put bkpfiles in subdir
        If filetype = filesetfile Then
          sFileName = sAbsPath
          If Not fso.folderexists(fncGetPathName(sFileName) & "val_bkp\") Then
             fso.createFolder (fncGetPathName(sFileName) & "val_bkp\")
          End If
          sFileName = fncGetPathName(sFileName) & "val_bkp\" & fncGetFileName(sFileName)
          Do Until Not fso.fileExists(sFileName)
            sFileName = fncGetPathName(sFileName) & _
              fso.getbasename(sFileName) & "_." & fso.getextensionname(sFileName)
          Loop
          fso.copyfile sAbsPath, sFileName
          sFileName = sAbsPath
        End If
                
'mg 20030910; below is orig before abov bkpdir add
'        If filetype = filesetfile Then
'          'make a backup of the preexisting file
'          sFileName = sAbsPath
'          Do Until Not fso.FileExists(sFileName)
'            sFileName = fncGetPathName(sFileName) & _
'              fso.getbasename(sFileName) & "_." & fso.getextensionname(sFileName)
'          Loop
'          fso.copyfile sAbsPath, sFileName
'          sFileName = sAbsPath
'        End If
        
        Set stream = fso.Createtextfile(sFileName, True)
        stream.write (isContent)
        stream.Close
        fncSaveFile = True
    End With
Errhandler:
End Function

Public Function fncOpenFile(lCandidateType As Long) As String
    With frmMain
        
        On Error GoTo Errhandler
        
        Select Case lCandidateType
            Case 0 'TYPE_SINGLEDTB
                SetSingleSelectFlags
                .CommonDialog1.DialogTitle = "Choose ncc"
                .CommonDialog1.Filter = "Navigation Control Center (ncc.html)|ncc.htm*"
                .CommonDialog1.FileName = "ncc.html"
            Case 1 'TYPE_MULTIVOLUME
                SetSingleSelectFlags
                .CommonDialog1.DialogTitle = "Choose ncc"
                .CommonDialog1.Filter = "Navigation Control Center (ncc.html)|ncc.htm*"
                .CommonDialog1.FileName = "ncc.html"
            Case 2 'TYPE_SINGLE_NCC
                SetSingleSelectFlags
                .CommonDialog1.DialogTitle = "Choose ncc"
                .CommonDialog1.Filter = "Navigation Control Center (ncc.html)|ncc.htm*"
                .CommonDialog1.FileName = "ncc.html"
            Case 3 'TYPE_SINGLE_SMIL
                SetMultiSelectFlags
                .CommonDialog1.DialogTitle = "Choose smil"
                .CommonDialog1.Filter = "SMIL file (*.smil)|*.smil"
                .CommonDialog1.FileName = "*.smil"
            Case 4 'TYPE_SINGLE_MSMIL
                SetSingleSelectFlags
                .CommonDialog1.DialogTitle = "Choose master smil"
                .CommonDialog1.Filter = "Master SMIL (master.smil)|master.smil"
                .CommonDialog1.FileName = "master.smil"
            Case 5 'TYPE_SINGLE_CONTENTDOC
                SetMultiSelectFlags
                .CommonDialog1.DialogTitle = "Choose content document (xhtml)"
                .CommonDialog1.Filter = "xhtml content document (*.htm*)|*.htm*"
                .CommonDialog1.FileName = ""
            Case 6 'TYPE_SINGLE_DISCINFO
                SetSingleSelectFlags
                .CommonDialog1.DialogTitle = "Choose discinfo document"
                .CommonDialog1.Filter = "Discinfo document (discinfo.htm*)|discinfo.htm*"
                .CommonDialog1.FileName = "discinfo.htm*"
        End Select
        
Choose:
        .CommonDialog1.ShowOpen
        If Not InStr(1, .CommonDialog1.FileName, ",", vbBinaryCompare) = 0 Then _
          MsgBox "Cannot have ',' in filename!", vbOKOnly: GoTo Choose
        
        fncOpenFile = .CommonDialog1.FileName
Errhandler:
        Exit Function
    End With
End Function

Public Sub SetDirSelectFlags()
'Dim objShell As New Shell, objFolder As Folder, vRoot As Variant
'
'  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose Mother folder of one or several DTBs", _
'    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
'    BIF_VALIDATE)
'
'  If objFolder Is Nothing Then Exit Sub
'  If Not objFolder.Self.IsFileSystem Then Exit Sub
'
'  objTextSaveLogPath.Text = objFolder.Self.Path
'  If Not Left$(objTextSaveLogPath.Text, 1) = "\" Then objTextSaveLogPath.Text = _
'    objTextSaveLogPath.Text & "\"
End Sub

Public Sub SetSingleSelectFlags()
    With frmMain
        .CommonDialog1.Flags = cdlOFNFileMustExist Or _
                                cdlOFNHideReadOnly Or _
                                cdlOFNNoChangeDir
        .CommonDialog1.CancelError = True
    End With
End Sub

Private Sub SetMultiSelectFlags()
    With frmMain
        .CommonDialog1.Flags = cdlOFNFileMustExist Or _
                                cdlOFNHideReadOnly Or _
                                cdlOFNNoChangeDir Or _
                                cdlOFNAllowMultiselect Or _
                                cdlOFNExplorer
        .CommonDialog1.MaxFileSize = 32767
        .CommonDialog1.CancelError = True
    End With
End Sub

Public Function fncFileCount(sAbsPath As String) As Long
Dim oFSO As Object, oFolder As Object, oFiles As Object, oFile As Object
    On Error GoTo Errhandler
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getfolder(oFSO.getabsolutepathname(sAbsPath))
    Set oFiles = oFolder.Files
    fncFileCount = oFiles.Count
Errhandler:
End Function

Public Function fncGetFileName(isAbsPath As String) As String
  Dim sTemp As String, sFileName As String
  If Not fncParseURI(isAbsPath, sTemp, sTemp, sFileName, sTemp) Then Stop
  fncGetFileName = sFileName
End Function

Public Function fncGetPathName(isAbsPath As String) As String
  Dim sTemp As String, sDrive As String, sPath As String
  If Not fncParseURI(isAbsPath, sDrive, sPath, sTemp, sTemp) Then Stop
  fncGetPathName = sDrive & sPath
End Function

Public Function fncParseURI(ByVal isHref As String, isDrive As String, isPath As String, _
  isFileName As String, isID As String, Optional isDefDrive As Variant) As Boolean
  
  fncParseURI = False
    
  Dim templCounter As Long, sDefPath As String, sDefDrive As String, tempsString As String
  
  If Not IsMissing(isDefDrive) Then _
    If Not fncParseURI(CStr(isDefDrive), sDefDrive, sDefPath, tempsString, tempsString) Then _
      Exit Function
  
  If (Not Right$(isHref, 1) = "\") And (Not Right$(isHref, 1) = "/") Then
    templCounter = InStrRev(isHref, "#", -1, vbBinaryCompare)
    If templCounter <> 0 Then
      isID = Right$(isHref, Len(isHref) - templCounter)
      isHref = Left$(isHref, templCounter - 1)
    End If
  End If
  
  isHref = LCase$(isHref)
  For templCounter = 1 To Len(isHref)
    If Mid$(isHref, templCounter, 1) = "/" Then Mid$(isHref, templCounter, 1) = "\"
  Next templCounter
  If Left$(isHref, 7) = "file:\\" Then isHref = Right$(isHref, Len(isHref) - 7)
  If Left$(isHref, 7) = "http:\\" Or Left$(isHref, 6) = "ftp:\\" Then Exit Function
     
  templCounter = InStr(1, isHref, ":", vbBinaryCompare)
  If templCounter <> 0 Then
    isDrive = Mid$(isHref, templCounter - 1, 2)
    isHref = Mid$(isHref, templCounter + 1, Len(isHref) - templCounter)
  Else
    isDrive = sDefDrive
  End If
  
  templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
  If templCounter <> 0 Then
    isPath = Left$(isHref, templCounter)
    isHref = Right$(isHref, Len(isHref) - templCounter)
  Else
    isPath = sDefPath
  End If
  
  templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
  If templCounter <> 0 And templCounter < Len(isHref) Then
    isFileName = Right$(isHref, Len(isHref) - templCounter)
    isHref = Left$(isHref, templCounter - 1)
  Else
    isFileName = isHref
  End If
  
  If Left$(isPath, 1) = "." Then
    If Left$(isPath, 2) = ".." Then templCounter = 2 Else templCounter = 1
    isPath = Right$(isPath, Len(isPath) - templCounter)
    
    If Not IsMissing(isDefDrive) Then
      Dim tempsD As String, tmpsP As String
      fncParseURI CStr(isDefDrive), tempsD, tmpsP, tempsString, tempsString
      If Not tempsD = isDrive Then GoTo SetDefault
      
      If Right$(tmpsP, 1) = "\" And Left$(isPath, 1) = "\" Then _
        isPath = Right$(isPath, Len(isPath) - 1)
      
      If templCounter = 1 Then
        isPath = tmpsP & isPath
      Else
        templCounter = InStrRev(tmpsP, "\", -1, vbBinaryCompare)
        If templCounter > 0 Then isPath = Left$(tmpsP, templCounter) & isPath
      End If
    Else
SetDefault:
    End If
  End If
  
  fncParseURI = True
End Function
