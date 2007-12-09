Attribute VB_Name = "mFileHandler"
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

Public Function fncCreateDirectoryChain( _
    isPath As String, _
    objOwner As oRegenerator _
    )
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim lPointer As Long

  If Not Right(isPath, 1) = "\" Then isPath = isPath & "\"

' Create all non-existing directories in a full path specification given
  lPointer = InStr(lPointer + 1, isPath, "\", vbBinaryCompare)
  Do Until lPointer = 0
    If Not oFSO.folderexists(Left$(isPath, lPointer)) Then
      oFSO.createFolder (Left$(isPath, lPointer))
    End If
    lPointer = InStr(lPointer + 1, isPath, "\", vbBinaryCompare)
  Loop
  Set oFSO = Nothing
End Function

Public Function fncFolderExists( _
    sCandidate As String, _
    objOwner As oRegenerator _
    ) As Boolean
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo ErrHandler
    fncFolderExists = False
    If oFSO.folderexists(sCandidate) Then fncFolderExists = True
ErrHandler:
    Set oFSO = Nothing
End Function

Public Function fncFileExists( _
    sCandidate As String, _
    objOwner As oRegenerator _
    ) As Boolean
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")

    On Error GoTo ErrHandler
    fncFileExists = False
    If oFSO.FileExists(sCandidate) Then fncFileExists = True
ErrHandler:
    Set oFSO = Nothing
End Function

Public Function fncGetFileName(sAbsPath As String) As String
Dim oFSO As Object, oFile As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    fncGetFileName = (oFSO.getFileName(sAbsPath))
    Set oFSO = Nothing: Set oFile = Nothing
End Function

Public Function fncGetParentFolderName(sAbsPath As String) As String
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    fncGetParentFolderName = oFSO.GetParentFolderName(sAbsPath)
    Set oFSO = Nothing
End Function

Public Function fncGetParentFolderName2(sAbsPath As String) As String
Dim oFSO As Object, file As Object
Dim s As String
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set file = oFSO.getFile(sAbsPath)
    s = file.ParentFolder
    fncGetParentFolderName2 = s
    Set oFSO = Nothing: Set file = Nothing
    
End Function

Public Function fncGetExtensionFromFileObject(oFile As Object) As String
Dim oFSO As Object
  On Error GoTo ErrHandler
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  fncGetExtensionFromFileObject = oFSO.getExtensionName(oFile.Path)
ErrHandler:
 Set oFSO = Nothing
End Function

Public Function fncGetExtensionFromString(sFullPath As String) As String
Dim oFile As Object, oFSO As Object
  On Error GoTo ErrHandler
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFile = oFSO.getFile(sFullPath)
  If Not oFile Is Nothing Then
    fncGetExtensionFromString = fncGetExtensionFromFileObject(oFile)
    Set oFSO = Nothing
    Exit Function
  End If
ErrHandler:
  fncGetExtensionFromString = "fncGetExtensionFromString"
  Set oFSO = Nothing
End Function


Public Function fncGetFileAsString(sAbsPath As String) As String
Dim oStream As Object, oFSO As Object
    On Error GoTo ErrHandler
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.opentextfile(sAbsPath)
    fncGetFileAsString = oStream.ReadAll
ErrHandler:
    oStream.Close
End Function

Public Function fncGetMotherFolder(sAbsPath As String) As String
'returns only the current mother folder (whereas ofso.getparentfoldername returns the whole path...)
Dim lBegin As Long
  On Error GoTo ErrHandler

  sAbsPath = Replace(sAbsPath, "/", "\")
  If Right(sAbsPath, 1) = "\" Then sAbsPath = Mid(sAbsPath, 1, Len(sAbsPath) - 1)
  lBegin = InStrRev(sAbsPath, "\") + 1
  fncGetMotherFolder = Mid(sAbsPath, lBegin)
  Exit Function
ErrHandler:

End Function

Public Function fncReadFile(sAbsPath As String) As String
Dim oFSO As Object
Dim oStream As Object
    On Error GoTo ErrHandler
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.opentextfile(sAbsPath)
    fncReadFile = oStream.ReadAll
ErrHandler:
    oStream.Close
    Set oFSO = Nothing
End Function

Public Function fncSaveFile(sAbsPath As String, sContent As String, _
  objOwner As oRegenerator) As Boolean
  
Dim oFSO As Object
Dim oStream As Object
    On Error GoTo ErrHandler
    fncSaveFile = False
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sAbsPath, True)
    oStream.Write (sContent)
    oStream.Close
    fncSaveFile = True
ErrHandler:

End Function

Public Function fncMoveFile( _
    sFileToMove As String, _
    sDestination As String, _
    objOwner As oRegenerator _
    ) As Boolean
Dim oFSO As Object, oFile As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")

    fncMoveFile = False
    On Error GoTo ErrHandler
    oFSO.MoveFile sFileToMove, sDestination
    fncMoveFile = True
ErrHandler:
    Set oFSO = Nothing: Set oFile = Nothing
    If Not fncMoveFile Then
      objOwner.addlog ("<errH in='fncMoveFile' file='" & sFileToMove & "' dest='" & sDestination & "'>fncMoveFile ErrHandler</errH>")
    End If
End Function

Public Function fncCopyFile( _
    sFileToMove As String, _
    sDestination As String, _
    objOwner As oRegenerator _
    ) As Boolean
Dim oFSO As Object, oFile As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    fncCopyFile = False
    On Error GoTo ErrHandler
    oFSO.copyFile sFileToMove, sDestination
    fncCopyFile = True
ErrHandler:
    Set oFSO = Nothing: Set oFile = Nothing
    If Not fncCopyFile Then objOwner.addlog ("<errH in='fncCopyFile' file='" & sFileToMove & "' dest='" & sDestination & "'>fncCopyFile ErrH</errH>")
End Function

Public Function fncGetExtension(ByVal sFileName As String) As String
Dim oFSO As Object, oFile As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    fncGetExtension = Chr(46) & oFSO.getExtensionName(sFileName)
End Function

Public Function fncFileCount(sAbsPath As String) As Long
Dim oFSO As Object, oFolder As Object, oFiles As Object, oFile As Object
    
    On Error GoTo ErrHandler

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getfolder(oFSO.GetAbsolutePathName(sAbsPath))
    Set oFiles = oFolder.Files
    fncFileCount = oFiles.Count

ErrHandler:
End Function

Public Function fncDeleteFolder(sAbsPath As String, objOwner As oRegenerator) _
  As Boolean
  
Dim oFSO As Object, oFolder As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo ErrHandler
    fncDeleteFolder = False
    If oFSO.folderexists(oFSO.GetAbsolutePathName(sAbsPath)) Then
        oFSO.DeleteFolder oFSO.GetAbsolutePathName(sAbsPath), True
    End If
    fncDeleteFolder = True
ErrHandler:
  If Not fncDeleteFolder Then objOwner.addlog "<errH in='fncDeleteFolder'>fncDeleteFolder ErrH</errH>"
End Function

Public Function fncCreateFolder( _
    sNewFolderPath As String, objOwner As oRegenerator, _
    Optional bMakeUnique As Boolean _
    ) As String
Dim oFSO As Object, oFolder As Object
Set oFSO = CreateObject("Scripting.FileSystemObject")
  
  On Error GoTo ErrHandler
    
  If Not oFSO.folderexists(sNewFolderPath) Then
    Set oFolder = oFSO.createFolder(sNewFolderPath)
  ElseIf bMakeUnique Then
    Do
      sNewFolderPath = Mid(sNewFolderPath, 1, Len(sNewFolderPath) - 1) & "_\"
    Loop Until Not oFSO.folderexists(sNewFolderPath)
    Set oFolder = oFSO.createFolder(sNewFolderPath)
  Else
   objOwner.addlog "<error in='fncCreateFolder'>attempt to create already existing folder without setting bMakeUnique to true</error>"
   GoTo ErrHandler
  End If
  fncCreateFolder = sNewFolderPath
  Exit Function
ErrHandler:
  objOwner.addlog "<errH in='fncCreateFolder' file='" & sNewFolderPath & "'>fncCreateFolder ErrH</errH>"
End Function

Public Function fncRemoveStringFromEachFileInFolder( _
    ByVal sFolderPath As String, _
    ByVal sSafetyString As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
    
    'On Error GoTo ErrHandler

fncRemoveStringFromEachFileInFolder = False
'only the audio, img and other types end up here
'removes the safetystring made during rename proc
'if still a collision, appends underscores

Dim oFSO As Object, oFolder As Object, oFiles As Object, oFile As Object, oSubFolders As Object, oSubFolder As Object
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFolder = oFSO.getfolder(sFolderPath)
  
  Set oSubFolders = oFolder.subFolders
  For Each oSubFolder In oSubFolders
    If Not (oSubFolder.Attributes And 64) Then
      If InStr(1, oSubFolder.Name, "rgn~") < 1 Then
        If Not fncRemoveStringFromEachFileInFolder(oSubFolder.Path, sSafetyString, objOwner) Then GoTo ErrHandler
      Else
        '-it was a regen backup file
      End If
    Else
      Debug.Print oSubFolder.Path & " is an alias"
    End If
  Next
  
  Set oFiles = oFolder.Files
  For Each oFile In oFiles
    If Not (oFile.Attributes And 64) Then
      'Debug.Print oFile.Path & " run"
      If InStr(1, oFile.Name, sSafetyString) > 0 Then
        Dim sNewFileName As String
        sNewFileName = Replace(oFile.Path, sSafetyString, "")
        If oFSO.FileExists(sNewFileName) Then
          Dim oCollisionFile As Object, sUniquePathName As String
          Set oCollisionFile = oFSO.getFile(sNewFileName)
          sUniquePathName = oCollisionFile.Path
          Do
            sUniquePathName = sUniquePathName & "_"
          Loop Until Not oFSO.FileExists(sUniquePathName)
          objOwner.addlog "<warning in='fncRemoveStringFromEachFileInFolder'>renamed unreferenced " & sNewFileName & " to " & sUniquePathName & "</errH>"
          oCollisionFile.Move sUniquePathName
        End If
        oFile.Move sNewFileName
      End If
    Else
      Debug.Print oFile.Path & " is an alias"
    End If
  Next
  fncRemoveStringFromEachFileInFolder = True

ErrHandler:
 Set oFSO = Nothing
 If Not fncRemoveStringFromEachFileInFolder Then objOwner.addlog "<errH in='fncRemoveStringFromEachFileInFolder' folder='" & sFolderPath & "'>fncRemoveStringFromEachFileInFolder ErrH</errH>"
End Function

Public Function fncRenameFilesetFile( _
    ByVal sOrigAbsPath As String, _
    ByVal sNewAbsPath As String, _
    ByVal sBackupPath As String, _
    objOwner As oRegenerator _
    ) As Boolean

Dim oFSO As Object, oFile As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
            
    On Error GoTo ErrHandler
    fncRenameFilesetFile = False
    
    If Not oFSO.FileExists(sOrigAbsPath) Then GoTo ErrHandler
        
    'mg fix 20030317 etc, 20050418
    If sOrigAbsPath <> sNewAbsPath Then 'else no need to rename
      If LCase$(sOrigAbsPath) = LCase$(sNewAbsPath) Then
        'if there are only case diffs
        Set oFile = oFSO.getFile(sOrigAbsPath)
        'oFSO.copyFile oFile.Path, sBackupPath
        'objOwner.addlog "<warning in='fncRenameFile'>moved preexisting file " & sOrigAbsPath & " to " & sBackupPath & " </warning>"
        oFile.Move sNewAbsPath
      Else
        'if there are more than case diffs
        'mg 20040216: if destination name exists as a file, move it
        If oFSO.FileExists(sNewAbsPath) Then
          objOwner.addlog "<warning in='fncRenameFile'>moved preexisting file " & sNewAbsPath & " to " & sBackupPath & " </warning>"
          Set oFile = oFSO.getFile(sNewAbsPath)
          oFile.Move sBackupPath
        End If
        Set oFile = oFSO.getFile(sOrigAbsPath)
        oFile.Name = oFSO.getFileName(sNewAbsPath)
      End If
    End If 'sOrigAbsPath <> sNewAbsPath
        
    fncRenameFilesetFile = True
ErrHandler:
  Set oFSO = Nothing: Set oFile = Nothing
  If Not fncRenameFilesetFile Then objOwner.addlog "<errH in='fncRenameFilesetFile' file='" & sOrigAbsPath & "'>fncRenameFilesetFile ErrH</errH>"
End Function

Public Function fncGetAbsolutePathName( _
    ByVal sInPath As String, _
    ByRef sReturn As String _
    ) As Boolean
Dim oFSO As Object

  Set oFSO = CreateObject("Scripting.FileSystemObject")
  On Error GoTo ErrHandler
  fncGetAbsolutePathName = False

  sReturn = oFSO.GetAbsolutePathName(sInPath)

  fncGetAbsolutePathName = True

ErrHandler:
  Set oFSO = Nothing
  'If Not fncGetAbsolutePathName Then objOwner.addlog "<errH in='fncGetAbsolutePathName' inpath='" & sInPath & "'>fncGetAbsolutePathName ErrH</errH>"
End Function

Public Function fncCreateFileNameLog(objOwner As oRegenerator) As Boolean
Dim i As Long
Dim sTemp As String
Dim oNode As IXMLDOMNode
Dim oElem As IXMLDOMElement
Dim oAttr As IXMLDOMAttribute
Dim oDoc As New MSXML2.DOMDocument40
Dim sID As String

    oDoc.async = False
    oDoc.validateOnParse = False
    oDoc.resolveExternals = False
    oDoc.setProperty "SelectionLanguage", "XPath"
    oDoc.setProperty "NewParser", True

  fncCreateFileNameLog = False
  On Error GoTo ErrHandler
  
  Dim sProcIns As String: sProcIns = ""
  sProcIns = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & " ?>" & vbCrLf
  If Not fncParseString(sProcIns & "<rename_nfo></rename_nfo>", oDoc, objOwner) Then Exit Function
  
  Set oNode = objOwner.objCommonMeta.DcIdentifier
  If Not oNode Is Nothing Then sID = oNode.selectSingleNode("@content").Text Else sID = ""
  
  If Not fncAppendChild(oDoc.documentElement, "renamed", objOwner, , "date", Format(Date, "yyyy-mm-dd"), "appversion", sAppVersion, "id", sID) Then Exit Function
  For i = 0 To objOwner.objFileSetHandler.aInFileSetMembers - 1
    If Not fncAppendChild(oDoc.documentElement, "file", objOwner, , "num", CStr(i), "origName", _
    fncGetFileName(objOwner.objFileSetHandler.aInFileSet(i).sAbsPath), "newName", objOwner.objFileSetHandler.aOutFileSet(i).sFileName) Then Exit Function
  Next i
  'pretty print
  If Not fncPrettyInstance(oDoc, "  ", objOwner) Then GoTo ErrHandler
  'save
  oDoc.save (objOwner.sBackupPath & "fileRename_nfo.xml")
  objOwner.addlog ("<status>saved \fileRename_nfo.xml</status>")
  fncCreateFileNameLog = True
ErrHandler:
  Set oDoc = Nothing
  If Not fncCreateFileNameLog Then objOwner.addlog "<errH in='fncCreateFileNameLog'>fncCreateFileNameLog ErrH</errH>"
End Function

Public Function fncMoveUnref(sDateString As String, objOwner As oRegenerator) As Boolean
Dim oFSO As Object, oFolder As Object, oFiles As Object, oFile As Object

    On Error GoTo ErrHandler
    fncMoveUnref = False
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getfolder(objOwner.sDtbFolderPath)
    Set oFiles = oFolder.Files

    For Each oFile In oFiles
        If Not objOwner.objFileSetHandler.fncItemExistsInOutArray(oFile.Name) Then
            If objOwner.sUnrefPath = "" Or Not fncFolderExists(objOwner.sUnrefPath, objOwner) Then
              objOwner.sUnrefPath = fncCreateFolder(objOwner.sDtbFolderPath & "rgn~unr_" & sDateString & "\", objOwner, True)
            End If
            Dim sFileName As String
            sFileName = oFile.Name
            If Not fncMoveFile(oFile.Path, objOwner.sUnrefPath, objOwner) Then GoTo ErrHandler
            '20040615 mg spaces in filenames breaks md5 checksum
            'trunc the filename
            'WARNING this should not be done in official version
            If Not fncIsValidUriChars(sFileName) Then
              'there was a non-ascii7 char, so rename the file in unref to this
              Dim oRenameFile As Object
              Set oRenameFile = oFSO.getFile(objOwner.sUnrefPath & sFileName)
              oRenameFile.Name = fncTruncToValidUriChars(sFileName)
              'If Not fncRenameFile(objOwner.sUnrefPath & sFileName, objOwner.sUnrefPath & fncTruncToValidUriChars(sFileName), "", objOwner) Then GoTo ErrHandler
            End If
        End If
    Next
    fncMoveUnref = True
ErrHandler:
  If Not fncMoveUnref Then objOwner.addlog "<errH in='fncMoveUnref'>fncMoveUnref ErrH</errH>"
End Function

