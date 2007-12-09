Attribute VB_Name = "mHelpersFile"
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

Public Function fncRenameFile( _
    ByVal sCurrentPath As String, _
    ByVal sNewName As String) _
    As Boolean

Dim oFSO As Object, oFile As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo ErrHandler
    fncRenameFile = False

    If Not fncFileExists(sCurrentPath) Then Exit Function
    If sCurrentPath <> sNewName Then 'else no need to rename
        Set oFile = oFSO.GetFile(sCurrentPath)
        oFile.Name = sNewName
    End If

    fncRenameFile = True
ErrHandler:
  If Not fncRenameFile Then addLog "fncRenameFile ErrH"
End Function


Public Function fncFolderExists(sCandidate As String) As Boolean
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo ErrHandler
    fncFolderExists = False

    If oFSO.folderExists(sCandidate) Then
        fncFolderExists = True
    Else
        addLog (sCandidate & " not found")
    End If
ErrHandler:
    Set oFSO = Nothing
End Function

Public Function fncFileExists(sCandidate As String, Optional bQuiet As Boolean) As Boolean
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo ErrHandler
    fncFileExists = False

    If oFSO.FileExists(sCandidate) Then
        fncFileExists = True
    Else
        If Not bQuiet Then addLog (sCandidate & " not found")
    End If
ErrHandler:
    Set oFSO = Nothing
End Function

Public Function fncGetFileName(sAbsPath As String) As String
Dim oFSO As Object, oFile As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    fncGetFileName = (oFSO.GetFileName(sAbsPath))
    Set oFSO = Nothing: Set oFile = Nothing
End Function

Public Function fncGetParentFolderName(sAbsPath As String) As String
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    fncGetParentFolderName = oFSO.GetParentFolderName(sAbsPath)
    Set oFSO = Nothing
End Function

Public Function fncReadFile(sAbsPath As String) As String
Dim oFSO As Object
Dim oStream As Object
    On Error GoTo ErrHandler
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.OpenTextFile(sAbsPath)
    fncReadFile = oStream.ReadAll
ErrHandler:
    oStream.Close
    Set oFSO = Nothing
End Function

Public Function fncSaveFile(sAbsPath As String, sContent As String) As Boolean
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
    If Not fncSaveFile Then addLog ("fncSaveFile ErrHandler")
End Function

Public Function fncMoveFile(sFileToMove As String, sDestination As String) As Boolean
Dim oFSO As Object, oFile As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")

    fncMoveFile = False
    On Error GoTo ErrHandler
    oFSO.MoveFile sFileToMove, sDestination
    fncMoveFile = True
ErrHandler:
    Set oFSO = Nothing: Set oFile = Nothing
    If Not fncMoveFile Then addLog ("fncMoveFile ErrHandler")
End Function

Public Function fncCreateFolder( _
    sNewFolderPath As String, _
    Optional bMakeUnique As Boolean _
    ) As String
Dim oFSO As Object, oFolder As Object
Set oFSO = CreateObject("Scripting.FileSystemObject")
  
  On Error GoTo ErrHandler
    
  If Not oFSO.folderExists(sNewFolderPath) Then
    Set oFolder = oFSO.CreateFolder(sNewFolderPath)
  ElseIf bMakeUnique Then
    Do
      sNewFolderPath = Mid(sNewFolderPath, 1, Len(sNewFolderPath) - 1) & "_\"
    Loop Until Not oFSO.folderExists(sNewFolderPath)
    Set oFolder = oFSO.CreateFolder(sNewFolderPath)
  Else
   addLog "attempt to create already existing folder without setting bMakeUnique to true"
   GoTo ErrHandler
  End If
  fncCreateFolder = sNewFolderPath
ErrHandler:
  
End Function

Public Function fncCopyFile(sFileToMove As String, sDestination As String) As Boolean
Dim oFSO As Object, oFile As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    'If bDebugMode Then oEvent.addStatusLog ("fncCopyFile: " & sFileToMove)
    fncCopyFile = False
    On Error GoTo ErrHandler
    oFSO.CopyFile sFileToMove, sDestination
    fncCopyFile = True
ErrHandler:
    Set oFSO = Nothing: Set oFile = Nothing
    If Not fncCopyFile Then addLog ("fncCopyFile ErrHandler")
End Function

Public Function fncGetExtension(ByVal sFileName As String) As String
Dim oFSO As Object, oFile As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    fncGetExtension = Chr(46) & oFSO.GetExtensionName(sFileName)
End Function
