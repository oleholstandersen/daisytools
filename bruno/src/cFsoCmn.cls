VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFsoCmn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim oFSO As Object

Private Sub Class_Initialize()
  'debug.Print "cFSO.initialize"
  Set oFSO = CreateObject("scripting.fileSystemObject")
End Sub

Private Sub Class_Terminate()
  'debug.Print "cFSO.terminate"
  Set oFSO = Nothing
End Sub

Public Function fncCreateDirectoryChain(isPath As String) As Boolean
Dim lPointer As Long
  On Error GoTo errh
  
  isPath = Replace$(isPath, "/", "\")
  lPointer = InStr(lPointer + 1, isPath, "\", vbBinaryCompare)
  Do Until lPointer = 0
    If Not oFSO.folderExists(Left$(isPath, lPointer)) Then
      oFSO.createFolder (Left$(isPath, lPointer))
    End If
    lPointer = InStr(lPointer + 1, isPath, "\", vbBinaryCompare)
  Loop
  
  fncCreateDirectoryChain = True
  Exit Function
errh:
  fncCreateDirectoryChain = False
End Function

Public Function fncMoveFile(sFileToMove As String, sDestination As String, bOverwrite As Boolean) As Boolean
Dim oFile As Object
    fncMoveFile = False
    On Error GoTo errhandler
    
    If bOverwrite Then
      If oFSO.fileExists(sDestination) Then
        oFSO.deletefile (sDestination)
      End If
    End If
    
    oFSO.MoveFile sFileToMove, sDestination
    fncMoveFile = True
errhandler:
    Set oFile = Nothing
End Function

Public Function fncGetFileName(sAbsPath As String) As String
Dim oFile As Object
    fncGetFileName = (oFSO.getfilename(sAbsPath))
    Set oFile = Nothing
End Function

Public Function fncGetParentFolderName(sAbsFilePath As String) As String
Dim oFile As Object, oFolder As Object
    
    On Error GoTo errh
    
    If oFSO.fileExists(sAbsFilePath) Then
      Set oFile = oFSO.getFile(sAbsFilePath)
      Set oFolder = oFile.ParentFolder
      fncGetParentFolderName = oFolder.Name
      Exit Function
    End If
        
errh:
    fncGetParentFolderName = ""
    Set oFile = Nothing
    Set oFolder = Nothing
End Function

Public Function fncGetFileNameLessExtension(ByVal sFileNameWithExtension As String) As String
Dim sBaseName As String
  sBaseName = Left(sFileNameWithExtension, InStrRev(sFileNameWithExtension, ".") - 1)
  fncGetFileNameLessExtension = sBaseName
End Function


Public Function fncCopyFile(sFileToMove As String, sDestination As String, bOverwrite As Boolean) As Boolean
Dim oFile As Object
    fncCopyFile = False
    On Error GoTo errhandler
    oFSO.CopyFile sFileToMove, sDestination, bOverwrite
    fncCopyFile = True
errhandler:
  Set oFile = Nothing
End Function

Public Function fncFileExists(sCandidate As String) As Boolean
    On Error GoTo errhandler
    fncFileExists = False
    If oFSO.fileExists(sCandidate) Then fncFileExists = True
errhandler:
End Function

Public Function fncFolderExists(sCandidate As String) As Boolean
    On Error GoTo errhandler
    fncFolderExists = False
    If oFSO.folderExists(sCandidate) Then fncFolderExists = True
errhandler:
End Function

Public Function fncGetFSOPathFromString(isString As String) As String
Dim oFile As Object
  On Error GoTo errh
  
  If oFSO.fileExists(isString) Then
    Set oFile = oFSO.getFile(isString)
    fncGetFSOPathFromString = oFile.Path
    Exit Function
  Else
   'file didnt exist
  End If

errh:
  fncGetFSOPathFromString = isString
End Function

Public Function fncIsFileDescendantOfFolder(sChildFileFullPath As String, sFolderPath As String) As Boolean
Dim oFileToTest As Object, oFolderToTest As Object, oCurrentFolder As Object
  'returns true if sChildFileFullPath exists and is descendant of folder
  
  fncIsFileDescendantOfFolder = True
  
  If oFSO.fileExists(sChildFileFullPath) Then
    If oFSO.folderExists(sFolderPath) Then
     Set oFileToTest = oFSO.getFile(sChildFileFullPath)
     Set oFolderToTest = oFSO.GetFolder(sFolderPath)
     'iterate upwards
     Set oCurrentFolder = oFileToTest.ParentFolder
     If oCurrentFolder.Path = oFolderToTest.Path Then Exit Function
     
     Do
       Set oCurrentFolder = oCurrentFolder.ParentFolder
     Loop Until oCurrentFolder.Path = oFolderToTest.Path Or oCurrentFolder.IsRootFolder
     
     If oCurrentFolder.Path = oFolderToTest.Path Then Exit Function

  
    End If 'oFSO.folderExists(sFolderPath)
  End If 'oFSO.fileExists(sChildFullPath)

  fncIsFileDescendantOfFolder = False
  
End Function



Public Function fncGetFolderFiles(sFolderPath As String) As Object
Dim oFiles As Object, oFolder As Object
  On Error GoTo errhandler
  
  Set oFolder = oFSO.GetFolder(sFolderPath)
  If Not oFolder Is Nothing Then
    Set fncGetFolderFiles = oFolder.Files
  End If
  
errhandler:
  Set oFiles = Nothing
  Set oFolder = Nothing
End Function

Public Function fncGetExtensionFromFileObject(oFile As Object) As String
  On Error GoTo errhandler
  fncGetExtensionFromFileObject = oFSO.getExtensionName(oFile.Path)
errhandler:
End Function

Public Function fncGetExtensionFromString(sFullPath As String) As String
Dim oFile As Object
  On Error GoTo errhandler
  Set oFile = oFSO.getFile(sFullPath)
  If Not oFile Is Nothing Then
    fncGetExtensionFromString = fncGetExtensionFromFileObject(oFile)
    Exit Function
  End If
errhandler:
  fncGetExtensionFromString = "fncGetExtensionFromString"
End Function

Public Function fncDeleteFile(oFile As Object) As Boolean
  On Error GoTo errh
  fncDeleteFile = False
  oFSO.deletefile (oFile.Path)
  fncDeleteFile = True
errh:
  If Not oFile Is Nothing Then fncDeleteFile = False
End Function

Public Function fncDeleteAllFilesInFolder(sFolderPath As String) As Boolean
Dim oFiles As Object, oFile As Object
  On Error GoTo errh
  fncDeleteAllFilesInFolder = False
  Set oFiles = fncGetFolderFiles(sFolderPath)
  If oFiles Is Nothing Then GoTo errh
  For Each oFile In oFiles
    fncDeleteFile oFile
  Next
  fncDeleteAllFilesInFolder = True
  Exit Function
errh:
  fncDeleteAllFilesInFolder = False
End Function

Public Function fncGetFileAsString(sAbsPath As String) As String
Dim oStream As Object
    On Error GoTo errhandler
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.opentextfile(sAbsPath)
    fncGetFileAsString = oStream.ReadAll
errhandler:
    oStream.Close
End Function
