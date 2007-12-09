Attribute VB_Name = "IteratorFileHandler"
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
Public Function fncCreateDirectoryChain( _
    isPath As String _
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
    sCandidate As String _
    ) As Boolean
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo ErrHandler
    fncFolderExists = False
    If oFSO.folderexists(sCandidate) Then fncFolderExists = True
ErrHandler:
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

End Function


