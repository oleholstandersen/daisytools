Attribute VB_Name = "mHelpers"
' Daisy 2.02 Regenerator Batch UI
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

' This function converts a checkbox control value to a boolean
' vbChecked = true, vbUnchecked = false, vbGrayed = false
Public Function fncCheck2Bol(iobjControl As Object) As Boolean
  If (iobjControl.Value = vbChecked) Then _
    fncCheck2Bol = True Else fncCheck2Bol = False
End Function

Public Function fncBol2Check(ibolValue As Boolean) As CheckBoxConstants
  If (ibolValue) Then _
    fncBol2Check = vbChecked Else fncBol2Check = vbUnchecked
End Function

' Create all non-existing directories in a full path specification given
Public Function fncCreateDirectoryChain(isPath As String)
  Dim lPointer As Long
  
  lPointer = InStr(lPointer + 1, isPath, "\", vbBinaryCompare)
  Do Until lPointer = 0
    If Not oFSO.folderexists(Left$(isPath, lPointer)) Then
      oFSO.createFolder (Left$(isPath, lPointer))
    End If
    lPointer = InStr(lPointer + 1, isPath, "\", vbBinaryCompare)
  Loop
End Function


' ***** URI Helpers *****

' Get filename from a URI
Public Function fncGetUriFileName(sAbsPath As String) As String
  Dim sTemp As String, sFileName As String
  If Not fncParseURI(sAbsPath, sTemp, sTemp, sFileName, sTemp) Then Exit Function
  fncGetUriFileName = sFileName
End Function

' Get directory name from a URI
Public Function fncGetPathName(sAbsPath As String) As String
  Dim sTemp As String, sDrive As String, sPath As String
  If Not fncParseURI(sAbsPath, sDrive, sPath, sTemp, sTemp) Then Exit Function
  fncGetPathName = sDrive & sPath
  'added mg 20030325
  If Not Right$(fncGetPathName, 1) = "\" Then fncGetPathName = fncGetPathName & "\"
End Function

' Remove the id from a URI
Public Function fncStripId(sAbsPath As String) As String
  Dim sD As String, sP As String, sF As String, sTemp As String
  fncParseURI sAbsPath, sD, sP, sF, sTemp
  fncStripId = sD & sP & sF
End Function

' Remove the id and add a basepath to a URI
Public Function fncStripIdAddPath( _
    ByVal sRelPath As String, ByVal sBasePath As String _
    ) As String
  Dim sD As String, sP As String, sF As String, sTemp As String
  fncParseURI sRelPath, sD, sP, sF, sTemp, sBasePath
  fncStripIdAddPath = sD & sP & sF
End Function

' Get the id from an URI
Public Function fncGetId( _
    ByVal sPath As String) As String
  Dim sTemp As String, sID As String
    
    fncParseURI sPath, sTemp, sTemp, sTemp, sID
    fncGetId = sID
End Function

' This function parses a URI (isHref) and returns it divided in several smaller
' variables (isDrive, isPath, isFileName, isID). It has also the option of parsing
' a relative path getting the absolute path by suppling the path that it is relating
' from (isDefDrive)
Public Function fncParseURI(ByVal isHref As String, isDrive As String, isPath As String, _
  isFileName As String, isID As String, Optional isDefDrive As Variant) As Boolean
  
  fncParseURI = False
    
  Dim templCounter As Long, sDefPath As String, sDefDrive As String, tempsString As String
  Dim sDrive As String, sPath As String, sFileName As String, sID As String
  
  If Not IsMissing(isDefDrive) Then _
    If Not fncParseURI(CStr(isDefDrive), sDefDrive, sDefPath, tempsString, tempsString) Then _
      Exit Function
  
  If (Not Right$(isHref, 1) = "\") And (Not Right$(isHref, 1) = "/") Then
    templCounter = InStrRev(isHref, "#", -1, vbBinaryCompare)
    If templCounter <> 0 Then
      sID = Right$(isHref, Len(isHref) - templCounter)
      isHref = Left$(isHref, templCounter - 1)
    End If
  End If
  
  For templCounter = 1 To Len(isHref)
    If Mid$(isHref, templCounter, 1) = "/" Then Mid$(isHref, templCounter, 1) = "\"
  Next templCounter
  
  If LCase$(Left$(isHref, 7)) = "file:\\" Then _
    isHref = Right$(isHref, Len(isHref) - 7)
  If LCase$(Left$(isHref, 7)) = "http:\\" Or Left$(isHref, 6) = "ftp:\\" Then _
    Exit Function
     
  templCounter = InStr(1, isHref, ":", vbBinaryCompare)
  If templCounter <> 0 Then
    sDrive = Mid$(isHref, templCounter - 1, 2)
    isHref = Mid$(isHref, templCounter + 1, Len(isHref) - templCounter)
  Else
    sDrive = sDefDrive
  End If
  
  templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
  If templCounter <> 0 Then
    sPath = Left$(isHref, templCounter)
    isHref = Right$(isHref, Len(isHref) - templCounter)
  End If
      
  templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
  If templCounter <> 0 And templCounter < Len(isHref) Then
    sFileName = Right$(isHref, Len(isHref) - templCounter)
    isHref = Left$(isHref, templCounter - 1)
  Else
    sFileName = isHref
  End If
  
'  If Left$(isPath, 1) = "." Then
'    If Left$(isPath, 2) = ".." Then templCounter = 2 Else templCounter = 1
'    isPath = Right$(isPath, Len(isPath) - templCounter)
'
'    If Not IsMissing(isDefDrive) Then
'      Dim tempsD As String, tmpsP As String
'      fncParseURI CStr(isDefDrive), tempsD, tmpsP, tempsString, tempsString
'      If Not tempsD = isDrive Then GoTo SetDefault
'
'      If Right$(tmpsP, 1) = "\" And Left$(isPath, 1) = "\" Then _
'        isPath = Right$(isPath, Len(isPath) - 1)
'
'      If templCounter = 1 Then
'        isPath = tmpsP & isPath
'      Else
'        templCounter = InStrRev(tmpsP, "\", -1, vbBinaryCompare)
'        If templCounter > 0 Then isPath = Left$(tmpsP, templCounter) & isPath
'      End If
'    Else
'SetDefault:
'    End If
'  End If
  
  Do Until Not Left$(sPath, 1) = "."
    If Left$(sPath, 3) = "..\" Then
      sPath = Right$(isPath, Len(sPath) - 3)
      templCounter = InStrRev(sDefPath, "\", Len(sDefPath) - 1, vbBinaryCompare)
      sDefPath = Left$(sDefPath, templCounter)
    ElseIf Left$(sPath, 2) = ".\" Then
      sPath = Right$(sPath, Len(sPath) - 2)
    End If
  Loop
  
  If Not Left$(sPath, 1) = "\" Then
    If Not Right$(sDefPath, 1) = "\" Then sDefPath = sDefPath & "\"
    sPath = sDefPath & sPath
  End If
  
  If Not (sDrive = "" And sPath = "" And sFileName = "") Then
    If fncValidatePathSections(sDrive, sPath, sFileName) Then fncParseURI = True
  Else
    fncParseURI = True
  End If
  
  isDrive = sDrive
  isPath = sPath
  isFileName = sFileName
End Function

Public Function fncValidatePath(isPath As String) As Boolean
  Dim sD As String, sP As String, sF As String, sTemp As String
  If fncParseURI(isPath, sD, sP, sF, sTemp) Then fncValidatePath = True
End Function

Function fncValidatePathSections(isDrive As String, isPath As String, _
  isFileName As String) As Boolean
  
  Dim lCounter As Long, lValue As Long

' Drive validation
  If Len(isDrive) <> 2 Then Exit Function
  lValue = Asc(Left$(isDrive, 1))
  If lValue < 65 Or (lValue > 90 And lValue < 97) Or lValue > 122 Then Exit Function
  
' Path validation
  If Not fncIsValidUriChars(isPath) Then Exit Function
  
' File validation
  If Not fncIsValidUriChars(isFileName) Then Exit Function
  
  fncValidatePathSections = True
End Function

' This function checks wheter the given URI contains valid characters or not
'
Private Function fncIsValidUriChars(ByVal isURI As String) As Boolean
Dim i As Long
    fncIsValidUriChars = False
    For i = 1 To Len(isURI)
        If Not fncIsValidUriChar(Asc(Mid$(isURI, i, 1))) Then Exit Function
    Next i
    fncIsValidUriChars = True
End Function

' This function checks wheter the given character is a valid URI character
'
Private Function fncIsValidUriChar(ByVal sChar As String) As Boolean
    fncIsValidUriChar = False
    If Not ((CInt(sChar) > 44 And CInt(sChar) < 58) Or _
      (CInt(sChar) > 63 And CInt(sChar) < 93) Or _
      (CInt(sChar) > 94 And CInt(sChar) < 123) Or _
      (CInt(sChar) = 32 Or CInt(sChar) = 39)) Then _
        Exit Function
    fncIsValidUriChar = True
End Function

' This function converts a given string to valid URI characters (removes all non-valid
' characters)
Public Function fncConvertToUri(isInput As String, isOutput As String) As Boolean
  Dim lCounter As Long
  
  isOutput = ""
  
  For lCounter = 1 To Len(isInput)
    If fncIsValidUriChar(Asc(Mid$(isInput, lCounter, 1))) Then _
      isOutput = isOutput & Mid$(isInput, lCounter, 1)
  Next lCounter
  
  fncConvertToUri = True
End Function
