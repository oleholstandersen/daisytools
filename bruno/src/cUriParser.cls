VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUriParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Function fncGetUriFileName(sAbsPath As String) As String
  Dim sTemp As String, sFileName As String
  If Not fncParseURI(sAbsPath, sTemp, sTemp, sFileName, sTemp) Then Exit Function
  fncGetUriFileName = sFileName
End Function

Public Function fncGetPathName(sAbsPath As String) As String
  Dim sTemp As String, sDrive As String, sPath As String
  If Not fncParseURI(sAbsPath, sDrive, sPath, sTemp, sTemp) Then Exit Function
  fncGetPathName = sDrive & sPath
End Function

Public Function fncStripId(sAbsPath As String) As String
  Dim sD As String, sP As String, sF As String, sTemp As String
  fncParseURI sAbsPath, sD, sP, sF, sTemp
  fncStripId = sD & sP & sF
End Function

Public Function fncStripIdAddPath( _
    ByVal sRelPath As String, ByVal sBasePath As String _
    ) As String
  Dim sD As String, sP As String, sF As String, sTemp As String
  fncParseURI sRelPath, sD, sP, sF, sTemp, sBasePath
  fncStripIdAddPath = sD & sP & sF
End Function

Public Function fncGetId( _
    ByVal sPath As String _
    ) As String
  Dim sTemp As String, sID As String
    fncParseURI sPath, sTemp, sTemp, sTemp, sID
    fncGetId = sID
End Function

' This is a multi-purpose path separator, this function takes any relative or
' absolute path and separates it into Drive, Path, Filename, ID. It can also add
' an absolute path to a relative path.
Public Function fncParseURI(ByVal isHref As String, isDrive As String, isPath As String, _
  isFileName As String, isID As String, Optional isDefDrive As Variant) As Boolean
  
  fncParseURI = False
    
  Dim templCounter As Long, sDefPath As String, sDefDrive As String, tempsString As String
  Dim sDefFile As String
  
  If Not IsMissing(isDefDrive) Then
'    If Not (Left$(isDefDrive, 1) = "\" Or Left$(isDefDrive, 1) = "/") Then _
'      isDefDrive = isDefDrive & "\"
  
    If Not fncParseURI(CStr(isDefDrive), sDefDrive, sDefPath, sDefFile, tempsString) Then _
      Exit Function
  End If
  
  If (Not Right$(isHref, 1) = "\") And (Not Right$(isHref, 1) = "/") Then
         templCounter = InStrRev(isHref, "#", -1, vbBinaryCompare)
    If templCounter <> 0 Then
      isID = Right$(isHref, Len(isHref) - templCounter)
      isHref = Left$(isHref, templCounter - 1)
    End If
  End If
  
'  'debug.Print "No lcase$ in fncParseUri -> Check if errors"
  
  For templCounter = 1 To Len(isHref)
    If Mid$(isHref, templCounter, 1) = "/" Then Mid$(isHref, templCounter, 1) = "\"
  Next templCounter
  
  If LCase$(Left$(isHref, 7)) = "file:\\" Then _
    isHref = Right$(isHref, Len(isHref) - 7)
  If LCase$(Left$(isHref, 7)) = "http:\\" Or Left$(isHref, 6) = "ftp:\\" Then _
    Exit Function
  If LCase$(Left$(isHref, 7)) = "mailto:" Then Exit Function
     
  templCounter = InStr(1, isHref, ":", vbBinaryCompare)
  If templCounter <> 0 Then
    isDrive = Mid$(isHref, templCounter - 1, 2)
    isHref = Mid$(isHref, templCounter + 1, Len(isHref) - templCounter)
  Else
    isDrive = sDefDrive
  End If
  
  templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
  If templCounter <> 0 Then
    isPath = sDefPath & Left$(isHref, templCounter)
    isHref = Right$(isHref, Len(isHref) - templCounter)
  Else
    isPath = sDefPath
  End If
  
  templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
  If templCounter <> 0 And templCounter < Len(isHref) Then
    isFileName = Right$(isHref, Len(isHref) - templCounter)
    isHref = Left$(isHref, templCounter - 1)
  ElseIf isHref = "" Then
    isFileName = sDefFile
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

Public Function fncGetAttrValueFromString( _
    ByVal sString As String, _
    ByVal sAttrName As String _
    ) As String
    
Dim lStartPos As Long, lEndPos As Long
  sString = Replace(sString, Chr(39), Chr(34))
  'all quotes are now doubles
  lStartPos = InStr(1, sString, sAttrName & "=" & Chr(34)) + Len(sAttrName) + 2
  lEndPos = InStr(lStartPos, sString, Chr(34))
  fncGetAttrValueFromString = Mid(sString, lStartPos, lEndPos - lStartPos)
End Function


