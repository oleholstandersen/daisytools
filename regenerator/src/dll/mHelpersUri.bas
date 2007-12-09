Attribute VB_Name = "mHelpersUri"
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
  
'  Debug.Print "No lcase$ in fncParseUri -> Check if errors"
  
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


'Public Function fncParseURI(ByVal isHref As String, isDrive As String, isPath As String, _
'  isFileName As String, isID As String, Optional isDefDrive As Variant) As Boolean
'
'  fncParseURI = False
'
'  Dim templCounter As Long, sDefPath As String, sDefDrive As String, tempsString As String
'  Dim sDefFile As String
'
'  If Not IsMissing(isDefDrive) Then _
'    If Not fncParseURI(CStr(isDefDrive), sDefDrive, sDefPath, sDefFile, tempsString) Then _
'      Exit Function
'
'  If (Not Right$(isHref, 1) = "\") And (Not Right$(isHref, 1) = "/") Then
'    templCounter = InStrRev(isHref, "#", -1, vbBinaryCompare)
'    If templCounter <> 0 Then
 '     isID = Right$(isHref, Len(isHref) - templCounter)
 '     isHref = Left$(isHref, templCounter - 1)
 '   End If
 ' End If
 '
 ' 'Debug.Print "No lcase$ in fncParseUri -> Check if errors"
 '
 ' For templCounter = 1 To Len(isHref)
 '   If Mid$(isHref, templCounter, 1) = "/" Then Mid$(isHref, templCounter, 1) = "\"
 ' Next templCounter
 '
 ' If LCase$(Left$(isHref, 7)) = "file:\\" Then _
 '   isHref = Right$(isHref, Len(isHref) - 7)
 ' If LCase$(Left$(isHref, 7)) = "http:\\" Or Left$(isHref, 6) = "ftp:\\" Then _
 '   Exit Function
 ' If LCase$(Left$(isHref, 7)) = "mailto:" Then Exit Function
 '
 ' templCounter = InStr(1, isHref, ":", vbBinaryCompare)
 ' If templCounter <> 0 Then
 '   isDrive = Mid$(isHref, templCounter - 1, 2)
 '   isHref = Mid$(isHref, templCounter + 1, Len(isHref) - templCounter)
 ' Else
 '   isDrive = sDefDrive
'  End If
 '
 '' templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
 ' If templCounter <> 0 Then
 '   isPath = sDefPath & Left$(isHref, templCounter)
 '   isHref = Right$(isHref, Len(isHref) - templCounter)
 ' Else
'    isPath = sDefPath
'  End If
'
'  templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
'  If templCounter <> 0 And templCounter < Len(isHref) Then
'    isFileName = Right$(isHref, Len(isHref) - templCounter)
'    isHref = Left$(isHref, templCounter - 1)
'  ElseIf isHref = "" Then
'    isFileName = sDefFile
'  Else
'    isFileName = isHref
'  End If
'
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
'
'  fncParseURI = True
'End Function

Public Function fncIsValidUriChars(sCandidate As String) As Boolean
Dim i As Long
  fncIsValidUriChars = False
  For i = 1 To Len(sCandidate)
    If Not fncIsValidUriChar(Asc(Mid$(sCandidate, i, 1))) Then Exit Function
  Next i
  fncIsValidUriChars = True
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
    sString = Replace$(sString, Chr(229), Chr(97) & Chr(97), , , vbTextCompare)
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
