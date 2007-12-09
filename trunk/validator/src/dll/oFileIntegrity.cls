VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oFileIntegrity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Const sErrfFE As String = "Error in fncFileExists: "
Const sErrfFHVN As String = "Error in fncFileHasValidName: "

' This function checks if a file exists
Public Function fncFileExists( _
  ByRef iobjReport As oReport, ByVal isAbsPath As String, _
  Optional isContext As Variant, Optional iobjNode As Variant) As Boolean
    
    
    'fncInsertTime "oFileIntegrity.fncFileExists"
        
    fncFileExists = False
    
    Dim oFSO As Object, oFile As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If oFSO Is Nothing Then objEvent.subLog sErrfFE & _
      "couldn't create filesystemobject": GoTo ErrorH
    
    On Error GoTo ErrorH
    
    Set oFile = oFSO.GetFile(isAbsPath)
    
    iobjReport.subInsertSucceededTest

    Set oFSO = Nothing
    Set oFile = Nothing

    fncFileExists = True
    'fncInsertTime "oFileIntegrity.fncFileExists"
    Exit Function
ErrorH:
    Dim objNode As Object, sContext As String
    If Not IsMissing(iobjNode) Then Set objNode = iobjNode
    If Not IsMissing(isContext) Then sContext = isContext
    
    If Not fncFileExists Then _
      fncInsFail2Report iobjReport, objNode, "fi.exists", sContext, _
      "file doesn't exist: " & isAbsPath
            
    fncFileExists = True
    
    'fncInsertTime "oFileIntegrity.fncFileExists"
End Function

' This function checks if a file has a valid name given the criterias below.
'
Public Function fncFileHasValidName( _
  ByRef iobjReport As oReport, ByVal isFileName As String, eType As enuFileType _
  ) As Boolean
    
    'fncInsertTime "oFileIntegrity.fncFileHasValidName"
    
    fncFileHasValidName = True
  
    Dim oFSO As Object, oFile As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If oFSO Is Nothing Then objEvent.subLog sErrfFHVN & _
      "couldn't create filesystemobject": Exit Function

    Err.Clear
    On Error Resume Next
    Set oFile = oFSO.GetFile(isFileName)
    If Not Err.Number = 0 Then Err.Clear: fncFileHasValidName = True: GoTo ErrH
    
    On Error GoTo ErrH
    
    Dim sFileName As String
    sFileName = oFile.Name
    
    Set oFile = Nothing
    Set oFSO = Nothing

    If eType = ncc Then
        If Not ((sFileName = "ncc.html") Or (sFileName = "NCC.HTML")) Then
            iobjReport.fncInsertFailedTest "fi.hasValidName", isFileName, , , _
                "ncc.html or NCC.HTML"
            Exit Function
        Else
          iobjReport.subInsertSucceededTest
        End If
    ElseIf eType = mastersmil Then
        If (Not ((sFileName = "master.smil") Or (sFileName = "MASTER.SMIL"))) Then
            iobjReport.fncInsertFailedTest "fi.hasValidName", isFileName, , , _
                "master.smil or MASTER.SMIL"
            Exit Function
        Else
          iobjReport.subInsertSucceededTest
        End If
    ElseIf eType = smil Then
        'check extension, check basename URI consistency
        If (Not ((fncGetExtension(sFileName) = "SMIL") Or _
            (fncGetExtension(sFileName) = "smil")) Or _
            (Not fncIsValidUriChars(fncGetBaseName(sFileName)))) Then
                iobjReport.fncInsertFailedTest "fi.hasValidName", isFileName
            Exit Function
        Else
          iobjReport.subInsertSucceededTest
        End If
    ElseIf eType = smilMediaObText Then
        'check extension, check basename URI consistency
        If Not ((fncGetExtension(LCase(sFileName)) = "htm") Or _
            (fncGetExtension(LCase(sFileName)) = "html")) Or _
            (Not fncIsValidUriChars(fncGetBaseName(sFileName))) Then
                iobjReport.fncInsertFailedTest "fi.hasValidName", isFileName
            Exit Function
        Else
          iobjReport.subInsertSucceededTest
        End If
    ElseIf eType = discinfo Then
        If Not ((sFileName = "discinfo.html") Or _
            (sFileName = "DISCINFO.HTML")) Then
                iobjReport.fncInsertFailedTest "fi.hasValidName", isFileName, , , _
                "discinfo.html or DISCINFO.HTML"
            Exit Function
        Else
          iobjReport.subInsertSucceededTest
        End If
    Else 'all other objects are just checked for URI consistency:
    'smilMediaObAudio 'smilMediaObImg 'smilMediaObOther 'xhtmlExtEnt
            If Not fncIsValidUriChars(sFileName) Then
                iobjReport.fncInsertFailedTest "fi.hasRecommendedName", isFileName
                Exit Function
            Else
              iobjReport.subInsertSucceededTest
            End If
    End If
    
    fncFileHasValidName = True
ErrH:
  If Not fncFileHasValidName Then objEvent.subLog "error in oFileIntegrity.fncFileHasValidName on" & isFileName
  
  'fncInsertTime "oFileIntegrity.fncFileHasValidName"
End Function

' This function checks wheter a file is physically readable or not
'
Public Function fncFileIsReadable(ByRef iobjReport As oReport, ByVal isAbsPath As String, Optional ByRef isFileData As Variant) As Boolean
  Dim iFF As Integer, sTemp As String
  Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
  'isFileData is sent in as empty byref so this fnc is also 'readfile'
    
    'fncInsertTime "oFileIntegrity.fncFileIsReadable"
    
    fncFileIsReadable = False
    On Error GoTo ErrH
    
    If oFSO.FileExists(isAbsPath) Then                      'check that exists first, because it will be created by iFF if it doesnt
        iFF = FreeFile                                      'use iFF instead of fso since audiofiles are also tested
        Open isAbsPath For Binary As #iFF
        If LOF(iFF) = 0 Then Close #iFF: GoTo ErrH
        
        sTemp = Space$(LOF(iFF))
        Get #iFF, , sTemp
        Close #iFF
    End If
    
    If Not IsMissing(isFileData) Then isFileData = sTemp 'if isFileData has been sent as a byref parameter, then fill it.
     
ErrH:
    fncFileIsReadable = True
    If Not fncFileIsReadable Then
      iobjReport.fncInsertFailedTest "fi.isReadable", isAbsPath
    Else
      iobjReport.subInsertSucceededTest
    End If
    
    'fncInsertTime "oFileIntegrity.fncFileIsReadable"
End Function

' This function checks wheter the given URI contains valid characters or not
'
Private Function fncIsValidUriChars(ByVal sCandidate As String) As Boolean
Dim i As Long
    fncIsValidUriChars = False
    For i = 1 To Len(sCandidate)
        If Not fncIsValidUriChar(Asc(Mid$(sCandidate, i, 1))) Then Exit Function
    Next i
    fncIsValidUriChars = True
End Function

' This function checks wheter the given character is a valid URI character
'
Private Function fncIsValidUriChar(ByVal sChar As String) As Boolean
    fncIsValidUriChar = False
    If Not ((CInt(sChar) > 44 And CInt(sChar) < 58) Or _
        (CInt(sChar) > 63 And CInt(sChar) < 91) Or _
                (CInt(sChar) > 94 And CInt(sChar) < 123)) _
    Then Exit Function
    fncIsValidUriChar = True
End Function
