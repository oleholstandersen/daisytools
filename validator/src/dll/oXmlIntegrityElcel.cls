VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oXmlIntegrityElcel"
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

Public Function documentValidatesToGivenDTD(iobjreport As oReport, _
  isAbsPath As String, isDTD As String) As Boolean
  
  documentValidatesToGivenDTD = False
  
  If bolUseOnlineDtds Then isDTD = ""
  
  If fncValidateDocument( _
    iobjreport, isAbsPath, isDTD, "xi.dtdGiven" _
    ) Then
        documentValidatesToGivenDTD = True
  Else
    objEvent.subLog ("failure in fncValidateDocument: " & _
        fncGetFileName(isAbsPath) & " " & isDTD)
  End If
End Function

Public Function documentValidatesToCustomDTD(iobjreport As oReport, _
  isAbsPath As String, isDTD As String) As Boolean
  
  documentValidatesToCustomDTD = False
  
  If fncValidateDocument( _
    iobjreport, isAbsPath, isDTD, "xi.dtdCustom" _
    ) Then
        documentValidatesToCustomDTD = True
  Else
    objEvent.subLog ("failure in fncValidateDocument: " & _
        fncGetFileName(isAbsPath) & " " & isDTD)
  End If
    
End Function

Private Function fncValidateDocument(iobjreport As oReport, _
  isAbsPath As String, isDTD As String, isObjectReporting As String) As Boolean
  
  Dim oFSO As Object, oFile As Object
  Dim sDtdSwitch As String, slocalProxySettings As String
  
  On Error GoTo ErrH
  
  fncValidateDocument = False
  
  'Set dtd flags
  'online+doctype: "" (use the doctype statement)
  'online+custom: daisyorgpath & d202smil.dtd
  'local+given: dtdPath & smil10.dtd
  'local+custom: dtdPath & d202smil.dtd
  
  If bolUseOnlineDtds And isDTD = "" Then sDtdSwitch = ""
  If bolUseOnlineDtds And (Not isDTD = "") Then _
    sDtdSwitch = "--dtd=http://www.daisy.org/publications/dtds/" & isDTD & " "
  If (Not bolUseOnlineDtds) Then sDtdSwitch = "--dtd=" & sDtdPath & isDTD & " "
   
  If bolUseProxy Then slocalProxySettings = "--httpproxy=" & _
    sProxy & " "
  
  On Error Resume Next
  'Create a batch file. This way we can provide flags to 'the xmlvalid.exe'
  'and redirect the ouput to the file 'output.txt'
  
  Set oFSO = CreateObject("scripting.filesystemobject")
  Set oFile = oFSO.CreateTextFile(sElcelTempPath & "valid.bat")
  oFile.Write sElcelExePath & "xmlvalid.exe " & sDtdSwitch & slocalProxySettings & _
    isAbsPath & " > " & sElcelTempPath & "output.txt"
  Set oFile = Nothing
  
  'If error creating file, exit
  If Not Err.Number = 0 Then Exit Function
  
  'Run the batchfile
  Shell sElcelTempPath & "valid.bat", vbHide
  
  'Since the 'shell' function is running in another thread, we'll have to wait until
  'the 'output.txt' file is present
  Do
    DoEvents
  Loop Until oFSO.fileexists(sElcelTempPath & "output.txt")

  'Allthough the 'Shell' thread has create the file, the data may not have been written.
Again:
  Set oFile = oFSO.GetFile(sElcelTempPath & "output.txt")
  DoEvents
  If oFile.Size = 0 Then GoTo Again
  
  'The 'output.txt' file now contains the output from 'xmlvalid.exe', we get it and
  'open it
  Set oFile = oFSO.OpenTextFile(sElcelTempPath & "output.txt", 1, -2)
  
  Dim tempsString As String, templValue As Long, templValue2 As Long
  Dim lColumn As Long, lLine As Long, tempsOut As String, sError As String
  
  Do
    'Read ONE LINE from the file
    tempsString = oFile.ReadLine
    
    'If that line contains the phrase 'is valid' then we don't care
    If InStr(1, tempsString, "is valid", vbTextCompare) = 0 Then
    
      'Here we extract line and column information that is reported by 'xmlvalid.exe'
      'in format '[', line, ':', column, ']'
      templValue = InStr(1, tempsString, "[", vbBinaryCompare)
      If templValue = 0 Then GoTo Skip
      templValue2 = InStr(templValue + 1, tempsString, ":", vbBinaryCompare)
      If templValue2 = 0 Then GoTo Skip
  
      lLine = CLng(Mid$(tempsString, templValue + 1, templValue2 - templValue - 1))
    
      templValue = InStr(templValue2 + 1, tempsString, "]", vbBinaryCompare)
      If templValue = 0 Then GoTo Skip
    
      lColumn = CLng(Mid$(tempsString, templValue2 + 1, templValue - templValue2 - 1))
        
Skip:
      ' After that comes a ':' and the error string
      templValue2 = InStr(templValue + 1, tempsString, ":", vbBinaryCompare)
      If templValue2 = 0 Then Exit Do
    
      sError = Mid$(tempsString, templValue2 + 2, Len(tempsString) - templValue2 - 1)
      
      ' report the error
      iobjreport.fncInsertFailedTest isObjectReporting, isAbsPath, lLine, lColumn, _
        sError
    End If
  Loop Until oFile.AtEndOfStream Or (Not Err.Number = 0)
  
  'This is crashprotection, setting variables to nothing can take some time
  Set oFile = Nothing
  Set oFSO = Nothing
  Do
    DoEvents
  Loop Until (oFile Is Nothing And oFSO Is Nothing)
  
  On Error Resume Next
Again2:
  Err.Clear
  
  'Delete the temporary files
  Kill sElcelTempPath & "output.txt"
  If Not Err.Number = 0 Then GoTo Again2
  Kill sElcelTempPath & "valid.bat"

  fncValidateDocument = True
ErrH:

End Function

