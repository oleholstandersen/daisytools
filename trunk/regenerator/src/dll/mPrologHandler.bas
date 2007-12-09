Attribute VB_Name = "mPrologHandler"
' Daisy 2.02 Regenerator DLL
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

Private bHasSkipExtension As Boolean

Public Function fncFixProlog( _
    ByRef sRawFile As String, _
    ByVal lFileType As Long, _
    ByVal sOutCharsetName As String, _
    ByRef objOwner As oRegenerator _
    ) As String
  fncFixProlog = fncAddProlog(fncStripProlog(sRawFile, lFileType, objOwner), lFileType, sOutCharsetName, objOwner)
End Function

Private Function fncStripProlog( _
    sInput As String, _
    lFileType As Long, _
    objOwner As oRegenerator _
    ) As String
Dim sRootElemName As String

    If lFileType = TYPE_NCC Or lFileType = TYPE_SMIL_CONTENT Then
        sRootElemName = "<html"
        'check if the prolog is extended
        If lFileType = TYPE_SMIL_CONTENT Then
          bHasSkipExtension = bDTDIsExtendedInternallyForSkippableStructures(sInput)
        End If
    ElseIf lFileType = TYPE_SMIL_1 Or lFileType = TYPE_SMIL_MASTER Then
        sRootElemName = "<smil"
    Else
        'this shouldnt happen
        sRootElemName = ""
        objOwner.addlog "<error in='fncStripProlog'>no root element name found</error>"
    End If
    
    If InStr(1, sInput, sRootElemName, vbTextCompare) > 0 Then
        fncStripProlog = Mid(sInput, (InStr(1, sInput, sRootElemName, vbTextCompare)))
    Else
        fncStripProlog = sInput
    End If
    
End Function

Private Function fncAddProlog( _
    sRawFile As String, _
    lFileType As Long, _
    sOutCharsetName As String, _
    objOwner As oRegenerator _
    ) As String

Dim sProcIns As String: sProcIns = ""
Dim sDocType As String: sDocType = ""
Dim sDtdExtension As String
  sProcIns = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & sOutCharsetName & Chr(34) & " ?>" & vbCrLf

  If lFileType = (TYPE_SMIL_1) Or lFileType = (TYPE_SMIL_MASTER) Then
    sDocType = "<!DOCTYPE smil PUBLIC " & Chr(34) & _
           "-//W3C//DTD SMIL 1.0//EN" & Chr(34) & vbCrLf & vbTab & _
           Chr(34) & "http://www.w3.org/TR/REC-smil/SMIL10.dtd" & Chr(34) & ">"
  ElseIf (lFileType = TYPE_NCC) Or (lFileType = TYPE_SMIL_CONTENT) Then
    If bHasSkipExtension Then
      sDtdExtension = vbCrLf & "[" & vbCrLf & "<!ATTLIST span bodyref CDATA #IMPLIED>" & vbCrLf & "]"
    Else
      sDtdExtension = ""
    End If
    
    sDocType = "<!DOCTYPE html PUBLIC " & Chr(34) & _
           "-//W3C//DTD XHTML 1.0 Transitional//EN" & Chr(34) & vbCrLf & vbTab & _
           Chr(34) & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" & Chr(34) & sDtdExtension & ">" & vbCrLf
  Else
    objOwner.addlog "<error in='fncAddProlog'>erroneous lFileType encountered in fncAddProlog</error>"
  End If
  
  fncAddProlog = sProcIns & sDocType & sRawFile
End Function

Public Function fncTweakSaxProlog( _
    ByRef sXml As String, _
    sOutCharsetName As String _
    )
  sXml = Replace(sXml, "standalone=" & Chr(34) & "yes" & Chr(34), "encoding=" & Chr(34) & sOutCharsetName & Chr(34), 1, 1, vbTextCompare)
  sXml = Replace(sXml, " [" & vbCrLf & "]", "")
End Function

Private Function bDTDIsExtendedInternallyForSkippableStructures( _
  ByRef sXml As String) As Boolean
  Dim lTest1 As Long, lTest2 As Long, lTest3 As Long, lTest4 As Long, lTest5 As Long
  
  On Error GoTo ErrH
  
  bDTDIsExtendedInternallyForSkippableStructures = False
  'check that all strings of internal declaration are there
  lTest1 = InStr(1, sXml, "ATTLIST", vbBinaryCompare): If lTest1 < 1 Then Exit Function
  lTest2 = InStr(1, sXml, "bodyref", vbBinaryCompare): If lTest2 < 1 Then Exit Function
  lTest3 = InStr(1, sXml, "CDATA", vbBinaryCompare): If lTest3 < 1 Then Exit Function
  lTest4 = InStr(1, sXml, "#IMPLIED", vbBinaryCompare): If lTest4 < 1 Then Exit Function
  'also check that all these string occur before root
  lTest5 = InStr(1, sXml, "<html", vbBinaryCompare)
  If (lTest1 > lTest5) Or (lTest2 > lTest5) Or (lTest3 > lTest5) Or (lTest4 > lTest5) Then Exit Function
  'appears to be an extended DTD
  bDTDIsExtendedInternallyForSkippableStructures = True
 
ErrH:
End Function
