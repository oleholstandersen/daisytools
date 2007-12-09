Attribute VB_Name = "mXmlLoad"
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

Public Function fncParseFile( _
    ByVal isAbsPath As String, _
    ByRef ioDom As MSXML2.DOMDocument40 _
    ) As Boolean
 
    fncParseFile = False
    
    If Not ioDom.Load(isAbsPath) Then
        addLog "Parse error in " & fncGetFileName(isAbsPath) & ": " & ioDom.parseError.reason & vbCrLf & _
           "filepos: " & ioDom.parseError.filepos & vbCrLf & _
           "line: " & ioDom.parseError.Line & vbCrLf & _
           "linepos: " & ioDom.parseError.linepos & vbCrLf & _
           "srctext: " & ioDom.parseError.srcText & vbCrLf & _
           "Url: " & ioDom.parseError.url & vbCrLf
    Else
        fncParseFile = True
    End If
    
End Function

Public Function fncParseString( _
    ByVal isContent As String, _
    ByRef ioDom As MSXML2.DOMDocument40 _
    ) As Boolean
    
    fncParseString = False

    If Not ioDom.loadXML(isContent) Then
        addLog "Parse error: " & ioDom.parseError.reason & vbCrLf & _
           "filepos: " & ioDom.parseError.filepos & vbCrLf & _
           "line: " & ioDom.parseError.Line & vbCrLf & _
           "linepos: " & ioDom.parseError.linepos & vbCrLf & _
           "srctext: " & ioDom.parseError.srcText & vbCrLf & _
           "Url: " & ioDom.parseError.url '& vbCrLf & vbCrLf & _
           'isContent
    Else
        fncParseString = True
    End If
End Function
