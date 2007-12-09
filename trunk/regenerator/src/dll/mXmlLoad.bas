Attribute VB_Name = "mXmlLoad"
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

Public Function fncParseFile( _
    ByVal isAbsPath As String, _
    ByRef ioDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
    
    fncParseFile = False
    
    If Not ioDom.Load(isAbsPath) Then
        objOwner.addlog "<error>Parse error in " & fncGetFileName(isAbsPath) & ": " & ioDom.parseError.reason & vbCrLf & _
           "filepos: " & ioDom.parseError.filepos & vbCrLf & _
           "line: " & ioDom.parseError.line & vbCrLf & _
           "linepos: " & ioDom.parseError.linepos & vbCrLf & _
           "srctext: " & ioDom.parseError.srcText & vbCrLf & _
           "Url: " & ioDom.parseError.url & vbCrLf & "</error>"
    Else
        fncParseFile = True
    End If
    
End Function

Public Function fncParseString( _
    ByVal isContent As String, _
    ByRef ioDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
    
    fncParseString = False

    If Not ioDom.loadXML(isContent) Then
        objOwner.addlog "<error>Parse error: " & ioDom.parseError.reason & vbCrLf & _
           "filepos: " & ioDom.parseError.filepos & vbCrLf & _
           "line: " & ioDom.parseError.line & vbCrLf & _
           "linepos: " & ioDom.parseError.linepos & vbCrLf & _
           "srctext: " & ioDom.parseError.srcText & vbCrLf & _
           "Url: " & ioDom.parseError.url & vbCrLf & vbCrLf & "</error>"
    Else
        fncParseString = True
    End If
End Function

Public Function fncParseStringSax( _
    ByRef sXmlData As String, _
    ByRef objOwner As oRegenerator, _
    Optional ByRef returnDom As MSXML2.DOMDocument40 _
    ) As Boolean

  fncParseStringSax = False
  On Error GoTo ErrHandler
   
  Dim rdr As New SAXXMLReader40
  Dim wrt As New MXXMLWriter40

  On Error GoTo ErrHandler
  fncParseStringSax = False
  Set rdr.contentHandler = wrt
  Set rdr.dtdHandler = wrt
  Set rdr.errorHandler = wrt
  rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt
  rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
  rdr.putFeature "preserve-system-identifiers", True
  wrt.output = ""
  wrt.byteOrderMark = False
  'wrt.encoding = sOutCharsetName
  wrt.standalone = True
  wrt.indent = True
  wrt.omitXMLDeclaration = False
  rdr.parse sXmlData
  sXmlData = wrt.output
  
  If Not fncParseString(sXmlData, returnDom, objOwner) Then GoTo ErrHandler
  
  fncParseStringSax = True

  
ErrHandler:
  
  If Not fncParseStringSax Then objOwner.addlog "<errH in='fncParseStringSax'>fncParseStringSax ErrH</errH>"
End Function

