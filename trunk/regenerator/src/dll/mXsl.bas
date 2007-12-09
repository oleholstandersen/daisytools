Attribute VB_Name = "mXsl"
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

Public Function fncRunXslTransform( _
    ByRef oDocToTransForm As MSXML2.DOMDocument40, _
    ByVal sXsltPath As String _
) As Boolean

  'fncRunXslTransform = True: Exit Function

Dim oXsl As New MSXML2.DOMDocument40: oXsl.async = False
'
'sXsltPath = sAppPath & "xml_pretty_printer\xsl\xml_pp_clean.xsl"


  'If bDebugMode Then objowner.addlog "fncRunXslTransform in"
  fncRunXslTransform = False
  If fncParseFile(sXsltPath, oXsl) Then
    ' Parse results into byref DOM Document.
    Dim oDomOutput As New MSXML2.DOMDocument40
        oDomOutput.async = False
        oDomOutput.validateOnParse = False
        oDomOutput.resolveExternals = False
        oDomOutput.preserveWhiteSpace = True
        
    oDocToTransForm.transformNodeToObject oXsl, oDomOutput
    Set oDocToTransForm = oDomOutput
  Else
    objOwner.addlog fncGetFileName(sXsltPath) & " parse error"
    GoTo ErrHandler
  End If
  
  fncRunXslTransform = True
  'If bDebugMode Then objowner.addlog "fncRunXslTransform out"
ErrHandler:
  If Not fncRunXslTransform Then objOwner.addlog "fncRunXslTransform ErrH"
End Function
