Attribute VB_Name = "mXmlValidate"
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

Public Function fncDtdValidateString(sDomData As String, sLocalName As String)
Dim i As Long
Dim oDom As New MSXML2.DOMDocument40
    oDom.async = False
    oDom.validateOnParse = True
    oDom.resolveExternals = True
    oDom.preserveWhiteSpace = True
    
    If fncParseString(sDomData, oDom) Then objOwner.addlog sLocalName & " valid."
    
    Set oDom = Nothing
    
End Function



