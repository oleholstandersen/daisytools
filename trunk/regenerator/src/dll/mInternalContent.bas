Attribute VB_Name = "mInternalContent"
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

Public Function fncInternalContent( _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByRef lCurrentArrayItem As Long, _
    ByRef bolPreserveBiblioMeta As Boolean, _
    ByRef bolAddCss As Boolean, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
    
Dim oContentDom As New MSXML2.DOMDocument40
    oContentDom.async = False
    oContentDom.validateOnParse = False
    oContentDom.resolveExternals = False
    oContentDom.preserveWhiteSpace = False
    oContentDom.setProperty "SelectionLanguage", "XPath"
    oContentDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
    oContentDom.setProperty "NewParser", True

    On Error GoTo ErrHandler
    fncInternalContent = False
      
    If Not fncParseString( _
        objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sDomData, _
        oContentDom, objOwner _
        ) Then GoTo ErrHandler
    
    If Not fncDoNonNccBiblioMetaData( _
        oContentDom, _
        TYPE_SMIL_CONTENT, _
        bolPreserveBiblioMeta, _
        objOwner _
        ) Then GoTo ErrHandler
        
   If Not fncFixFragmentCase(oContentDom, TYPE_SMIL_CONTENT, objOwner) Then GoTo ErrHandler
      
   'mg20030911: fix pagenum case (class attr value)
   If Not fncFixAttrValueCase(oContentDom, objOwner, TYPE_SMIL_CONTENT, lCurrentArrayItem) Then GoTo ErrHandler
      
   If Not fncFixPageNums(oContentDom, objOwner, lCurrentArrayItem) Then GoTo ErrHandler
   
   If Not fncSetLangAttrs(oContentDom, objOwner, lCurrentArrayItem) Then GoTo ErrHandler
   
   If bolAddCss Then
    If Not fncAddCss(oContentDom, objOwner) Then GoTo ErrHandler
   End If
   
   'move http-equiv elem to first sibling pos
   If Not fncMoveSiblingToTop(oContentDom, "//head/meta[@http-equiv='Content-type']", objOwner) Then GoTo ErrHandler

   ' mg20030325 do what tidy should be doing; also causes illegal char in output
   If Not fncStripEmptyElem(oContentDom, "//p", objOwner) Then GoTo ErrHandler

   'set the modded dom back to array
   objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sDomData = oContentDom.xml
   
  fncInternalContent = True

ErrHandler:
  Set oContentDom = Nothing
  DoEvents
  If Not fncInternalContent Then objOwner.addlog "<errH in='fncInternalContent' arrayItem='" & lCurrentArrayItem & "'>fncInternalContent ErrH" & Err.Number & " : " & Err.Description & "</errH>"
End Function
