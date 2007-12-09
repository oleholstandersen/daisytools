Attribute VB_Name = "mXmlSave"
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

Public Function fncSaveAllXmlInOutArray( _
    sFolderPath As String, _
    sOutCharsetName As String, _
    objOwner As oRegenerator _
    ) As Boolean
Dim i As Long, sDomData As String
Dim oSaveDom As New MSXML2.DOMDocument40
    oSaveDom.async = False
    oSaveDom.validateOnParse = False
    oSaveDom.resolveExternals = False
    oSaveDom.preserveWhiteSpace = False
    oSaveDom.setProperty "NewParser", True
    oSaveDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
  
  On Error GoTo ErrHandler
  fncSaveAllXmlInOutArray = False

  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    Select Case objOwner.objFileSetHandler.aOutFileSet(i).eType
      Case TYPE_NCC, TYPE_SMIL_1, TYPE_SMIL_CONTENT, TYPE_SMIL_MASTER
        sDomData = objOwner.objFileSetHandler.aOutFileSet(i).sDomData
        sDomData = fncFixProlog(sDomData, objOwner.objFileSetHandler.aOutFileSet(i).eType, sOutCharsetName, objOwner)
                
        'then save
        If LCase$(sOutCharsetName) = "shift_jis" Or LCase$(sOutCharsetName) = "big5" Then
          'saving with dom makes all shiftjis chars into numentities, dont know why. mg20030314
          If Not fncXmlSaveSax(sDomData, sFolderPath & objOwner.objFileSetHandler.aOutFileSet(i).sFileName, sOutCharsetName, objOwner.objFileSetHandler.aOutFileSet(i).eType, objOwner) Then GoTo ErrHandler
        Else
          If Not fncParseString(sDomData, oSaveDom, objOwner) Then GoTo ErrHandler
          Dim sSavePath As String
          sSavePath = sFolderPath & objOwner.objFileSetHandler.aOutFileSet(i).sFileName
          If fncFileExists(sSavePath, objOwner) Then
           'very unlikely, but the dest exists already
             Dim oFSO As Object, oFile As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")
             Dim sUniquePath As String
             
             Set oFile = oFSO.getFile(sSavePath)
             If fncFolderExists(objOwner.sUnrefPath, objOwner) Then
               'move it to unrefpath
               sUniquePath = objOwner.sUnrefPath & oFile.Name
               If oFSO.fileExists(sUniquePath) Then
                 Do
                   sUniquePath = sUniquePath & "_"
                 Loop Until Not oFSO.fileExists(sUniquePath)
               End If
               
             Else
               'rename in dtb dir only
               sUniquePath = oFile.Path
               Do
                 sUniquePath = sUniquePath & "_"
               Loop Until Not oFSO.fileExists(sUniquePath)
               
             End If
             oFile.Move sUniquePath
             objOwner.addlog "<warning in='fncSaveAllXmlInOutArray'>moved " & sSavePath & " to " & sUniquePath & " before saving new file with same name</warning>"
          End If
          
          If Not fncXmlSaveDom(oSaveDom, sSavePath, objOwner) Then GoTo ErrHandler
        End If
      Case Else
        'do nothing
    End Select
    DoEvents
  Next i

  fncSaveAllXmlInOutArray = True
ErrHandler:
  Set oSaveDom = Nothing
  If Not fncSaveAllXmlInOutArray Then
    'Stop
    objOwner.addlog "<errH in='fncSaveAllXmlInOutArray'>fncSaveAllXmlInOutArray ErrH</errH>"
  End If
End Function

Public Function fncXmlSaveDom( _
    ByRef oSaveDom As MSXML2.DOMDocument40, _
    ByVal sDestAbsPath As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean

  fncXmlSaveDom = False
  On Error GoTo ErrHandler

  oSaveDom.save (sDestAbsPath)
  fncXmlSaveDom = True

ErrHandler:
  If Not fncXmlSaveDom Then
  'Stop
  objOwner.addlog "<errH in='fncXmlSaveDom'>fncXmlSaveDom ErrH</errH>"
  End If
End Function

Public Function fncXmlSaveSax( _
    ByRef sXmlData As String, _
    ByVal sDestAbsPath As String, _
    ByVal sOutCharsetName As String, _
    lFileType As Long, _
    ByRef objOwner As oRegenerator _
    ) As Boolean

  fncXmlSaveSax = False
  On Error GoTo ErrHandler
      
  Dim rdr As New SAXXMLReader40
  Dim wrt As New MXXMLWriter40
  Dim sContent As String

  On Error GoTo ErrHandler
  fncXmlSaveSax = False
  Set rdr.contentHandler = wrt
  Set rdr.dtdHandler = wrt
  Set rdr.errorHandler = wrt
  rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt
  rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
  rdr.putFeature "preserve-system-identifiers", True
  wrt.output = ""
  wrt.byteOrderMark = False
  wrt.encoding = sOutCharsetName
  wrt.standalone = True
  wrt.indent = True
  wrt.omitXMLDeclaration = False
  rdr.parse sXmlData
  
  sXmlData = wrt.output
    
  fncTweakSaxProlog sXmlData, sOutCharsetName
    
  If Not fncSaveFile(sDestAbsPath, sXmlData, objOwner) Then GoTo ErrHandler
  
  fncXmlSaveSax = True
  
ErrHandler:
  If Not fncXmlSaveSax Then objOwner.addlog "<errH in='fncXmlSaveSax'>fncXmlSaveSax ErrH</errH>"
End Function



