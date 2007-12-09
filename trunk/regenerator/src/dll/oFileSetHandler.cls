VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oFileSetHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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

Public objOwner As oRegenerator

Private aprivInFileSet() As oFileSetCollection
Private aprivInFileSetMembers As Long

Private aprivOutFileSet() As oFileSetCollection
Private aprivOutFileSetMembers As Long

Private aprivOutFileSetContentDocs() As Long '?? ?? is this used?
Private aprivOutFileSetContentDocMembers As Long '?? is this used?

Public Property Get aInFileSet(lItem As Long) As oFileSetCollection
  If lItem < 0 Or lItem > aprivInFileSetMembers - 1 Then Exit Property
  Set aInFileSet = aprivInFileSet(lItem)
End Property

Public Property Get aInFileSetMembers() As Long
  aInFileSetMembers = aprivInFileSetMembers
End Property

Public Property Get aOutFileSet(lItem As Long) As oFileSetCollection
  If lItem < 0 Or lItem > aprivOutFileSetMembers - 1 Then Exit Property
  Set aOutFileSet = aprivOutFileSet(lItem)
End Property

Public Property Get aOutFileSetMembers() As Long
  aOutFileSetMembers = aprivOutFileSetMembers
End Property

Public Property Get aOutFileSetContentDocs(lItem As Long) As Long
  If lItem < 0 Or lItem > aOutFileSetContentDocMembers Then Exit Property
  aOutFileSetContentDocs = aprivOutFileSetContentDocs(lItem)
End Property

Public Property Get aOutFileSetContentDocMembers() As Long
  aOutFileSetContentDocMembers = aprivOutFileSetContentDocMembers
End Property

'***************************************************************
'fncCreateInputFileSetArray takes the xml string data of the ncc;
'parses the ncc, gets smilfiles, adds them to arrays
'parses the smilfiles, gets mediaobjects, adds them to arrays
'for ncc and content doc, checks for auxilliary files, adds to array
'if there is a mastersmil it ignores it (always renders a new one)
'***************************************************************

' enuInputEncoding

Public Function fncCreateFileSetArrays( _
    ByRef oNccDom As MSXML2.DOMDocument40, _
    ByVal lInputEncoding As Long, _
    ByVal sOutCharsetName As String _
    ) As Boolean
Dim oSmilDom As New MSXML2.DOMDocument40
    oSmilDom.async = False
    oSmilDom.validateOnParse = False
    oSmilDom.resolveExternals = False
    oSmilDom.preserveWhiteSpace = True
    oSmilDom.setProperty "NewParser", True
Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
  
  objOwner.addlog "<status>tidying and parsing fileset members...</status>"
  fncCreateFileSetArrays = False
  On Error GoTo ErrHandler
  
  'extract all smilfiles from ncc:
    
   Set oNodes = oNccDom.selectNodes("//h1/a" & _
                                  "| //h2/a" & _
                                  "| //h3/a" & _
                                  "| //h4/a" & _
                                  "| //h5/a" & _
                                  "| //h6/a" & _
                                  "| //span/a" & _
                                  "| //div/a") '!xht

    If (Not oNodes.length = 0) Or (Not oNodes Is Nothing) Then
      Dim sCurrSrc As String
      If Not fncNccSmilSequenceIsCorrect(oNodes) Then GoTo ErrHandler
      'for each href value in ncc
      For Each oNode In oNodes
        sCurrSrc = fncStripIdAddPath(oNode.selectSingleNode("@href").nodeValue, objOwner.sDtbFolderPath) 'get the filename from the nodevalue
        If fncFileExists(sCurrSrc, objOwner) Then
        'add references in input array to the original smil files
          If Not fncIsObjectInInputArray(sCurrSrc) Then
            fncAddObjectToInputArray TYPE_SMIL_1, sCurrSrc
            'try to tidy the smilfile, if ok add to output array
            Dim sSmil As String

            If TidyLib.fncRunTidy(sCurrSrc, sSmil, lInputEncoding, TYPE_SMIL_1, sOutCharsetName, objOwner) Then
              If sSmil <> "" Then
                If fncParseString(sSmil, oSmilDom, objOwner) Then
                  If Not fncIsObjectInOutputArray(fncGetFileName(sCurrSrc)) Then
                    fncAddObjectToOutputArray TYPE_SMIL_1, fncGetFileName(sCurrSrc), oSmilDom.xml, oNode.Text
                  End If
                  'while the smil is parsed, extract its media objects
                  If Not fncExtractSmilMediaObjects(oSmilDom, sCurrSrc, lInputEncoding, sOutCharsetName) Then
                    objOwner.addlog ("<error>mediaobject extraction fail on " & sCurrSrc & "</error>")
                    GoTo ErrHandler
                  End If
                Else
                  objOwner.addlog "<error>could not parse tidy processed smil " & fncGetFileName(sCurrSrc) & "</error>"
                End If 'fncParseString(sSmil, oSmilDom)
                'Set oSmilDom = Nothing
              Else
                objOwner.addlog "<error>tidy process on input smil " & fncGetFileName(sCurrSrc) & " returned null</error>"
                GoTo ErrHandler
              End If 'sSmil <> ""
            End If 'fncRunTidy(sSmil)
          End If 'If Not fncIsObjectInInputArray
        Else
          objOwner.addlog ("<error>smilfile " & fncGetFileName(sCurrSrc) & " referenced in ncc not found</error>")
          GoTo ErrHandler
        End If 'fncFileExists(sCurrSrc)
        DoEvents
      Next
    Else
      objOwner.addlog "<error>0 smil URI nodes found in ncc</error>"
      GoTo ErrHandler
    End If ' Not oNodes.length = 0
    
    'While ncc(oNccDom) is parsed, extract its auxilliary files
    If Not fncExtractXhtmlAuxilliary(oNccDom, TYPE_NCC) Then GoTo ErrHandler
    
    fncCreateFileSetArrays = True
    objOwner.addlog "<status>input fileset: " & aInFileSetMembers & " members.</status>"
    
    If Not fncCreateContentDocMembers Then GoTo ErrHandler
ErrHandler:
    Set oSmilDom = Nothing
    If Not fncCreateFileSetArrays Then objOwner.addlog "<errH in='fncCreateFileSetArray'>fncCreateFileSetArray ErrH</errH>"
End Function

Private Function fncExtractSmilMediaObjects( _
    oSmilDom As MSXML2.DOMDocument40, _
    sCurrSrc As String, _
    lInputEncoding As Long, _
    ByVal sOutCharsetName As String _
    ) As Boolean

  fncExtractSmilMediaObjects = False
  If Not fncExtractSmilMediaObjectType(oSmilDom, "//text/@src", sCurrSrc, lInputEncoding, TYPE_SMIL_CONTENT, sOutCharsetName) Then GoTo Break
  If Not fncExtractSmilMediaObjectType(oSmilDom, "//audio/@src", sCurrSrc, lInputEncoding, TYPE_SMIL_AUDIO, sOutCharsetName) Then GoTo Break
  If Not fncExtractSmilMediaObjectType(oSmilDom, "//img/@src", sCurrSrc, lInputEncoding, TYPE_SMIL_IMG, sOutCharsetName) Then GoTo Break
  fncExtractSmilMediaObjects = True
  
Break:
  
End Function

Private Function fncExtractSmilMediaObjectType( _
    ByRef oSmilDom As MSXML2.DOMDocument40, _
    ByVal sXpath As String, _
    ByVal sCurrSmil As String, _
    ByVal lInputEncoding As Long, _
    ByVal lFileType As Long, _
    ByVal sOutCharsetName As String _
    ) As Boolean

Dim i As Long
Dim sCurrSrc As String
Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode, bolResult As Boolean

  On Error GoTo ErrHandler
  
  fncExtractSmilMediaObjectType = False
  
  Set oNodes = oSmilDom.selectNodes(sXpath)
    If Not (oNodes.length = 0) Or (oNodes Is Nothing) Then
      For Each oNode In oNodes
        sCurrSrc = fncStripIdAddPath(oNode.nodeValue, objOwner.sDtbFolderPath) 'get the filename from the nodevalue
        'add references in input array to the original files
        If Not fncIsObjectInInputArray(sCurrSrc) Then
          If Not fncAddObjectToInputArray(lFileType, sCurrSrc) Then
            'objOwner.addlog "<error>File missing in fncExtractSmilMediaObjectType: " & sCurrSrc & "</error>"
            'Exit Function
            GoTo SkipObject
          End If
          If lFileType = TYPE_SMIL_CONTENT Then
            'try to tidy the contentdoc, if ok add to output array
            Dim sContent As String
            If TidyLib.fncRunTidy(sCurrSrc, sContent, lInputEncoding, TYPE_SMIL_CONTENT, sOutCharsetName, objOwner) Then
              If sContent <> "" Then
                Dim oContentDom As New MSXML2.DOMDocument40
                    oContentDom.async = False
                    oContentDom.validateOnParse = False
                    oContentDom.resolveExternals = False
                    oContentDom.preserveWhiteSpace = True
                    oContentDom.setProperty "SelectionLanguage", "XPath"
                    oContentDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
                    oContentDom.setProperty "NewParser", True
                If fncParseString(sContent, oContentDom, objOwner) Then
                  If Not fncIsObjectInOutputArray(fncGetFileName(sCurrSrc)) Then
                    fncAddObjectToOutputArray lFileType, fncGetFileName(sCurrSrc), oContentDom.xml
                  End If
                  'While content is parsed, extract its auxilliary files
                  If Not fncExtractXhtmlAuxilliary(oContentDom, TYPE_SMIL_CONTENT) Then GoTo ErrHandler
                Else
                  objOwner.addlog "<error>could not parse tidy processed content doc " & fncGetFileName(sCurrSrc) & ": " & vbCrLf & sContent & "</error>"
                  GoTo ErrHandler
                  'revisit need to remove the input array item here?:
                  'aprivInFileSetMembers = aprivInFileSetMembers - 1
                End If 'fncParseString(sContent, oContentDom)
                Set oContentDom = Nothing
              Else
                objOwner.addlog "<error>tidy process on input smil " & fncGetFileName(sCurrSrc) & " returned null</error>"
                GoTo ErrHandler
              End If 'sContent <> ""
            End If 'fncRunTidy(sContent)
          Else 'if other type than content doc (= audio or img)
            If InStr(1, sCurrSrc, "rgn_empty", vbTextCompare) > 0 Then
              'this book has been regenerated before,
              'this is an "rgn_empty" audio element, change its lFileType
              fncAddObjectToOutputArray TYPE_SMIL_AUDIO_INSERTED, fncGetFileName(sCurrSrc), ""
            Else
              fncAddObjectToOutputArray lFileType, fncGetFileName(sCurrSrc), ""
            End If
          End If 'If lFileType = smil_content Then
SkipObject:
        End If 'If Not fncIsObjectInInputArray
        DoEvents
      Next
    Else

    End If ' Not oNodes.length = 0
  fncExtractSmilMediaObjectType = True
  
ErrHandler:
  If Not fncExtractSmilMediaObjectType Then objOwner.addlog "<errH in='fncExtractSmilMediaObjectType'>fncExtractSmilMediaObjectType ErrHandler</errH>"
  
End Function

Private Function fncIsObjectInInputArray( _
  sAbsPath As String _
  ) As Boolean
Dim i As Long
  
    fncIsObjectInInputArray = False
    For i = 0 To aInFileSetMembers - 1
        If LCase$(aInFileSet(i).sAbsPath) = LCase$(sAbsPath) Then
            fncIsObjectInInputArray = True
            Exit For
        End If
    Next i

End Function

Public Function fncIsObjectInOutputArray( _
  sFileName As String _
  ) As Boolean
Dim i As Long
  
    fncIsObjectInOutputArray = False
    For i = 0 To aOutFileSetMembers - 1
        If LCase$(aOutFileSet(i).sFileName) = LCase$(sFileName) Then
            fncIsObjectInOutputArray = True
            Exit For
        End If
    Next i

End Function

'Public Function fncAddObjectToInputArrayOld( _
'  ByVal lType As Long, _
'  ByVal sAbsPath As String _
'  ) As Boolean
'        On Error GoTo ErrHandler
'        fncAddObjectToInputArray = False
'        If fncFileExists(sAbsPath, objOwner) Then
'            ReDim Preserve aprivInFileSet(aInFileSetMembers)
'            Set aprivInFileSet(aInFileSetMembers) = New oFileSetCollection
'            With aprivInFileSet(aInFileSetMembers)
'              .sAbsPath = sAbsPath
'              .eType = lType
'            End With
'            aprivInFileSetMembers = aprivInFileSetMembers + 1
'            fncAddObjectToInputArray = True
'        Else
'            objOwner.addlog "<error in='fncAddObjectToInputArray'> " & sAbsPath & " not found.</error>"
'            Exit Function 'this exit added mg 20030219
'            'GoTo ErrHandler
'        End If
'ErrHandler:
'        If Not fncAddObjectToInputArray Then objOwner.addlog "<errH in='fncAddObjectToInputArray'>fncAddObjectToInputArray ErrH</errH>"
'End Function

Public Function fncAddObjectToInputArray( _
  ByVal lType As Long, _
  ByVal sAbsPath As String _
  ) As Boolean
  Dim oFSO As Object
  Dim oFile As Object
  Set oFSO = CreateObject("Scripting.FileSystemObject")
        'mg20050418, changed from fncAddObjectToInputArrayOld
        On Error GoTo ErrHandler
        fncAddObjectToInputArray = False
        If oFSO.FileExists(sAbsPath) Then
            Set oFile = oFSO.getFile(sAbsPath)
            sAbsPath = oFile.Path
                    
            ReDim Preserve aprivInFileSet(aInFileSetMembers)
            Set aprivInFileSet(aInFileSetMembers) = New oFileSetCollection
            With aprivInFileSet(aInFileSetMembers)
              .sAbsPath = sAbsPath
              .eType = lType
            End With
            aprivInFileSetMembers = aprivInFileSetMembers + 1
            fncAddObjectToInputArray = True
        Else
            objOwner.addlog "<error in='fncAddObjectToInputArray'> " & sAbsPath & " not found.</error>"
            Exit Function 'this exit added mg 20030219
            'GoTo ErrHandler
        End If
ErrHandler:
        If Not fncAddObjectToInputArray Then objOwner.addlog "<errH in='fncAddObjectToInputArray'>fncAddObjectToInputArray ErrH</errH>"
End Function


Public Function fncAddObjectToOutputArray( _
  ByVal lType As Long, _
  ByVal sFileName As String, _
  ByVal sDomData As String, _
  Optional ByVal sSmilTitle As String, _
  Optional ByVal lOwnerType As Long _
  ) As Boolean
        
        On Error GoTo ErrHandler
        
        'mg20031022, lcase smil extensions
        If lType = TYPE_SMIL_1 Then
          sFileName = Replace$(sFileName, ".SMIL", ".smil", , , vbTextCompare)
        End If
        
        If lType = TYPE_SMIL_AUDIO Then
          sFileName = Replace$(sFileName, ".WAV", ".wav", , , vbTextCompare)
          sFileName = Replace$(sFileName, ".MP2", ".mp2", , , vbTextCompare)
          sFileName = Replace$(sFileName, ".MP3", ".mp3", , , vbTextCompare)
        End If
        
        fncAddObjectToOutputArray = False
            ReDim Preserve aprivOutFileSet(aOutFileSetMembers)
            Set aprivOutFileSet(aOutFileSetMembers) = New oFileSetCollection
            With aprivOutFileSet(aOutFileSetMembers)
              .eType = lType
              .sDomData = sDomData
              .sFileName = sFileName
              .sSmilTitle = fncEscapeQuotes(sSmilTitle) 'this only used for smil members(indicated heading name in ncc)
              .lOwnerType = lOwnerType
            End With
            aprivOutFileSetMembers = aprivOutFileSetMembers + 1
        fncAddObjectToOutputArray = True
ErrHandler:
        If Not fncAddObjectToOutputArray Then objOwner.addlog "<errH in='fncAddObjectToOutputArray'>fncAddObjectToOutputArray error</errH>"
End Function

Public Function fncResetArrays() As Boolean
    fncResetArrays = False
    aprivInFileSetMembers = 0: ReDim Preserve aprivInFileSet(aprivInFileSetMembers)
    aprivOutFileSetMembers = 0: ReDim Preserve aprivOutFileSet(aprivOutFileSetMembers)
    aprivOutFileSetContentDocMembers = 0: ReDim Preserve aprivOutFileSetContentDocs(aprivOutFileSetContentDocMembers)
    fncResetArrays = True
End Function

Private Function fncExtractXhtmlAuxilliary( _
    ByRef oXhtDom As MSXML2.DOMDocument40, _
    ByVal lOwnerType As Long _
    ) As Boolean

Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode

  fncExtractXhtmlAuxilliary = False
  On Error GoTo ErrHandler
  
  'extract all aux types in xhtml:
  Set oNodes = oXhtDom.selectNodes( _
    "//head/link[@rel='stylesheet' and @type='text/css']/@href | //img/@src") '!xht
    '//REVISIT what about relative paths in URIs (here and everywhere)
    
  If Not (oNodes.length = 0) Or (oNodes Is Nothing) Then
    For Each oNode In oNodes
      Dim sCurrSrc As String
      ' mg20030316
      'old:
      'sCurrSrc = fncStripIdAddPath(oNode.nodeValue, objOwner.sDtbFolderPath)
      'new: this handles relative paths "../bla.jpg" in these values
      sCurrSrc = ""
      'If Not fncGetAbsolutePathName(objOwner.sDtbFolderPath & oNode.nodeValue, sCurrSrc) Then GoTo ErrHandler
      sCurrSrc = fncStripIdAddPath(oNode.nodeValue, objOwner.sDtbFolderPath)
      
      If fncFileExists(LCase$(sCurrSrc), objOwner) Then
        If Not fncIsObjectInInputArray(sCurrSrc) Then
          fncAddObjectToInputArray TYPE_OTHER, LCase$(sCurrSrc)
          fncAddObjectToOutputArray TYPE_OTHER, LCase$(fncGetFileName(oNode.nodeValue)), "", "", lOwnerType
          'mg20050310: parse the css and look for url() objects
          If LCase$(fncGetExtensionFromString(sCurrSrc)) = "css" Then
            If Not fncExtractCssAuxilliary(sCurrSrc) Then GoTo ErrHandler
          End If
        End If
      Else
        objOwner.addlog "<warning in='fncExtractXhtmlAuxilliary'>warning: aux file " & oNode.nodeValue & " not found on filesystem</warning>"
      End If
      DoEvents
    Next
  Else
    
  End If 'Not (oNodes.length = 0) Or (oNodes Is Nothing)
  
  fncExtractXhtmlAuxilliary = True
  
ErrHandler:
  If Not fncExtractXhtmlAuxilliary Then objOwner.addlog "<errH in='fncExtractXhtmlAuxilliary'>fncExtractXhtmlAuxilliary ErrH</errH>"
End Function

Private Function fncExtractCssAuxilliary(ByRef sCurrSrc As String) As Boolean
Dim sCssUrlRefs() As String
Dim sCssUrlRefFullPath As String
  On Error GoTo ErrHandler
  fncExtractCssAuxilliary = False
  'check css for url() entries
  If (fncGetCssUrlValues(sCurrSrc, sCssUrlRefs)) Then
    Dim i As Long
    For i = 1 To UBound(sCssUrlRefs)
       sCssUrlRefFullPath = fncGetPathName(sCurrSrc) & sCssUrlRefs(i)
      'check integrity and add
      If fncFileExists(LCase$(sCssUrlRefFullPath), objOwner) Then
        If Not fncIsObjectInInputArray(sCssUrlRefFullPath) Then
          fncAddObjectToInputArray TYPE_OTHER, LCase$(sCssUrlRefFullPath)
          fncAddObjectToOutputArray TYPE_OTHER, LCase$(fncGetFileName(sCssUrlRefFullPath)), "", ""
        End If
      Else
        objOwner.addlog "<warning in='fncExtractCssAuxilliary'>warning: File referenced by CSS not found on filesystem: " & sCssUrlRefFullPath & "</warning>"
      End If '.fncFileExists
    Next i
  Else
    'no url values found, or error
  End If
  fncExtractCssAuxilliary = True
ErrHandler:
End Function

Private Function fncGetCssUrlValues(ByRef sCurrSrc As String, ByRef sCssUrlRefs() As String) As Boolean
Dim sCssFile As String
Dim bUrlFound As Boolean
Dim lCurrentUrlPos As Long
Dim lUrlOpenParen As Long
Dim lUrlCloseParen As Long
Dim sUrlString As String
Dim sCssUrlRefsCount As Long
'populate the byref input array with any found url() values
'return false on error or if array length = 0
  fncGetCssUrlValues = False
  On Error GoTo ErrHandler
  
  sCssFile = fncGetFileAsString(sCurrSrc)
  If sCssFile <> "" Then
    lCurrentUrlPos = 1
    Do
      bUrlFound = False
      lCurrentUrlPos = InStr(lCurrentUrlPos, sCssFile, "url", vbTextCompare)
      If lCurrentUrlPos > 0 Then
        bUrlFound = True
        lUrlOpenParen = InStr(lCurrentUrlPos, sCssFile, "(", vbTextCompare)
        lUrlCloseParen = InStr(lCurrentUrlPos, sCssFile, ")", vbTextCompare)
        'get the string within the paren
        sUrlString = Mid(sCssFile, lUrlOpenParen + 1, lUrlCloseParen - (lUrlOpenParen + 1))
        'Stop
        'remove single or double quotes
        sUrlString = Trim$(sUrlString)
        If (Mid(sUrlString, 1, 1) = Chr(39)) Or (Mid(sUrlString, 1, 1) = Chr(34)) Then
          'remove the leading char
          sUrlString = Mid(sUrlString, 2, Len(sUrlString))
        End If
        
        If (Mid(sUrlString, Len(sUrlString), 1) = Chr(39)) Or (Mid(sUrlString, Len(sUrlString), 1) = Chr(34)) Then
          'remove the trailing char
          sUrlString = Mid(sUrlString, 1, Len(sUrlString) - 1)
        End If
        
        'add it to array
        sCssUrlRefsCount = sCssUrlRefsCount + 1
        ReDim Preserve sCssUrlRefs(sCssUrlRefsCount)
        sCssUrlRefs(sCssUrlRefsCount) = sUrlString
        'indicate there is content on the array
        fncGetCssUrlValues = True
        'move cursor forward
        lCurrentUrlPos = lUrlCloseParen
      End If
    Loop Until bUrlFound = False
  
  End If 'sCssFile <> ""
  
  Exit Function
ErrHandler:
  fncGetCssUrlValues = False
End Function

Public Function fncCheckArrayConsistency() As Boolean
Dim i As Long
  
  On Error GoTo ErrHandler
  fncCheckArrayConsistency = False
  
  'check that both arrays has equal number of members (mastersmil not created yet)
  If aInFileSetMembers <> aOutFileSetMembers Then
    objOwner.addlog "<error in='fncCheckArrayConsistency'>differing member count in input and output arrays</error>"
    Exit Function
  End If
  
  'check that placement corresponds, exclude ncc, since it is always namechanged
  For i = 1 To aInFileSetMembers - 1
   'mg20030315 added lcase below
   If LCase$(fncGetFileName(aInFileSet(i).sAbsPath)) <> LCase$(aOutFileSet(i).sFileName) Then
     objOwner.addlog "<error in='fncCheckArrayConsistency' arrayItem='" & CStr(i) & "'>" & fncGetFileName(aInFileSet(i).sAbsPath) & " and " & aOutFileSet(i).sFileName & "is not the same. Array item number " & i & "."
     Exit Function
   End If
   DoEvents
  Next i

  ' check that there are no dupes
  ' this is run from within fncNamesAndUris so save time by disabling it here
  'If Not fncAllFileNamesAreUnique() Then GoTo ErrHandler

  fncCheckArrayConsistency = True
ErrHandler:
  If Not fncCheckArrayConsistency Then objOwner.addlog "<errH in='fncCheckArrayConsistency'>fncCheckArrayConsistency ErrH</errH>"
End Function

Public Function fncAllFileNamesAreUnique() As Boolean
Dim i As Long, lCount As Long, k As Long
Dim sTest As String
    
  On Error GoTo ErrHandler
  fncAllFileNamesAreUnique = False
  
  For i = 1 To UBound(aprivOutFileSet)
    sTest = aprivOutFileSet(i).sFileName
    lCount = 0
    For k = 1 To UBound(aprivOutFileSet)
      If aprivOutFileSet(k).sFileName = sTest Then lCount = lCount + 1
    Next k
    If lCount > 1 Then
      objOwner.addlog ("<error in='fncAllFileNamesAreUnique'>duplicate filename found in out array: " & aprivOutFileSet(i).sFileName & ": aborting</error>")
      Exit Function
    End If
    DoEvents
  Next i
    
  fncAllFileNamesAreUnique = True

ErrHandler:
  
End Function

Public Function fncItemExistsInOutArray(ByVal sItemName As String) As Boolean
Dim i As Long
  
  sItemName = LCase$(sItemName)
  fncItemExistsInOutArray = False
  For i = 0 To aOutFileSetMembers - 1
    If LCase$(aOutFileSet(i).sFileName) = sItemName Then
      fncItemExistsInOutArray = True
      Exit Function
    End If
  Next i

End Function

Public Function fncGetPreviousItemOfType( _
  ByVal lType As Long, _
  ByRef lPreviousArrayItem As Long, _
  ByVal lCurrentArrayItem As Long _
  ) As Boolean
Dim i As Long
Dim bFound As Boolean
  
  fncGetPreviousItemOfType = False
  
  bFound = False
  
  For i = (lCurrentArrayItem - 1) To 0 Step -1
    If aOutFileSet(i).eType = lType Then
      bFound = True
      Exit For
    End If
  Next i
 
  If bFound Then
    lPreviousArrayItem = i
  Else
    lPreviousArrayItem = -1
  End If
  
  fncGetPreviousItemOfType = True
  
ErrHandler:
  If Not fncGetPreviousItemOfType Then objOwner.addlog "<errH in='fncGetPreviousItemOfType'>fncGetPreviousItemOfType ErrH</errH>"
End Function

Private Function fncNccSmilSequenceIsCorrect( _
    oNccAnchorNodes As IXMLDOMNodeList _
    ) As Boolean
Dim oAnchorNode As IXMLDOMNode
Dim oHrefNode As IXMLDOMNode
Dim sCurrentFileName As String
Dim aSmilFiles() As String, lSmilFiles As Long
Dim i As Long, bBreakPoint As Boolean
  'this function checks that smil references in ncc
  'do not reoccur out of order
  
  On Error GoTo ErrHandler
  fncNccSmilSequenceIsCorrect = False
  
  lSmilFiles = 0
  For Each oAnchorNode In oNccAnchorNodes
    bBreakPoint = False
    Set oHrefNode = oAnchorNode.selectSingleNode("@href")
    sCurrentFileName = LCase$(fncStripId(oHrefNode.Text))
    
    'add the current smilfile to array
    ReDim Preserve aSmilFiles(lSmilFiles)
    aSmilFiles(lSmilFiles) = sCurrentFileName
    
    'go backwards in the array;
    'check that no other name comes inbetween same name
    For i = lSmilFiles To 0 Step -1
      If aSmilFiles(i) <> sCurrentFileName Then
        bBreakPoint = True
      End If
      If bBreakPoint And aSmilFiles(i) = sCurrentFileName Then
           objOwner.addlog "<error in='fncNccSmilSequenceIsCorrect'>illegal smilfile sequence in ncc: " _
             & sCurrentFileName & " first occured on pos " & i & "</error>"
           Exit Function
      End If
    Next i
    'up smilfilecount for next iteration
    lSmilFiles = lSmilFiles + 1
    DoEvents
  Next

'  For i = 0 To lSmilFiles - 1
'    objowner.addlog aSmilFiles(i)
'  Next i

  fncNccSmilSequenceIsCorrect = True
ErrHandler:
  If Not fncNccSmilSequenceIsCorrect Then objOwner.addlog "<errH in='fncNccSmilSequenceIsCorrect'>fncNccSmilSequenceIsCorrect ErrH</errH>"
End Function

Public Function fncGetArrayItemFromName(isFileName As String) As Long
Dim i As Long
'gets a filename as input, finds it in the output array, and returns this items position

  For i = 0 To aOutFileSetMembers - 1
   If LCase$(aOutFileSet(i).sFileName) = LCase$(isFileName) Then
     fncGetArrayItemFromName = i
     Exit Function
   End If
   DoEvents
  Next i
  
  fncGetArrayItemFromName = -1
  objOwner.addlog "<errH in='fncGetArrayItemFromName'>fncGetArrayItemFromName fail</errH>"
  
End Function

Private Function fncCreateContentDocMembers() As Boolean
  'this function createas an array that contains
  'the position of each contentdoc in aOutFileSet
  'used to save time when iterating
Dim i As Long
  On Error GoTo ErrHandler
  fncCreateContentDocMembers = False
  
  For i = 1 To aprivOutFileSetMembers - 1
    If aprivOutFileSet(i).eType = TYPE_SMIL_CONTENT Then
      ReDim Preserve aprivOutFileSetContentDocs(aprivOutFileSetContentDocMembers)
      aprivOutFileSetContentDocs(aprivOutFileSetContentDocMembers) = i
      aprivOutFileSetContentDocMembers = aprivOutFileSetContentDocMembers + 1
    End If
    DoEvents
  Next i
  'objowner.addlog "content docs in this book: " & CStr(aOutFileSetContentDocMembers)
  fncCreateContentDocMembers = True
  
'  For i = 0 To aOutFileSetContentDocMembers - 1
'    objowner.addlog aOutFileSet(aOutFileSetContentDocs(i)).sFileName
'  Next
  
ErrHandler:
  If Not fncCreateContentDocMembers Then objOwner.addlog "<errH in='fncCreateContentDocMembers'>fncCreateContentDocMembers ErrH</errH>"
End Function

Private Sub Class_Terminate()
  Dim lCounter As Long
  
  For lCounter = 0 To aprivOutFileSetMembers - 1
    Set aprivOutFileSet(lCounter) = Nothing
  Next lCounter

  For lCounter = 0 To aprivInFileSetMembers - 1
    Set aprivInFileSet(lCounter) = Nothing
  Next lCounter
  
  ReDim aprivOutFileSet(0)
  ReDim aprivInFileSet(0)
End Sub

Public Function fncTerminateObject() As Boolean
  Dim bolResult As Boolean, lCounter As Long
  
  bolResult = True
  
  For lCounter = 0 To aprivOutFileSetMembers - 1
    Set aprivOutFileSet(lCounter) = Nothing
    If Not aprivOutFileSet(lCounter) Is Nothing Then bolResult = False
  Next lCounter

  For lCounter = 0 To aprivInFileSetMembers - 1
    Set aprivInFileSet(lCounter) = Nothing
    If Not aprivInFileSet(lCounter) Is Nothing Then bolResult = False
  Next lCounter
  
  ReDim aprivOutFileSet(0)
  ReDim aprivInFileSet(0)
  aprivOutFileSetMembers = 0
  aprivInFileSetMembers = 0
  
  fncTerminateObject = bolResult
End Function

Public Function fncPrintFileSet(objOwner As oRegenerator) As Boolean
Dim i As Long
Dim sTemp As String
      On Error GoTo ErrHandler
      fncPrintFileSet = False
      objOwner.addlog "<fileset>"
      For i = 0 To UBound(aprivInFileSet) - 1
        sTemp = "<file item='" & CStr(i) & "' origName='" & fncGetFileName(aprivInFileSet(i).sAbsPath) & "' "
        If i <= UBound(aprivOutFileSet) Then
         sTemp = sTemp & "newName='" & aprivOutFileSet(i).sFileName
        Else
         sTemp = sTemp & "newName='not set"
        End If
         sTemp = sTemp & "'/>"
        objOwner.addlog sTemp
      Next i
      objOwner.addlog "</fileset>"
      fncPrintFileSet = True
      
ErrHandler:
      If Not fncPrintFileSet Then objOwner.addlog ("<errH in='fncPrintFileSet'>fncPrintFileSet ErrH</errH>")
End Function

