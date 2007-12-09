Attribute VB_Name = "mNamesAndURis"
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

Private Type eIdPairs
  sTextId As String
  sParId As String
End Type


  '***************************************************************
  '* renaming files, fixing invalid URIs, fixing case in fragments:
  '*
  '* bolSeqRename = user wants to rename smil and audio sequentially
  '* bolUseDcIdNum = use the numeric part of dc:identifer in hex as prefix of filename
  '* sWantedPrefix = dont use numerals but this string
  '* if bolUseDcIdNum = true then sWantedPrefix shall be ignored even if <> ""
  '* sequence:
  '* If bolSeqRename Then do sequrename
  '* Then check/fix invalid URIs regardless of bolSeqRename
  '* but if bolSeqRename = true only need to check content docs
  '* because sequrename refuses to do invalid prefixes
  '* Finally, check that sequfilerename an/or URIrename
  '* did not cause files having same name
  '***************************************************************

Public Function fncNamesAndUris( _
    bolSeqRename As Boolean, _
    bolUseDcIdNum As Boolean, _
    sWantedPrefix As String, _
    objOwner As oRegenerator _
        ) As Boolean

  On Error GoTo ErrHandler
  fncNamesAndUris = False
  objOwner.addlog "<status>initiating file and URI checking...</status>"
  
  If bolSeqRename Then
    If Not fncDoSequRenameNew( _
        bolUseDcIdNum, _
        sWantedPrefix, _
        objOwner _
    ) _
    Then GoTo ErrHandler
  End If 'bolSeqRename
  'Stop
  'Debug.Print objOwner.objFileSetHandler.aOutFileSet(10).sFileName
  'Debug.Print objOwner.objFileSetHandler.aOutFileSet(10).sDomData
  'Stop
  
  If Not fncFixInvalidURIs(bolSeqRename, objOwner) Then GoTo ErrHandler
    
  'check that no dupe filenames have occured
  If Not objOwner.objFileSetHandler.fncAllFileNamesAreUnique() Then GoTo ErrHandler

  fncNamesAndUris = True
ErrHandler:
  If Not fncNamesAndUris Then objOwner.addlog "<errH in='fncNamesAndUris'>fncNamesAndUris ErrH</errH>"
End Function

Private Function fncFixInvalidURIs( _
    bolSeqRename As Boolean, _
    objOwner As oRegenerator _
    ) As Boolean
Dim i As Long, sNewName As String, sOrigName As String
  On Error GoTo ErrHandler
  fncFixInvalidURIs = False
  
  'Note: excluded the ncc.html ref Replace$ in fncDoSequRename;
  'it is always done here since fncDoSeqRename is not always run
  If Not fncGetFileName(objOwner.objFileSetHandler.aInFileSet(0).sAbsPath) = "ncc.html" Then
    'objOwner.addlog "replacing references to ncc.html..."
    If Not fncReplaceThisFileNameRefInAllDocs _
      (fncGetFileName(objOwner.objFileSetHandler.aInFileSet(0).sAbsPath), "ncc.html", objOwner) Then GoTo ErrHandler
  End If
  
  'then proceed with invalid URI checking;
  'if bolSeqRename then this needs only to be done on contentdocs
  'since fncDoSequRename never can use output other than filenames giving valid URIs
  'for smilfiles and audiofiles

  For i = 1 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    '//REVISIT  should be done on type 'other' as well?
    'If (bolSeqRename And aOutFileSet(i).eType = smil_content) Or (Not bolSeqRename) Then
    If ((bolSeqRename) _
      And (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_CONTENT _
      Or objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_IMG)) _
      Or (Not bolSeqRename) Then
        sOrigName = objOwner.objFileSetHandler.aOutFileSet(i).sFileName
        If Not fncIsValidUriChars(sOrigName) Then
          sNewName = fncTruncToValidUriChars(sOrigName)
          If sNewName = "" Then
            'if sOrigName was only charwide then "" will be returned
            '//REVISIT add here a class that assigns four random ascii chars
            objOwner.addlog "<error in='fncFixInvalidURIs' arrayItem='" & CStr(i) & "'>new validUriChars name was zero length: aborting...</error>"
            GoTo ErrHandler
          End If
          objOwner.addlog "<message> filename " & sOrigName & " replaced with " & sNewName & "</message>"
          If Not fncReplaceThisFileNameRefInAllDocs(sOrigName, sNewName, objOwner) Then GoTo ErrHandler
          objOwner.objFileSetHandler.aOutFileSet(i).sFileName = sNewName
      End If
    End If
    DoEvents
  Next i
  
  fncFixInvalidURIs = True
ErrHandler:
  If Not fncFixInvalidURIs Then objOwner.addlog "<errH in='fncFixInvalidURIs' arrayItem='" & CStr(i) & "'>fncFixInvalidURIs ErrH</errH>"
End Function

Public Function fncReplaceThisFileNameRefInAllDocs( _
    sOrigName As String, _
    sNewName As String, _
    objOwner As oRegenerator _
    ) As Boolean
Dim i As Long, lEnd As Long
  On Error GoTo ErrHandler
  
  fncReplaceThisFileNameRefInAllDocs = False

  lEnd = objOwner.objFileSetHandler.aOutFileSetMembers - 1
  For i = 0 To lEnd
    If (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_NCC) Or _
      (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1) Or _
      (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_CONTENT) Then
      'The prefixing """ should assure that we do not accidentally
      'destroy a long filename while replacing a short one.
      objOwner.objFileSetHandler.aOutFileSet(i).sDomData = Replace$( _
                            objOwner.objFileSetHandler.aOutFileSet(i).sDomData, _
                            """" & sOrigName, _
                            """" & sNewName, 1, -1, vbTextCompare)
    End If
    DoEvents
  Next i

  fncReplaceThisFileNameRefInAllDocs = True

ErrHandler:
  If Not fncReplaceThisFileNameRefInAllDocs Then objOwner.addlog "<errH in='fncReplaceThisFileNameRefInAllDocs'>fncReplaceThisFileNameRefInAllDocs ErrH</errH>"
End Function

Private Function fncDoSequRenameNew( _
    ByRef bolUseDcIdNum As Boolean, _
    ByRef sWantedPrefix As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean

Dim i As Long, i2 As Long, i3 As Long, k As Long
Dim sExtension As String
Dim sPrefix As String
Dim sSafetyString As String

'fixed bug reported by brandon cnib
  On Error GoTo ErrHandler
  
  fncDoSequRenameNew = False
  
  objOwner.addlog "<status>sequential rename...</status>"
  
  'generate prefix
  If Not fncGeneratePrefix(sPrefix, bolUseDcIdNum, sWantedPrefix, objOwner) Then GoTo ErrHandler
  sPrefix = sPrefix & Chr(95)
  
  'create a string that prevents matches with pre-existing strings
  sSafetyString = Replace(Replace(Replace(Now, " ", ""), "-", ""), ":", "")
  
  'rename in outarray
  i2 = 0: i3 = 0
  
  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1) Then
      i2 = i2 + 1
      sExtension = fncGetExtension(objOwner.objFileSetHandler.aOutFileSet(i).sFileName)
      objOwner.objFileSetHandler.aOutFileSet(i).sFileName = sSafetyString & fncGenerateSequId(sPrefix, i2, "0000") & sExtension
    ElseIf (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_AUDIO) Then
      i3 = i3 + 1
      sExtension = fncGetExtension(objOwner.objFileSetHandler.aOutFileSet(i).sFileName)
      'fix for the forbidden "mpeg" extension for mp2 files - mg 20030215
      If sExtension = ".mpeg" Then sExtension = ".mp2"
      objOwner.objFileSetHandler.aOutFileSet(i).sFileName = sSafetyString & fncGenerateSequId(sPrefix, i3, "0000") & sExtension
    End If

'    DoEvents
  Next i
  
  'loop inside all text files and Replace$ all occurences of old filenames with new filenames
  'all while using utf-16. doing this with DOM takes too much time.
  Dim sOrigName As String
  Dim sNewName As String
    
 For k = 1 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
'''  20050417 - fix brandons rename bug: reverse loop instead:
'''   For k = objOwner.objFileSetHandler.aOutFileSetMembers - 1 To 1 Step -1
    '(exclude the ref to ncc.html Replace$ here;
    'it is always done in fixinvaliduris
    
    'if the file that has been renamed is a smilfile
    'other smilfiles need not be opened and searched
    'because a smilfile never points to a smilfile

    If (objOwner.objFileSetHandler.aOutFileSet(k).eType = TYPE_SMIL_1) Then
      sOrigName = fncGetFileName(objOwner.objFileSetHandler.aInFileSet(k).sAbsPath)
      sNewName = objOwner.objFileSetHandler.aOutFileSet(k).sFileName
      
      For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
        If (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_NCC) Or _
          (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_CONTENT) Then
          objOwner.objFileSetHandler.aOutFileSet(i).sDomData = Replace$( _
                            objOwner.objFileSetHandler.aOutFileSet(i).sDomData, _
                            """" & sOrigName, _
                            """" & sNewName, 1, -1, vbTextCompare)
          
        End If
      Next i
      'if the file that has been renamed is an audiofile
      'only smilfiles need not be opened and searched
      'because a ncc and content never points to audio
    ElseIf (objOwner.objFileSetHandler.aOutFileSet(k).eType = TYPE_SMIL_AUDIO) Then
      
      sOrigName = fncGetFileName(objOwner.objFileSetHandler.aInFileSet(k).sAbsPath)
      sNewName = objOwner.objFileSetHandler.aOutFileSet(k).sFileName
      'Stop
      For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
        If (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1) Then
          objOwner.objFileSetHandler.aOutFileSet(i).sDomData = Replace$( _
                            objOwner.objFileSetHandler.aOutFileSet(i).sDomData, _
                            """" & sOrigName, _
                            """" & sNewName, 1, -1, vbTextCompare)

        End If
      Next i
    End If
    DoEvents
  Next k
  
  'finally remove the safety string
  Dim q As Long
  For q = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    '... in the outfileset filename
        If (objOwner.objFileSetHandler.aOutFileSet(q).eType = TYPE_SMIL_AUDIO) Or _
          (objOwner.objFileSetHandler.aOutFileSet(q).eType = TYPE_SMIL_1) Then
           objOwner.objFileSetHandler.aOutFileSet(q).sFileName = Replace$(objOwner.objFileSetHandler.aOutFileSet(q).sFileName, sSafetyString, "")
        End If
    '... and in the outfileset domdata
        If (objOwner.objFileSetHandler.aOutFileSet(q).eType = TYPE_NCC) Or _
          (objOwner.objFileSetHandler.aOutFileSet(q).eType = TYPE_SMIL_CONTENT) Or _
          (objOwner.objFileSetHandler.aOutFileSet(q).eType = TYPE_SMIL_1) Then
            objOwner.objFileSetHandler.aOutFileSet(q).sDomData = Replace$(objOwner.objFileSetHandler.aOutFileSet(q).sDomData, sSafetyString, "")
        End If
  Next q

'debug

'Stop
'Debug.Print objOwner.objFileSetHandler.aOutFileSet(64).sFileName
'Debug.Print objOwner.objFileSetHandler.aOutFileSet(64).sDomData
'Stop

'Dim cp As Long
'For cp = 1 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
'  Debug.Print "old: " & fncGetFileName(objOwner.objFileSetHandler.aInFileSet(cp).sAbsPath) & _
'   " new: " & objOwner.objFileSetHandler.aOutFileSet(cp).sFileName
'Next cp
'Stop

  fncDoSequRenameNew = True
  
ErrHandler:
  If Not fncDoSequRenameNew Then objOwner.addlog "<errH in='fncDoSequRenameNew'>fncDoSequRenameNew ErrH</errH>"
End Function

'Private Function fncGetNonExistingName(ByVal inName As String, ByRef objOwner As oRegenerator) As String
''task: to make sure that inName does not already exist in inputArray
''if, so, append underscore to filename until not exists
''Dim retName As String
'
'  On Error GoTo errH
'
'  If Not fncNameExistsInInArray(inName, objOwner) Then
'    'this was unique already
'  Else
'    Do
'      inName = inName & "_"
'     Loop Until Not fncNameExistsInInArray(inName, objOwner)
'  End If
'
'  fncGetNonExistingName = inName
'  Exit Function
'
'errH:
'  objOwner.addlog "<errH in='fncGetNonExistingName'>fncGetNonExistingName ErrH</errH>"
'End Function

'Private Function fncNameExistsInInArray(sFileName As String, objOwner As oRegenerator) As Boolean
'Dim i As Long
'  On Error GoTo errH
'
'  For i = 1 To objOwner.objFileSetHandler.aInFileSetMembers - 1
'    If LCase$(sFileName) = LCase$(fncGetFileName(objOwner.objFileSetHandler.aInFileSet(i).sAbsPath)) Then
'      fncNameExistsInInArray = True
'      Exit Function
'    End If
'  Next i
'  fncNameExistsInInArray = False
'  Exit Function
'errH:
'  objOwner.addlog "<errH in='fncNameExistsInInArray'>fncNameExistsInInArray ErrH</errH>"
'End Function

'Private Function fncDoSequRename( _
'    ByRef bolUseDcIdNum As Boolean, _
'    ByRef sWantedPrefix As String, _
'    ByRef objOwner As oRegenerator _
'    ) As Boolean
'Stop 'doing this in fncDoSequRenameNew
'Dim i As Long, i2 As Long, i3 As Long, k As Long
'Dim sExtension As String
'Dim sPrefix As String
'
' On Error GoTo ErrHandler
'
'  fncDoSequRename = False
'
'  objOwner.addlog "<status>sequential rename...</status>"
'
'  'generate prefix
'  If Not fncGeneratePrefix(sPrefix, bolUseDcIdNum, sWantedPrefix, objOwner) Then GoTo ErrHandler
'  sPrefix = sPrefix & Chr(95)
'
'  'rename in outarray
'  i2 = 0: i3 = 0
'  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
'    If (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1) Then
'      i2 = i2 + 1
'      sExtension = fncGetExtension(objOwner.objFileSetHandler.aOutFileSet(i).sFileName)
'      objOwner.objFileSetHandler.aOutFileSet(i).sFileName = fncGenerateSequId(sPrefix, i2, "0000") & sExtension
'    ElseIf (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_AUDIO) Then
'      i3 = i3 + 1
'      sExtension = fncGetExtension(objOwner.objFileSetHandler.aOutFileSet(i).sFileName)
'      'fix for the forbidden "mpeg" extension for mp2 files - mg 20030215
'      If sExtension = ".mpeg" Then sExtension = ".mp2"
'      objOwner.objFileSetHandler.aOutFileSet(i).sFileName = fncGenerateSequId(sPrefix, i3, "0000") & sExtension
'    End If
'    DoEvents
'  Next i
'
'  'loop inside all text files and Replace$ all occurences of old filenames with new filenames
'  'all while using utf-16. doing this with DOM takes too much time.
'  Dim sOrigName As String
'  Dim sNewName As String
'
'  For k = 1 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
'    '(exclude the ref to ncc.html Replace$ here;
'    'it is always done in fixinvaliduris
'
'    'if the file that has been renamed is a smilfile
'    'other smilfiles need not be opened and searched
'    'because a smilfile never points to a smilfile
'
'    If (objOwner.objFileSetHandler.aOutFileSet(k).eType = TYPE_SMIL_1) Then
'      sOrigName = fncGetFileName(objOwner.objFileSetHandler.aInFileSet(k).sAbsPath)
'      sNewName = objOwner.objFileSetHandler.aOutFileSet(k).sFileName
'
'      For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
'        If (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_NCC) Or _
'          (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_CONTENT) Then
'          objOwner.objFileSetHandler.aOutFileSet(i).sDomData = Replace$( _
'                            objOwner.objFileSetHandler.aOutFileSet(i).sDomData, _
'                            """" & sOrigName, _
'                            """" & sNewName, 1, -1, vbTextCompare)
'
'        End If
'      Next i
'      'if the file that has been renamed is an audiofile
'      'only smilfiles need not be opened and searched
'      'because a ncc and content never points to audio
'    ElseIf (objOwner.objFileSetHandler.aOutFileSet(k).eType = TYPE_SMIL_AUDIO) Then
'
'      sOrigName = fncGetFileName(objOwner.objFileSetHandler.aInFileSet(k).sAbsPath)
'      sNewName = objOwner.objFileSetHandler.aOutFileSet(k).sFileName
'      'Stop
'      For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
'        If (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1) Then
'          objOwner.objFileSetHandler.aOutFileSet(i).sDomData = Replace$( _
'                            objOwner.objFileSetHandler.aOutFileSet(i).sDomData, _
'                            """" & sOrigName, _
'                            """" & sNewName, 1, -1, vbTextCompare)
'
'        End If
'      Next i
''      Stop
'
'
'    End If
'    DoEvents
'  Next k
'
''  Stop
''  Debug.Print objOwner.objFileSetHandler.aOutFileSet(10).sFileName
''  Debug.Print objOwner.objFileSetHandler.aOutFileSet(10).sDomData
''  Stop
'
'  fncDoSequRename = True
'
'ErrHandler:
'  If Not fncDoSequRename Then objOwner.addlog "<errH in='fncDoSequrename'>fncDoSequrename ErrH</errH>"
'End Function

Private Function fncGeneratePrefix( _
    ByRef sPrefix As String, _
    ByVal bolUseDcIdNum, _
    ByVal sWantedPrefix As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean

Dim oNode As IXMLDOMNode
Dim oAttrNode As IXMLDOMNode
Dim sDcIdentifier As String
  
  On Error GoTo ErrHandler
  fncGeneratePrefix = False
  
  If bolUseDcIdNum Then
    If Not objOwner.objCommonMeta.DcIdentifier Is Nothing Then Set oNode = objOwner.objCommonMeta.DcIdentifier.cloneNode(True)
    If Not oNode Is Nothing Then
      Set oAttrNode = oNode.selectSingleNode("@content")
      sDcIdentifier = oAttrNode.Text
    End If
    
    If sDcIdentifier <> "" Then
      If IsNumeric(sDcIdentifier) Then
        sPrefix = LCase$(Hex(CLng(sDcIdentifier)))
      Else
        If Not fncMakeNumeric(sDcIdentifier, objOwner) Then
          objOwner.addlog "<message>numerization of id failed; using prefix 'dtb' instead.</message>"
          sPrefix = "dtb"
        Else
          'hex returns max eight hexadecimal characters
          If CDbl(sDcIdentifier) > 999999999 Then
           sDcIdentifier = Mid(sDcIdentifier, 1, 9)
           objOwner.addlog "<message>truncated numeric portion of identifier to " & sDcIdentifier & " for use in hex filenames</message>"
          End If
          
          If sDcIdentifier <> "" Then
            sPrefix = LCase$(Hex(CLng(sDcIdentifier)))
          Else
            objOwner.addlog "<message>zero numbers to use in dc:identifier; using prefix 'dtb' instead.</message>"
            sPrefix = "dtb"
          End If
        End If
      End If 'IsNumeric(sDcIdentifier)
    Else
      objOwner.addlog "<message>found dc:identifer not ok for use in rename. Using 'dtb' instead</message>"
      sPrefix = "dtb"
    End If '(sDcIdentifer <> "" )
  Else 'If not bolUseDcIdNum
    If sWantedPrefix <> "" Then
     If Not fncIsValidUriChars(sWantedPrefix) Then
       sPrefix = fncTruncToValidUriChars(sWantedPrefix)
       If sPrefix = "" Then
         objOwner.addlog "<warning>only invalid URI chars entered as file rename prefix - using 'dtb' instead</warning>"
         sPrefix = "dtb"
       End If 'sPrefix = ""
     Else
      sPrefix = sWantedPrefix
     End If
    Else
      objOwner.addlog "<message>empty prefix found - using 'dtb' instead</message>"
      sPrefix = "dtb"
    End If 'sWantedPrefix <> ""
  End If 'If  bolUseDcIdNum

  fncGeneratePrefix = True
ErrHandler:
  If Not fncGeneratePrefix Then objOwner.addlog "<errH in='fncGeneratePrefix'>fncGeneratePrefix ErrH</errH>"
End Function

Private Function fncMakeNumeric(ByRef sPrefix As String, objOwner As oRegenerator) As Boolean
Dim i As Long, sCh As String
  On Error GoTo ErrHandler
  fncMakeNumeric = False
  
  Do
    For i = 0 To Len(sPrefix)
      sCh = Mid(sPrefix, i + 1, 1)
      If Not IsNumeric(sCh) Then
        sPrefix = Replace$(sPrefix, sCh, "")
      End If
    Next
  Loop Until IsNumeric(sPrefix) Or Len(sPrefix) = 0

  fncMakeNumeric = True
  
ErrHandler:
  If Not fncMakeNumeric Then objOwner.addlog "<errH in='fncMakeNumeric'>fncMakeNumeric ErrH</errH>"
End Function

Private Function fncLcaseThisUriFragmentCase( _
    ByRef sUri As String _
    ) As String
Dim sID As String
  
  sID = fncGetId(sUri)
  sUri = Replace$(sUri, sID, LCase$(sID), , , vbTextCompare)
  fncLcaseThisUriFragmentCase = sUri

End Function

Public Function fncFixFragmentCase( _
    oMemberDom As MSXML2.DOMDocument40, _
    lFileType As Long, _
    objOwner As oRegenerator _
    ) As Boolean
Dim oNodes As IXMLDOMNodeList
Dim oNode As IXMLDOMNode
Dim sXpath As String
 
 On Error GoTo ErrHandler
 
 fncFixFragmentCase = False
  
  
  Select Case lFileType
    Case TYPE_NCC, TYPE_SMIL_CONTENT
      'create a nodelist of all URI and id carriers in xhtml
      sXpath = "//@href | //@id"
    Case TYPE_SMIL_1
      'create a nodelist of all URI and id carriers in smil
      'only needs to be done on text src and text id
      'since audio id´s are redone by default
      
      'mg 20050310 added pars with existing ids as well
      'sXpath = "//text/@src | //text/@id"
      sXpath = "//text/@src | //text/@id | //par/@id"
    Case Else
  End Select
  
  Set oNodes = oMemberDom.selectNodes(sXpath)
  
  If Not oNodes Is Nothing Then
    For Each oNode In oNodes
      Select Case oNode.nodeName
        Case "id"
          oNode.Text = LCase$(oNode.Text)
        Case Else
          oNode.Text = fncLcaseThisUriFragmentCase(oNode.Text)
      End Select
      
    Next
  Else
    objOwner.addlog "<message in='fncFixFragmentCase'>fncFixFragmentCase inparam contained zero URI nodes</message>"
  End If 'Not oNodes Is Nothing
  
  fncFixFragmentCase = True
  
ErrHandler:
  If Not fncFixFragmentCase Then objOwner.addlog "<errH in='fncFixFragmentCase'>fncFixFragmentCase ErrH</errH>"
End Function

Public Function fncDisableBrokenXhtmlLinks( _
    fncEstimateBrokenXhtmlLinks As Boolean, _
    objOwner As oRegenerator) _
    As Boolean
Dim oXhtmlDom As New MSXML2.DOMDocument40
    oXhtmlDom.async = False
    oXhtmlDom.validateOnParse = False
    oXhtmlDom.resolveExternals = False
    oXhtmlDom.preserveWhiteSpace = False
    oXhtmlDom.setProperty "SelectionLanguage", "XPath"
    oXhtmlDom.setProperty "NewParser", True
Dim oSmilDom As New MSXML2.DOMDocument40
    oSmilDom.async = False
    oSmilDom.validateOnParse = False
    oSmilDom.resolveExternals = False
    oSmilDom.preserveWhiteSpace = False
    oSmilDom.setProperty "SelectionLanguage", "XPath"
    oSmilDom.setProperty "NewParser", True

Dim i As Long
Dim oAnchorNodes As IXMLDOMNodeList
Dim oAnchorNode As IXMLDOMNode, oAnchorNodeItem As Long
Dim sUriBase As String, sUriFragment As String, sUri As String
Dim lDisabledNodes As Long, lEstimatedNodes As Long
Dim sLastOpenedDocument As String
Dim lArrayItem As Long
Dim bolEstimationSuccess As Boolean

'goes through the smilrefs in ncc and contentdoc
'disables, or if estimate is on, tries to estimate their orig pos
'note: the target (in the smilfile) is not modded by this func,
'only the xhtml href link
  
  On Error GoTo ErrHandler
  fncDisableBrokenXhtmlLinks = False
  objOwner.addlog ("<status>checking for broken links...</status>")
  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_NCC) Or (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_CONTENT) Then
      lDisabledNodes = 0
      lEstimatedNodes = 0
      If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(i).sDomData, oXhtmlDom, objOwner) Then GoTo ErrHandler
      'get all href nodes in nccor content doc
      Set oAnchorNodes = oXhtmlDom.selectNodes("//a[contains(@href,'.smil#')]")
      If Not oAnchorNodes Is Nothing Then
        oAnchorNodeItem = -1
        For Each oAnchorNode In oAnchorNodes
          oAnchorNodeItem = oAnchorNodeItem + 1
          sUri = oAnchorNode.selectSingleNode("@href").Text
          sUriBase = fncStripId(sUri)
          sUriFragment = fncGetId(sUri)
                             
          'only reload targetdoc if not the same as previous iterat
          If LCase$(sLastOpenedDocument) <> LCase$(sUriBase) Then
            lArrayItem = objOwner.objFileSetHandler.fncGetArrayItemFromName(sUriBase)
            If lArrayItem > 0 Then
              If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(lArrayItem).sDomData, oSmilDom, objOwner) Then GoTo ErrHandler
            Else
              'a nonexisting smil file was referenced
              objOwner.addlog "<warning in='fncDisableBrokenXhtmlLinks'>" & sUriBase & "not found in array</warning>"
              GoTo SkipObject
            End If
          End If 'sLastOpenedDocument <> sUriBase
            
          'now oSmilDom contains sUriBase file
          If Not fncNodeExists(oSmilDom, "//*[@id='" & sUriFragment & "']", "", objOwner) Then
            'the smil fragment was not found in dom
            bolEstimationSuccess = False
            If fncEstimateBrokenXhtmlLinks Then
              'try to estimate a position
              If Not fncEstimateNode(oAnchorNode, oAnchorNodeItem, oAnchorNodes, _
                 sUriBase, sUriFragment, lArrayItem, oSmilDom, _
                 bolEstimationSuccess, objOwner) Then GoTo ErrHandler
              If bolEstimationSuccess Then lEstimatedNodes = lEstimatedNodes + 1
            End If
            If Not bolEstimationSuccess Then
              'mg20040216; if node is a heading, dont disable
              Dim sNodeName As String
              sNodeName = oAnchorNode.parentNode.nodeName
              If (sNodeName <> "h1") And (sNodeName <> "h2") And (sNodeName <> "h3") And (sNodeName <> "h4") And (sNodeName <> "h5") And (sNodeName <> "h6") Then
                If Not fncDisableNode(oAnchorNode.parentNode, objOwner) Then GoTo ErrHandler
                lDisabledNodes = lDisabledNodes + 1
              End If
            End If
          End If 'Not fncNodeExists
          sLastOpenedDocument = sUriBase
SkipObject:
        Next 'For Each oAnchorNode In oAnchorNodes
      End If 'Not oAnchorNodes Is Nothing Then
      objOwner.objFileSetHandler.aOutFileSet(i).sDomData = oXhtmlDom.xml
      If lEstimatedNodes > 0 Then objOwner.addlog "<warning class='estimatedNodes'>warning:" & CStr(lEstimatedNodes) & " nodes estimated in " & objOwner.objFileSetHandler.aOutFileSet(i).sFileName & "</warning>"
      If lDisabledNodes > 0 Then objOwner.addlog "<warning class='disabledNodes'>warning:" & CStr(lDisabledNodes) & " nodes disabled in " & objOwner.objFileSetHandler.aOutFileSet(i).sFileName & "</warning>"
    End If 'is ncc or content
  Next i
  fncDisableBrokenXhtmlLinks = True

ErrHandler:
  Set oAnchorNode = Nothing
  Set oAnchorNodes = Nothing
  Set oXhtmlDom = Nothing
  Set oSmilDom = Nothing
  If Not fncDisableBrokenXhtmlLinks Then objOwner.addlog "<errH in='fncDisableBrokenXhtmlLinks'>fncDisableBrokenXhtmlLinks ErrH</errH>"
End Function


Private Function fncEstimateNode( _
 ByRef oAnchorNode As IXMLDOMNode, _
 ByRef oAnchorNodeItem As Long, _
 ByRef oAnchorNodes As IXMLDOMNodeList, _
 ByRef sUriBase As String, _
 ByRef sUriFragment As String, _
 ByRef lArrayItem As Long, _
 ByRef oSmilDom As MSXML2.DOMDocument40, _
 ByRef bolEstimationSuccess As Boolean, _
 ByRef objOwner As oRegenerator _
 ) As Boolean
 
 'this func is called
 'when a ncc or content doc href points to nonexisting SMIL target
 'it will try to reinsert a link
 'at a position somewhere inbetween pre and post links
 'oAnchorNode is the link in ncc or content that is broken, its byref and will be filled
 'in here if success
 'oAnchorNodeItem is the current nodelist item number
 'oAnchorNodes is the seq list of anchornodes with "#.smil" href content
 'sUriBase is filename of smilfile pointed to
 'sUriFragment is (nonexisting) id of target pointed to
 'bolestimationsuccess shall return whether an estimation mod was actually done
 
Dim sPrevUri As String, sPrevUriBase As String, sPrevUriFragment As String
Dim sNextUri As String, sNextUriBase As String, sNextUriFragment As String
Dim sNewUri As String
Dim lPrevArrayItem As Long, lNextArrayItem As Long
Dim sPrevTargetNodeName As String
Dim sNextTargetNodeName As String

Dim oPrevDom As New MSXML2.DOMDocument40
    oPrevDom.async = False
    oPrevDom.validateOnParse = False
    oPrevDom.resolveExternals = False
    oPrevDom.preserveWhiteSpace = False
    oPrevDom.setProperty "SelectionLanguage", "XPath"
    oPrevDom.setProperty "NewParser", True
 
Dim oNextDom As New MSXML2.DOMDocument40
    oNextDom.async = False
    oNextDom.validateOnParse = False
    oNextDom.resolveExternals = False
    oNextDom.preserveWhiteSpace = False
    oNextDom.setProperty "SelectionLanguage", "XPath"
    oNextDom.setProperty "NewParser", True
 
 On Error GoTo ErrHandler
 fncEstimateNode = False
    
 'find previous link target pos
 If Not oAnchorNodeItem = 0 Then
   'if its not the very first anchor in the document
   sPrevUri = oAnchorNodes.Item(oAnchorNodeItem - 1).selectSingleNode("@href").Text
   sPrevUriBase = LCase$(fncStripId(sPrevUri))
   sPrevUriFragment = fncGetId(sPrevUri)
   
   'check if the previous link is okay; if not, null the strings
   If sPrevUriBase = sUriBase Then
      'prev link is in same file as input broken link
      If Not fncNodeExists(oSmilDom, "//*[@id='" & sPrevUriFragment & "']", sPrevTargetNodeName, objOwner) Then
         sPrevUri = ""
         sPrevTargetNodeName = ""
      End If
   Else
     'prev link is in other file than input broken link
     lPrevArrayItem = objOwner.objFileSetHandler.fncGetArrayItemFromName(sPrevUriBase)
     If lPrevArrayItem > 0 Then
       If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(lPrevArrayItem).sDomData, oPrevDom, objOwner) Then GoTo ErrHandler
     Else
       'a nonexisting smil file was referenced
       sPrevUri = ""
       objOwner.addlog "<error in='fncEstimateNode'>" & sPrevUriFragment & "not found in array</error>"
     End If
     If Not fncNodeExists(oPrevDom, "//*[@id='" & sPrevUriFragment & "']", sPrevTargetNodeName, objOwner) Then
       sPrevUri = ""
       sPrevTargetNodeName = ""
     End If
   End If
 End If
 
 'find next link target pos
 If Not oAnchorNodeItem = (oAnchorNodes.length - 1) Then
   'if its not the very last anchor in the document
   sNextUri = oAnchorNodes.Item(oAnchorNodeItem + 1).selectSingleNode("@href").Text
   sNextUriBase = LCase$(fncStripId(sNextUri))
   sNextUriFragment = fncGetId(sNextUri)
  
   'check if the previous link is okay; if not, null the strings
   If sNextUriBase = sUriBase Then
      'Next link is in same file as input broken link
      If Not fncNodeExists(oSmilDom, "//*[@id='" & sNextUriFragment & "']", sNextTargetNodeName, objOwner) Then
         sNextUri = ""
         sNextTargetNodeName = ""
      End If
   Else
     'Next link is in other file than input broken link
     lNextArrayItem = objOwner.objFileSetHandler.fncGetArrayItemFromName(sNextUriBase)
     If lNextArrayItem > 0 Then
       If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(lNextArrayItem).sDomData, oNextDom, objOwner) Then GoTo ErrHandler
     Else
       'a nonexisting smil file was referenced
       sNextUri = ""
       objOwner.addlog "<error in='fncEstimateNode'>" & sNextUriBase & "not found in array</error>"
     End If
     If Not fncNodeExists(oNextDom, "//*[@id='" & sNextUriFragment & "']", sNextTargetNodeName, objOwner) Then
       sNextUri = ""
       sNextTargetNodeName = ""
     End If
   End If
 End If
 
' Debug.Print sPrevUri
' Debug.Print sPrevTargetNodeName
' Debug.Print sNextUri
' Debug.Print sNextTargetNodeName
' Stop
'
' now there are two strings for prev and next links;
' if either is "" means that target not existed
' but at least prev has already been visited by this function if it didnt resolve so...
'
' if both exists
'   if all three links are in same smilfile: put on par inbetween,
'   if no par inbetween then put on same as prev
'   if prev is in same as input link: put on following par
'   if next is in same as input link: put on preceeding par
'   if prev and next are in different smilfiles than input link:
'     put input at top of current
' if only prev exists
'   if in same smilfile: place on par following prev
'   if in different smilfile (and this diff file is previous in array): place on top of input links own smilfile
' if only next exists
'   if in same smilfile: place on par preceding next
'   if in different smilfile (and this diff file is following in array): place on top of input links own smilfile
' if none existed
'   abort, or put link on top par of smilfile?


  'if both exists
  If sPrevUri <> "" And sNextUri <> "" Then
  '  if all three links are in same smilfile
     If (sUriBase = sPrevUriBase) And (sUriBase = sNextUriBase) Then
  '     put on par following preceeding,
        If fncGetAdjacentParId(sUriBase, sPrevUriFragment, sPrevTargetNodeName, sNextTargetNodeName, POSITION_NEXT, sNewUri, oSmilDom, objOwner) And sNewUri <> "" Then oAnchorNode.selectSingleNode("@href").Text = sNewUri
  '  elseif prev is in same file as input link
     ElseIf (sUriBase = sPrevUriBase) Then
  '     put on par following the previous (resolving) par
        If fncGetAdjacentParId(sUriBase, sPrevUriFragment, sPrevTargetNodeName, sNextTargetNodeName, POSITION_NEXT, sNewUri, oSmilDom, objOwner) And sNewUri <> "" Then oAnchorNode.selectSingleNode("@href").Text = sNewUri
  '  elseif next is in same file as input link
     ElseIf (sUriBase = sNextUriBase) Then
  '     put on preceeding par
        If fncGetAdjacentParId(sUriBase, sNextUriFragment, sPrevTargetNodeName, sNextTargetNodeName, POSITION_PREC, sNewUri, oSmilDom, objOwner) And sNewUri <> "" Then oAnchorNode.selectSingleNode("@href").Text = sNewUri
  '  elseif prev and next are in different smilfiles than input link
     Else 'If (sUriBase <> sPrevUriBase) And (sUriBase <> sNextUriBase) Then
  '     put input at top of current
        If fncGetAdjacentParId(sUriBase, "not needed", sPrevTargetNodeName, sNextTargetNodeName, POSITION_TOP, sNewUri, oSmilDom, objOwner) And sNewUri <> "" Then oAnchorNode.selectSingleNode("@href").Text = sNewUri
    End If '(sUriBase = sPrevUriBase) And (sUriBase = sNextUriBase)


  'if only prev exists
  ElseIf sPrevUri <> "" And sNextUri = "" Then
  '  if in same smilfile
     If (sPrevUriBase = sUriBase) Then
  '    place on par following prev
       If fncGetAdjacentParId(sUriBase, sPrevUriFragment, sPrevTargetNodeName, sNextTargetNodeName, POSITION_NEXT, sNewUri, oSmilDom, objOwner) And sNewUri <> "" Then oAnchorNode.selectSingleNode("@href").Text = sNewUri
  '  if in different smilfile (and this diff file is previous in array)
     Else
  '    place on top of input links own smilfile
       If fncGetAdjacentParId(sUriBase, "not needed", sPrevTargetNodeName, sNextTargetNodeName, POSITION_TOP, sNewUri, oSmilDom, objOwner) And sNewUri <> "" Then oAnchorNode.selectSingleNode("@href").Text = sNewUri
     End If '(sPrevUriBase = sUriBase)
  'if only next exists
  ElseIf sPrevUri = "" And sNextUri <> "" Then
  '  if in same smilfile
     If (sNextUriBase = sUriBase) Then
  '    place on par preceding next
       If fncGetAdjacentParId(sUriBase, sNextUriFragment, sPrevTargetNodeName, sNextTargetNodeName, POSITION_PREC, sNewUri, oSmilDom, objOwner) And sNewUri <> "" Then oAnchorNode.selectSingleNode("@href").Text = sNewUri
  '  if in different smilfile (and this diff file is following in array)
     Else
  '    place on top of input links own smilfile
       If fncGetAdjacentParId(sUriBase, "not needed", sPrevTargetNodeName, sNextTargetNodeName, POSITION_TOP, sNewUri, oSmilDom, objOwner) And sNewUri <> "" Then oAnchorNode.selectSingleNode("@href").Text = sNewUri
     End If '(sNextUriBase = sUriBase)
  'if none existed
  ElseIf sPrevUri = "" And sNextUri = "" Then
  '   abort, or put link on top par of smilfile?
      If fncGetAdjacentParId(sUriBase, "not needed", sPrevTargetNodeName, sNextTargetNodeName, POSITION_TOP, sNewUri, oSmilDom, objOwner) And sNewUri <> "" Then oAnchorNode.selectSingleNode("@href").Text = sNewUri
  Else
    ' this shouldnt happen
    Debug.Print "unexpected else case in fncEstimateNode"
  End If 'sPrevUri, sNextUri tests

  If sNewUri <> "" Then
    'Dim sBrokenUri As String
    objOwner.addlog "<warning in='fncEstimateNode'>warning: broken URI " & sUriBase & "#" & sUriFragment & " estimated to: " & sNewUri & "</warning>"
    bolEstimationSuccess = True
  End If
 
  fncEstimateNode = True

ErrHandler:
  Set oPrevDom = Nothing
  Set oNextDom = Nothing
  If Not fncEstimateNode Then objOwner.addlog "<errH in='fncEstimateNode'>fncEstimateNode ErrH</errH>"
End Function

Private Function fncGetAdjacentParId( _
  ByVal sUriBase As String, _
  ByVal sResolvingFragment As String, _
  ByVal sPrevTargetNodeName As String, _
  ByVal sNextTargetNodeName As String, _
  ByVal lPosition As Long, _
  ByRef sNewUri As String, _
  ByRef oSmilDom As MSXML2.DOMDocument40, _
  ByRef objOwner As oRegenerator _
  ) As Boolean
Dim oContextPar As IXMLDOMNode
Dim oIteratPar As IXMLDOMNode
Dim oParNodes As IXMLDOMNodeList
Dim lParNodesItem As Long
'Dim lParItem As Long
  On Error GoTo ErrHandler
  fncGetAdjacentParId = False
  'sUriBase is name of smilfile where context par resides
  'sResolvingFragment is value of par/text id attribute where context par resides
  'POSITION const tells which adjacent par to return
  '  -if POSITION_NEXT, return the par following sUriFragment
  'sNewUri is the value to return byref
  
  
  Set oParNodes = oSmilDom.selectNodes("//par")
  If (Not oParNodes Is Nothing) Then
    If lPosition = POSITION_TOP Then
      'return a uri that points to top par
      Set oContextPar = oSmilDom.selectSingleNode("//par[1]")
      If Not oContextPar Is Nothing Then
        If Not fncGetUriFromContextPar(oContextPar, sNewUri, sUriBase, sPrevTargetNodeName, sNextTargetNodeName, objOwner) Then GoTo ErrHandler
      Else
        sNewUri = ""
      End If 'Not oContextPar Is Nothing
      fncGetAdjacentParId = True
      GoTo ErrHandler
    End If 'lPosition = POSITION_TOP
  
    'set the context node
    Set oContextPar = oSmilDom.selectSingleNode("//*[@id='" & sResolvingFragment & "']")
    
    'make sure the context par really is a par (may be a text child)
    If oContextPar.nodeName = "text" Then Set oContextPar = oContextPar.parentNode
    If oContextPar.nodeName <> "par" Then
      'we are somewhere undefined
      Exit Function
    End If
        
    If (Not oContextPar Is Nothing) Then
      lParNodesItem = -1
      For Each oIteratPar In oParNodes
        lParNodesItem = lParNodesItem + 1
        If oIteratPar.xml = oContextPar.xml Then
          'we are at the context node in the par nodelist
          If lPosition = POSITION_NEXT Then
'            Set oContextPar = oParNodes.Item(lParNodesItem + 1)
             'mg20040217; if not context par is last par
             If oParNodes.Item(lParNodesItem + 1) Is Nothing Then
               Set oContextPar = oParNodes.Item(lParNodesItem)
             Else
               Set oContextPar = oParNodes.Item(lParNodesItem + 1)
             End If

          ElseIf lPosition = POSITION_PREC Then
'            Set oContextPar = oParNodes.Item(lParNodesItem - 1)
             'mg20040217; if not context par is first par
             If oParNodes.Item(lParNodesItem - 1) Is Nothing Then
               Set oContextPar = oParNodes.Item(lParNodesItem)
             Else
               Set oContextPar = oParNodes.Item(lParNodesItem - 1)
             End If
             
          End If
          If Not oContextPar Is Nothing Then
            'send sNewUri away and fill it with a base#id
            If Not fncGetUriFromContextPar(oContextPar, sNewUri, sUriBase, sPrevTargetNodeName, sNextTargetNodeName, objOwner) Then GoTo ErrHandler
          Else
            sNewUri = ""
          End If 'Not oContextPar Is Nothing
          'no need to continue looping so exit
          fncGetAdjacentParId = True
          Exit Function
        End If 'oIteratPar.xml = oContextPar.xml
      Next
    Else
      sNewUri = ""
    End If 'Not oContextPar Is Nothing
  Else
    sNewUri = ""
  End If 'Not oParNodes Is Nothing
  fncGetAdjacentParId = True
ErrHandler:
  Set oContextPar = Nothing
  Set oIteratPar = Nothing
  Set oParNodes = Nothing

  If Not fncGetAdjacentParId Then objOwner.addlog "<errH in='fncGetAdjacentParId'>fncGetAdjacentParId ErrH</errH>"
End Function

Private Function fncGetUriFromContextPar( _
 ByRef oContextPar As IXMLDOMNode, _
 ByRef sNewUri As String, _
 ByVal sUriBase As String, _
 ByVal sPrevTargetNodeName As String, _
 ByVal sNextTargetNodeName As String, _
 ByRef objOwner As oRegenerator _
 ) As Boolean
 Dim sNewFragment As String
 Dim oTargetNode As IXMLDOMNode

  On Error GoTo ErrHandler
  fncGetUriFromContextPar = False
 
  'the func should retain par or text pointers;
  If (sPrevTargetNodeName = "par") And (sNextTargetNodeName = "par") Then
    Set oTargetNode = oContextPar.selectSingleNode("@id")
    If Not oTargetNode Is Nothing Then
      sNewFragment = oTargetNode.Text
    End If
  Else
    Set oTargetNode = oContextPar.selectSingleNode("text/@id")
    If Not oTargetNode Is Nothing Then
      sNewFragment = oTargetNode.Text
    End If
  End If
  
  If sNewFragment <> "" Then
    sNewUri = sUriBase & "#" & sNewFragment
  End If
  
  fncGetUriFromContextPar = True
ErrHandler:
  Set oTargetNode = Nothing
  If Not fncGetUriFromContextPar Then objOwner.addlog "<errH in='fncGetUriFromContextPar'>fncGetUriFromContextPar ErrH</errH>"
End Function

Private Function fncDisableNode(oNode As IXMLDOMNode, objOwner As oRegenerator) As Boolean
Dim oSpanWrapper As IXMLDOMNode
Dim oComment As IXMLDOMComment

  On Error GoTo ErrHandler
  fncDisableNode = False
  'create span wrapper
  Set oSpanWrapper = oNode.ownerDocument.createNode(NODE_ELEMENT, "span", "")
  If Not fncAppendAttribute(oSpanWrapper, "class", "disabled", objOwner) Then GoTo ErrHandler
  'add span element
  Set oSpanWrapper = oNode.parentNode.insertBefore(oSpanWrapper, oNode)
  'move oNode into span
  Set oNode = oSpanWrapper.appendChild(oNode)
  'create comment
  Set oComment = oNode.ownerDocument.createComment("")
  'add comment
  Set oComment = oSpanWrapper.parentNode.insertBefore(oComment, oSpanWrapper)
  'put the node into it
  oComment.appendData (" " & oSpanWrapper.xml & " ")
  'delete the node
  Set oSpanWrapper = oSpanWrapper.parentNode.removeChild(oSpanWrapper)
  
  fncDisableNode = True
ErrHandler:
  If Not fncDisableNode Then objOwner.addlog "<errH in='fncDisableNode'>fncDisableNode ErrH</errH>"
End Function

Public Function fncAdjustSmilTargetPointers(lSmilTarget, objOwner As oRegenerator) As Boolean
Dim oSmilDom As New MSXML2.DOMDocument40
    oSmilDom.async = False
    oSmilDom.validateOnParse = False
    oSmilDom.resolveExternals = False
    oSmilDom.preserveWhiteSpace = False
    oSmilDom.setProperty "SelectionLanguage", "XPath"
    oSmilDom.setProperty "NewParser", True

Dim aIdPairs() As eIdPairs
Dim lIdCount As Long
Dim i As Long, k As Long
Dim oIdNodes As IXMLDOMNodeList
Dim oIdNode As IXMLDOMNode
' get each par/text/@id
' step back up to the par parent, get its id
' replace text id value with par id value in ncc and content docs
' fnc assumes that all text and pars have ids; which should be true at this point in the regen process
  
  On Error GoTo ErrHandler
  fncAdjustSmilTargetPointers = False
  Dim sSmilTarget As String
  If lSmilTarget = SMILTARGET_PAR Then sSmilTarget = "par" Else sSmilTarget = "text"
  objOwner.addlog "<status>redirecting text targets to " & sSmilTarget & " </status>"
  'collect all text and par ids in an array
  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1 Then
      If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(i).sDomData, oSmilDom, objOwner) Then GoTo ErrHandler
      Set oIdNodes = oSmilDom.selectNodes("//par/text")
      For Each oIdNode In oIdNodes
        ReDim Preserve aIdPairs(lIdCount)
        aIdPairs(lIdCount).sTextId = oIdNode.selectSingleNode("@id").Text
        Set oIdNode = oIdNode.parentNode
        aIdPairs(lIdCount).sParId = oIdNode.selectSingleNode("@id").Text
        lIdCount = lIdCount + 1
      Next
    End If 'objOwner.objFileSetHandler.aOutFileSet(i).eType
  Next i
    
  'go through the array and run replace in ncc and content docs
  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_NCC) Or _
     (objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_CONTENT) Then
     If lSmilTarget = SMILTARGET_PAR Then
       For k = 0 To lIdCount - 1
         objOwner.objFileSetHandler.aOutFileSet(i).sDomData = Replace$( _
                            objOwner.objFileSetHandler.aOutFileSet(i).sDomData, _
                            "#" & aIdPairs(k).sTextId & Chr(34), _
                            "#" & aIdPairs(k).sParId & Chr(34), 1, -1, vbTextCompare)
       Next k
     ElseIf lSmilTarget = SMILTARGET_TEXT Then
       For k = 0 To lIdCount - 1
         objOwner.objFileSetHandler.aOutFileSet(i).sDomData = Replace$( _
                            objOwner.objFileSetHandler.aOutFileSet(i).sDomData, _
                            "#" & aIdPairs(k).sParId & Chr(34), _
                            "#" & aIdPairs(k).sTextId & Chr(34), 1, -1, vbTextCompare)
       Next k
     End If
     
    End If
    DoEvents
  Next i
  
  fncAdjustSmilTargetPointers = True
ErrHandler:
  If Not fncAdjustSmilTargetPointers Then objOwner.addlog "<errH in='fncAdjustSmilTargetPointers'>fncAdjustSmilTargetPointers ErrH</errH>"
End Function

Public Function fncSdar2Fix(objOwner As oRegenerator) As Boolean
  'check if the first par in the first smil has two textnodes
  'check if both reference the same ncc node, and that that node is the first in the ncc
  'if so,
  '  a)remove second textnode from par
  '  b)disable the referencing node in ncc

Dim oSmilDom As New MSXML2.DOMDocument40
    oSmilDom.async = False
    oSmilDom.validateOnParse = False
    oSmilDom.resolveExternals = False
    oSmilDom.preserveWhiteSpace = False
    oSmilDom.setProperty "SelectionLanguage", "XPath"
    oSmilDom.setProperty "NewParser", True

Dim oNccDom As New MSXML2.DOMDocument40
    oNccDom.async = False
    oNccDom.validateOnParse = False
    oNccDom.resolveExternals = False
    oNccDom.preserveWhiteSpace = False
    oNccDom.setProperty "SelectionLanguage", "XPath"
    oNccDom.setProperty "NewParser", True

Dim lSmilNum As Long, i As Long
Dim oFirstParNode As IXMLDOMNode
Dim oTextNodes As IXMLDOMNodeList
Dim oNccTargetNode As IXMLDOMNode
Dim sTargetId As String
Dim oDeletedNode As IXMLDOMNode

  fncSdar2Fix = False
    
  'get the first smilfile
  For i = 0 To objOwner.objFileSetHandler.aOutFileSetMembers - 1
    If objOwner.objFileSetHandler.aOutFileSet(i).eType = TYPE_SMIL_1 Then
      lSmilNum = i
      Exit For
    End If
  Next i
  
  If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(lSmilNum).sDomData, oSmilDom, objOwner) Then GoTo ErrHandler
  
  'get the first par
  Set oFirstParNode = oSmilDom.selectSingleNode("//par")
  If Not oFirstParNode Is Nothing Then
    Set oTextNodes = oFirstParNode.selectNodes("text")
    If oTextNodes.length < 2 Then
      fncSdar2Fix = True
      Exit Function
    End If
  Else
    'no par in smilfile, exit
    fncSdar2Fix = True
    Exit Function
  End If
  
  'if we are here, there are dupe text children in first par
  'make sure it is not more than two
  If oTextNodes.length > 2 Then
    objOwner.addlog "<warning>more than two text nodes found in first par of first smil</warning>"
    fncSdar2Fix = True
    Exit Function
  End If
    
  'parse the ncc
  If Not fncParseString(objOwner.objFileSetHandler.aOutFileSet(0).sDomData, oNccDom, objOwner) Then GoTo ErrHandler

  'check if both textnodes reference the same ncc node, and that that node is the first in the ncc
   If oTextNodes.Item(0).selectSingleNode("@src").Text = oTextNodes.Item(1).selectSingleNode("@src").Text Then
     sTargetId = fncGetId(oTextNodes.Item(1).selectSingleNode("@src").Text)
     Set oNccTargetNode = oNccDom.selectSingleNode("//*[@id='" & sTargetId & "']")
     If Not oNccTargetNode Is Nothing Then
       'check that it is the first
       If oNccTargetNode.previousSibling Is Nothing Then
         'it is the first: remove second textnode from par, disable the referencing node in ncc
         'first check that second node in ncc references second text node in first par

         If fncGetId(oNccTargetNode.nextSibling.firstChild.selectSingleNode("@href").Text) = oTextNodes.Item(1).selectSingleNode("@id").Text Then
           'yes, second node in ncc references second text node in first par
           'remove second textnode from par
           objOwner.addlog "<warning>SDAR 2 fix done on" & oTextNodes.Item(1).Text & " and " & oNccTargetNode.nextSibling.Text & "</warning>"
           Set oDeletedNode = oTextNodes.Item(1).parentNode.removeChild(oTextNodes.Item(1))
           'disable the referencing node in ncc
           If Not fncDisableNode(oNccTargetNode.nextSibling, objOwner) Then GoTo ErrHandler
         Else
           fncSdar2Fix = True
           Exit Function
         End If 'fncGetId(oNccTargetNode.nextSibling.f
       Else
         fncSdar2Fix = True
         Exit Function
       End If 'oNccTargetNode.previousSibling Is Nothing
     End If 'Not oNccTargetNode Is Nothing
   Else
    'they are not referecing the same ncc node
    fncSdar2Fix = True
    Exit Function
   End If
  
  'save the changes back to array
  objOwner.objFileSetHandler.aOutFileSet(lSmilNum).sDomData = oSmilDom.xml
  objOwner.objFileSetHandler.aOutFileSet(0).sDomData = oNccDom.xml

  fncSdar2Fix = True
ErrHandler:
  Set oSmilDom = Nothing
  Set oNccDom = Nothing
  If Not fncSdar2Fix Then objOwner.addlog "<errH in='fncSdar2Fix'>fncSdar2Fix ErrH</errH>"
End Function
