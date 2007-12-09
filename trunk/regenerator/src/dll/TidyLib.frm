VERSION 5.00
Begin VB.Form TidyLib 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "TidyLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private WithEvents tdoc As TidyDocument
Attribute tdoc.VB_VarHelpID = -1
Private sTidyMessage As String
                
Public Function fncRunTidy( _
    ByVal sInPath As String, _
    ByRef sTidied As String, _
    ByVal lInputEncoding As Long, _
    ByVal lFileType As Long, _
    ByVal sOutCharsetName As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim stat As Long
Dim oFSO As Object, oFile As Object
Set oFSO = CreateObject("Scripting.FileSystemObject")
  
  On Error GoTo ErrHandler
  fncRunTidy = False
        
  sTidyMessage = ""
  Set tdoc = New TidyDocument
  
  '20030318 do shiftjis entity fix to ovverride tidy bug: temp set entities to nonentitites
  If (lInputEncoding = CHARSET_SHIFTJIS) Or (lInputEncoding = CHARSET_BIG5) Then
    fncDoCjkEntityFix "pretidy", sInPath, "", objOwner
  End If
  
  'mg 20030525: if asian encoding and smilfile, remove the title meta before tidy parse
'  If (lFileType = TYPE_SMIL_1) And (lInputEncoding = CHARSET_SHIFTJIS Or lInputEncoding = CHARSET_BIG5) Then
'    If Not fncRemoveTitleMeta(sInPath, objOwner) Then GoTo ErrHandler
'  End If
  
  'mg 20030525 per above, but always remove
  If (lFileType = TYPE_SMIL_1) Then
    If Not fncRemoveTitleMeta(sInPath, objOwner) Then GoTo ErrHandler
  End If
   
  'mg 20041005 fix possible malformed ncc/content meta that crashes tidy
  'only do this if preservebibliometa is false
  If objOwner.bPreserveBiblioMeta = False Then
    If (lFileType = TYPE_SMIL_CONTENT) Or (lFileType = TYPE_NCC) Then
      If Not fncRemoveMalformedMeta(sInPath, objOwner) Then GoTo ErrHandler
    End If
  End If
       
  stat = 0
  stat = tdoc.SetErrorFile(sTidyLibPath & "tidyerrs.txt")
  If stat >= 0 Then
    Select Case lFileType
      Case TYPE_NCC, TYPE_SMIL_CONTENT
        Select Case lInputEncoding
          Case CHARSET_WESTERN
            stat = tdoc.LoadConfig(sTidyLibPath & "western_html.tidy")
          Case CHARSET_SHIFTJIS
            stat = tdoc.LoadConfig(sTidyLibPath & "shiftjis_html.tidy")
          Case CHARSET_BIG5
            stat = tdoc.LoadConfig(sTidyLibPath & "big5_html.tidy")
          Case CHARSET_UTF8
            stat = tdoc.LoadConfig(sTidyLibPath & "utf8_html.tidy")
          Case CHARSET_SPECIAL
            'mg 20030325 this select clause to allow for non utf-8 output of western encodings
            'as charset_western always does utf-8 output
            Select Case LCase$(sOutCharsetName)
              Case "windows-1252"
                stat = tdoc.LoadConfig(sTidyLibPath & "1252_html.tidy")
              Case "iso-8859-1"
                stat = tdoc.LoadConfig(sTidyLibPath & "1252_html.tidy")
              Case Else
                stat = tdoc.LoadConfig(sTidyLibPath & "raw_html.tidy")
            End Select
          Case Else
            objOwner.addlog "<error in='fncRunTidy'>unknown encoding in seterrorfile</error>"
            GoTo ErrHandler
        End Select
      Case TYPE_SMIL_1
        Select Case lInputEncoding
          Case CHARSET_WESTERN
            stat = tdoc.LoadConfig(sTidyLibPath & "western_smil.tidy")
          Case CHARSET_SHIFTJIS
            stat = tdoc.LoadConfig(sTidyLibPath & "shiftjis_smil.tidy")
          Case CHARSET_BIG5
            stat = tdoc.LoadConfig(sTidyLibPath & "big5_smil.tidy")
          Case CHARSET_UTF8
            stat = tdoc.LoadConfig(sTidyLibPath & "utf8_smil.tidy")
          Case CHARSET_SPECIAL
            'mg 20030325 this select clause to allow for non utf-8 output of western encodings
            'as charset_western always does utf-8 output
            Select Case LCase$(sOutCharsetName)
              Case "windows-1252", "iso-8859-1"
                stat = tdoc.LoadConfig(sTidyLibPath & "1252_smil.tidy")
              Case "iso-8859-1"
                stat = tdoc.LoadConfig(sTidyLibPath & "1252_smil.tidy")
              Case Else
                stat = tdoc.LoadConfig(sTidyLibPath & "raw_smil.tidy")
            End Select
          Case Else
            objOwner.addlog "<error in='fncRunTidy'>unknown encoding in seterrorfile</error>"
            GoTo ErrHandler
          End Select
    End Select
  End If
        
  If stat >= 0 Then
    stat = tdoc.ParseFile(sInPath)
  End If
  If stat >= 0 Then
    stat = tdoc.CleanAndRepair()
  End If
  If stat >= 0 Then
    stat = tdoc.RunDiagnostics()
  End If
  If stat >= 0 Then
    Select Case lInputEncoding
      Case CHARSET_SHIFTJIS, CHARSET_BIG5
        'go over disc since tidyATL does not return utf16 when asian encodings
        stat = tdoc.SaveFile(sTidyLibPath & "tidied.xml")
        Set oFile = oFSO.opentextfile(sTidyLibPath & "tidied.xml")
        If Not oFile.AtEndOfStream Then 'if the file is not empty
          sTidied = oFile.ReadAll
          oFile.Close
        Else
          GoTo ErrHandler
        End If
      Case Else
        sTidied = tdoc.SaveString()
    End Select
  End If
        
  If sTidyMessage <> "" Then
    If Asc(Mid(sTidyMessage, 1, 1)) = 13 Then sTidyMessage = Mid(sTidyMessage, 2, Len(sTidyMessage))
    If Asc(Mid(sTidyMessage, 1, 1)) = 10 Then sTidyMessage = Mid(sTidyMessage, 2, Len(sTidyMessage))
    objOwner.addlog "<message from='" & "tidyATL" & "' " & "file='" & fncGetFileName(sInPath) & "'>" & sTidyMessage & "</message>"
  End If
    
  '20030318 do shiftjis entity fix to ovverride tidy bug: return to orig value (numeric)
  If (lInputEncoding = CHARSET_SHIFTJIS) Or (lInputEncoding = CHARSET_BIG5) Then
    fncDoCjkEntityFix "posttidy", "", sTidied, objOwner
  End If
      
  'add a new prolog
  sTidied = fncFixProlog(sTidied, lFileType, sOutCharsetName, objOwner)
  
  'remove namespace
  If lFileType = TYPE_NCC Or lFileType = TYPE_SMIL_CONTENT Then
    sTidied = Replace$(sTidied, "xmlns=""http://www.w3.org/1999/xhtml""", "")
    sTidied = Replace$(sTidied, "xmlns='http://www.w3.org/1999/xhtml'", "")
  End If
  
  'mg 20030224; sometimes tidy leaves chars outside </smil>, trunc those
  If lFileType = TYPE_SMIL_1 Then
    'sTidied = Mid$(sTidied, 1, InStrRev(sTidied, "</smil>") + 6)
    'mg20031022: sometimes tidy leaves two </smil> close tags so change the above to:
    'get the whole string until first </smil> occurence
    sTidied = Mid$(sTidied, 1, InStr(sTidied, "</smil>") - 1)
    'append </smil> to the string
    sTidied = sTidied & "</smil>"
  End If
  
  'mg20030423
  'edmar reported this happening on nccs as well
  If lFileType = TYPE_NCC Or lFileType = TYPE_SMIL_CONTENT Then
    'sTidied = Mid$(sTidied, 1, InStrRev(sTidied, "</html>") + 6)
    'get the whole string until first </html> occurence
    sTidied = Mid$(sTidied, 1, InStr(sTidied, "</html>") - 1)
    'append </html> to the string
    sTidied = sTidied & "</html>"
  
  End If
  
  'mg 20030317: added this special fix for shiftjis and tidy
  If (lFileType = TYPE_NCC Or lFileType = TYPE_SMIL_CONTENT) And (lInputEncoding = CHARSET_SHIFTJIS Or lInputEncoding = CHARSET_BIG5) Then
   If Not fncDoCjkBodyTextFix(sTidied, objOwner) Then GoTo ErrHandler
  End If
     
  fncRunTidy = True
  
ErrHandler:
  Set tdoc = Nothing
  Set oFSO = Nothing
  If Not fncRunTidy Then
    Dim sTemp As String
    sTemp = fncGetFileName(sInPath)
    objOwner.addlog "<errH in='fncRunTidy' file='" & sTemp & "'>Tidy processing of " & sTemp & " failed.</errH>"
  End If
End Function
 
Private Sub tdoc_OnMessage( _
    ByVal level As TidyReportLevel, _
    ByVal line As Long, _
    ByVal col As Long, _
    ByVal msg As String)
  Dim lvl As String
  Dim lin As String
  
  If level = TidyInfo Then
  '  lvl = "Info: "
    Exit Sub
  ElseIf level = TidyAccess Then
    lvl = "Access: "
  ElseIf level = TidyWarning Then
    lvl = "Tidy Warning: "
    'Exit Sub
  ElseIf level = TidyConfig Then
    lvl = "Tidy Config: "
  ElseIf level = TidyError Then
    lvl = "Tidy Error: "
  ElseIf level = TidyBadDocument Then
    lvl = "Tidy BadDoc: "
  ElseIf level = TidyFatal Then
    lvl = "Tidy Fatal: "
  Else
    lvl = "Tidy ???: "
  End If
  
  If line > 0 Then
    lin = lvl & "Line " & line & " Col " & col & ", " & msg
  Else
    lin = lvl & msg
  End If
  
  sTidyMessage = sTidyMessage & vbCrLf & Replace(Replace(lin, "<", "&lt;"), ">", "&gt;")
  
End Sub

Private Function fncDoCjkBodyTextFix( _
    ByRef sTidied As String, _
    ByRef objOwner As oRegenerator) _
    As Boolean
Dim oDom As New MSXML2.DOMDocument40
    oDom.async = False
    oDom.validateOnParse = False
    oDom.resolveExternals = False
    oDom.preserveWhiteSpace = False
    oDom.setProperty "SelectionLanguage", "XPath"
    oDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
    oDom.setProperty "NewParser", True
Dim oBodyLastTextNode As IXMLDOMNode
   
  On Error GoTo ErrHandler
  fncDoCjkBodyTextFix = False
  
  'this function removes two garbage bytes that appear as textnode of <body> just before </body>.
  
  If Not fncParseString(sTidied, oDom, objOwner) Then fncDoCjkBodyTextFix = True: Exit Function
  
  Set oBodyLastTextNode = oDom.selectSingleNode("//body/text()[last()]")
  If Not oBodyLastTextNode Is Nothing Then
    If (Len(oBodyLastTextNode.Text) = 2) _
     And Asc(Mid(oBodyLastTextNode.Text, 1, 1)) = 255 _
     And Asc(Mid(oBodyLastTextNode.Text, 2, 1)) = 253 Then
     oBodyLastTextNode.Text = ""
    End If
   'mg 20030528 more brutal deletion:
   'If (Len(oBodyLastTextNode.Text) = 2) Then oBodyLastTextNode.Text = ""
   sTidied = oDom.xml
  End If
    
  fncDoCjkBodyTextFix = True
ErrHandler:
 Set oDom = Nothing
End Function

Private Function fncDoCjkEntityFix( _
  ByVal sPosition As String, _
  ByRef sInPath As String, _
  ByRef sTidied As String, _
  ByRef objOwner As oRegenerator _
  ) As Boolean
Dim oFSO As Object, oFile As Object
Dim sStringToTweak As String
Dim oNode As IXMLDOMNode

  On Error GoTo ErrHandler
  fncDoCjkEntityFix = False
  
  If objOwner.objXhtmlEntities.oEntityDom Is Nothing Or objOwner.objXhtmlEntities.oAllEntities Is Nothing Then GoTo ErrHandler

  Select Case sPosition
    Case "pretidy"
      'read the file from disk
      Set oFSO = CreateObject("Scripting.FileSystemObject")
      Set oFile = oFSO.opentextfile(sInPath)
      If Not oFile.AtEndOfStream Then 'if the file is not empty
        sStringToTweak = oFile.ReadAll
        oFile.Close
      Else
        GoTo ErrHandler
      End If
      'search for all entities and replace with temp value
      For Each oNode In objOwner.objXhtmlEntities.oAllEntities
        sStringToTweak = Replace$(sStringToTweak, "&" & oNode.Text & ";", "regen_" & oNode.Text, , , vbTextCompare)
      Next
      'save the modified file
      If Not fncSaveFile(sTidyLibPath & "shiftjistemp", sStringToTweak, objOwner) Then GoTo ErrHandler
      'set the byref sInPath to this temp storage location
      sInPath = sTidyLibPath & "shiftjistemp"
    Case "posttidy"
      'reset all temp entity values to original
      For Each oNode In objOwner.objXhtmlEntities.oAllEntities
        sTidied = Replace$(sTidied, "regen_" & oNode.Text, "&" & oNode.Text & ";", , , vbTextCompare)
      Next
  End Select

  fncDoCjkEntityFix = True
  
ErrHandler:
  Set oFSO = Nothing: Set oFile = Nothing
  If Not fncDoCjkEntityFix Then objOwner.addlog "<errH in='fncDoCjkEntityFix'>fncDoCjkEntityFix ErrHandler</errH>"
End Function

Private Function fncRemoveTitleMeta( _
  ByRef sInPath As String, _
  ByRef objOwner As oRegenerator _
  ) As Boolean
Dim sStartPos As Long
Dim sEndPos As Long
Dim sXmlString As String
Dim sTemp As String
Dim sTempPath As String

  fncRemoveTitleMeta = False
  
  sXmlString = fncReadFile(sInPath)
  sStartPos = InStr(1, sXmlString, "<meta name=""title""", vbTextCompare)
  If sStartPos = 0 Then
    fncRemoveTitleMeta = True
    Exit Function
  Else
    sEndPos = InStr(sStartPos, sXmlString, "/>", vbBinaryCompare)
    sTemp = Mid(sXmlString, 1, sStartPos - 1) & Mid(sXmlString, sEndPos + 2)
    'Debug.Print sTemp
    sXmlString = sTemp
    'change the inpath so that the orig file is not overwritten
    sTempPath = App.Path: If Right(sTempPath, 1) <> "\" Then sTempPath = sTempPath & "\"
    sInPath = sTempPath & "smiltmp.smil"
    If Not fncSaveFile(sInPath, sXmlString, objOwner) Then GoTo ErrHandler
  End If
  fncRemoveTitleMeta = True
ErrHandler:
  If Not fncRemoveTitleMeta Then objOwner.addlog "<errH in='fncRemoveTitleMeta'>fncRemoveTitleMeta ErrHandler</errH>"
End Function

Private Function fncRemoveMalformedMeta( _
  ByRef sInPath As String, _
  ByRef objOwner As oRegenerator _
  ) As Boolean
Dim sXml As String
Dim sHead As String
Dim lHeadBegin As Long, lHeadEnd As Long
Dim bMalformedTokenFound As Boolean
Dim sTempPath As String
'some nccs (and contentdocs?) contain malformed metas that makes tidy abort; examples:
'<meta name="dc:publisher" content="Albert Bonniers förlag" scheme="Albert Bonniers förlag" name=dc:publishe"/>
'<meta name="dc:identifier" content="91-0-057708-1" scheme="content=91-0-057708-1 name=dc:identifier"/>
'<meta name="ncc:produceddate" content="2002-08-13" scheme="content=2002-08-13 name=ncc:produceddate"/>

'mg20050426: added also for: <meta name="dc:language" content="sv" scheme="ISO 639" content=sv name=dc:languag" />

  fncRemoveMalformedMeta = False
  bMalformedTokenFound = False
  
  sXml = fncReadFile(sInPath)
  'get the content string of head
  lHeadBegin = InStr(1, sXml, "<head>", vbTextCompare) + 6
  lHeadEnd = InStr(1, sXml, "</head>", vbTextCompare)
  sHead = Mid(sXml, lHeadBegin, lHeadEnd - lHeadBegin)
  'Debug.Print "----------"
  'Debug.Print sHead
  'Stop
   
  Dim sTest1 As String, sTest2 As String, sTest3 As String, sTest4 As String, sTest5 As String
  sTest1 = "scheme=" & Chr(34) & "content="
  sTest2 = "name=dc:publishe" & Chr(34) '2 and three tested in combination
  sTest3 = "name=" & Chr(34) & "dc:publisher" & Chr(34)
  sTest4 = "name=dc:languag" & Chr(34) '4 and 5 tested in combination
  sTest5 = "name=" & Chr(34) & "dc:language" & Chr(34)
  'test for the occurence of these malformedness tokens
  If InStr(1, sHead, sTest1, vbTextCompare) > 0 Then bMalformedTokenFound = True
  If ((InStr(1, sHead, sTest2, vbTextCompare) > 0) And (InStr(1, sHead, sTest3, vbTextCompare) > 0)) Then bMalformedTokenFound = True
  If ((InStr(1, sHead, sTest4, vbTextCompare) > 0) And (InStr(1, sHead, sTest5, vbTextCompare) > 0)) Then bMalformedTokenFound = True
        
  If bMalformedTokenFound Then
    objOwner.addlog "<warning in='fncRemoveMalformedMeta'>Malformed metadata found in" & sInPath & ". All metadata is removed from head</warning>"
    'remove the whole meta set
    sXml = Replace$(sXml, sHead, " ")
    'save and mod sInpath for the routine above
    sTempPath = App.Path: If Right(sTempPath, 1) <> "\" Then sTempPath = sTempPath & "\"
    sInPath = sTempPath & "ncctmp.htm"
    If Not fncSaveFile(sInPath, sXml, objOwner) Then GoTo ErrHandler
  End If
  

  fncRemoveMalformedMeta = True
ErrHandler:
  If Not fncRemoveMalformedMeta Then objOwner.addlog "<errH in='fncRemoveMalformedMeta'>fncRemoveMalformedMeta ErrHandler</errH>"
End Function


