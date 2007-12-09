Attribute VB_Name = "mGlobal"
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

Public Const TYPE_NCC = 0
Public Const TYPE_SMIL_1 = 1
Public Const TYPE_SMIL_CONTENT = 2
Public Const TYPE_SMIL_AUDIO = 3
Public Const TYPE_SMIL_IMG = 4
Public Const TYPE_SMIL_MASTER = 5
Public Const TYPE_OTHER = 6
Public Const TYPE_SMIL_AUDIO_INSERTED = 7
Public Const TYPE_CSS_INSERTED = 8
Public Const TYPE_DELETED = 9

Public Const CHARSET_WESTERN = 0
Public Const CHARSET_SHIFTJIS = 1
Public Const CHARSET_BIG5 = 2
Public Const CHARSET_UTF8 = 3
Public Const CHARSET_SPECIAL = 4

Public Const RENDER_REPLACE = 0
Public Const RENDER_NEWFOLDER = 1

Public Const DTB_AUDIOONLY = 0
Public Const DTB_AUDIONCC = 1
Public Const DTB_AUDIOPARTTEXT = 2
Public Const DTB_AUDIOFULLTEXT = 3
Public Const DTB_TEXTPARTAUDIO = 4
Public Const DTB_TEXTNCC = 5
  
Public Const POSITION_NEXT = 0
Public Const POSITION_PREC = 1
Public Const POSITION_TOP = 2
  
Public Const SMILTARGET_NOCHANGE = 0
Public Const SMILTARGET_PAR = 1
Public Const SMILTARGET_TEXT = 2
  
Public sAppVersion As String
Public sResourcePath As String
Public sTidyLibPath As String
Public sAudioTemplatesPath As String

Public lInstancesRunning As Long

Public Type eTextSrcData
  sTextSrc As String
  sNewId As String
End Type

Public Type eNoteBodyData
  sTextSrc As String
  sNewId As String
  sOwnerDocName As String
End Type

'mods 20050419 by mg
' ncc:narrator maintained although meta import
'mods 20050417 by mg
' fixed a rename bug reported by brandon - fix in fncDoSequRenameNew

'mods 20050310 by MG
' Change of registry handling (to CURRENT_USER instead of LOCAL_MACHINE),
' based on problem reported by Heiko Becker, DZB
' Note: this means that all preferences and settings will be reset to defaults
' when using this version the first time. Remember to revisit all your
' settings and preferences before running the app.
' Now correctly identifies and brings along files referenced via url()
' statements in CSS.
' Corresponding fix to the ncc:files value calculation
' Removed possibility of ID case inconsistency (could occur when smil
' destination fragment was at <par> instead of <text>, and modify smil
' target set to "no change")
' GUI: fixed error dialog "property settings cancelled" when editing paths
' after having used "set all jobs to these settings"
' Added case correction on class attribute value on first h1 in NCC.
' Fixed bug in related to using path variables when moving books (Guillaume DuBourget)
' Implementors note: DLL Constant typo DTB_AUDIOFULLTTEXT corrected to DTB_AUDIOFULLTEXT

'mods 20041005 by mg
'- added a fix for Victor problem: when the string "id" occurs
'  as the two first characters in the last word of chapter text, Victor may reset.
'  Problem is caused by SMIL title metadata, and the fix done by Regenerator is
'  to insert a space character at the very end of chapter title text in
'  SMIL metadata.
'  Visuaide comments (2004-10-28):
'  The Vibe is unaffected by the problem. For the PRO/Classic/Classic+ a
'  fix is ready and a version will be available in January. A beta will be
'  available for organisations interested in early December. The schedule for
'  VR Soft has yet to be determined.

'- added forced remove of ncc and contentdoc metadata when malformed or erratic as occured in certain early versions of SigtunaDAR and LpStudioPlus (this caused Regenerator to abort).
'Examples:
'<meta name="dc:publisher" content="Albert Bonniers förlag" scheme="Albert Bonniers förlag" name=dc:publishe"/>
'<meta name="dc:identifier" content="91-0-057708-1" scheme="content=91-0-057708-1 name=dc:identifier"/>

'- fixed regenerator abort when render replace mode and write protected master.smil existed in original project

'mods 200402-03 by mg
' improved handling of sigtuna dar2 dtbs
' improved broken link estimation algorithm

'mods 20030913-15 by mg
' disable broken content doc links as well as ncc links
' added option to estimate a new position of broken links
' verbosity in log
' add stylesheet option: put one *.css in resourcepath and it will be included in all docs with no previous stylesheet link
' converts disallowed ncc children to div.group
' set langattrs goes for dc:language first, preexisting if dc:language not available
' added option to make true ncc only (removes content docs)
' move text ids to par (if not impacts skipbooks)
' fix for nonhxref report when target sat on par

'mods 20030911 by mg
'added lcase for ncc and contentdoc known class attr values(fncFixAttrValueCase)
'added wrap:0 in utf8 tidy config file and others

'mods 20030807 by mg
' added hack fncEscapeQuotes that removes " and ' from smil meta attr values

'mods 20030605 by mg
' fix in image copy routine
' added smil img rebuildLinkStructure

'mods 20030605 by mg
' changed the two-garbage-byte removal routine for Shift_Jis after testing on jap windows

'mods 20030330 by mg
' added support for skip-tweaked dtbs (modifying bodyref attrs in fncLinkMangler)

'mods 20030324 by mg
' added wips rename: smil endsynch attr to endsync

'mods 20030324 by mg
' fixed text src prob when intermittent orig ncc/content src pointers
' added http-equiv move to first pos in meta list (ncc and content)
' added global access prop of objregenerator: fncGetDcIdentifier

'mods 20030317 by mg
' added two shiftjis specific tidy fixes
' added xhtml entity collection object (used for shitftjis fix)
' fixed renamefile routine when there were only case diffs

'mods 20030314 by mg
' enus changed to consts
' new tidyATL with msgevents
' TidyATL output to file: now works with shiftjis and utf-8
' new saveroutine using sax and fso to bypass weird msxmldom save problem when shiftjis
' added pb2k layout fix as input option bool to fncRenderFiles

'mods 20030303 by mg
' fixed continue on: oFirstParTextNode Is Nothing in fncFixFirstPar
' fixed rename error on rgn_empty when re-regenerating

'mods 20030227 by mg
' added fncPrintFileSet ErrH
' added catch on hex() (only returns max 8 hexchars)
' added fncFileExists pre fncCheckutf8Encoding

'mods 20030222 by mg
' fixed missing id on audio moved to previous par
' fixed smil dur when audio moved to previous par
'
'mods 20030220 by mg
'  fixed identical text@src values in text@src in fncRebuildLinkStructure
'  scheme in mastersmil meta removed
'  disable invalid ncc links added

'mods 20030110 by MG
'  made a new ncc pretty print routine using sax
'  fixed uri hex escaping (id values) done by tidy - made all uris unresolvable when id value had nonascii char
'  fixed "NCC.html" output from fncRebuildLinkStructure
'  fixed "could not find mp3 file" msg when file was other than mp3
'  made array consistency check abort on fail

' mods 20030109 by MG:
'   fixed et loop when first par of first smil had dupe text events
'   pretty print filenamelog

' mods 20021230 by MG:
'   maxPageNormal printed with content 0 if no pagenormals in ncc
'   removed sSettingsDesc print to log

Sub Main()
    Dim sAppPath As String
    sAppVersion = App.Major & "." & App.Minor & "." & App.Revision

'  ****************************************************************
'  *********************** for dll debug **************************
'  ****************************************************************
'  ***** change the properties project type to active x exe  ******
'  *****       change component starttype to standalone      ******
'  *****          decomment and mod the lines below          ******
'
  sAppPath = App.Path: If Right(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
  sTidyLibPath = "E:\documents\daisyware\regenerator\src\ui_batch\tidylib\"
  sResourcePath = "E:\documents\daisyware\regenerator\src\ui_batch\resources\"
  sAudioTemplatesPath = "E:\documents\daisyware\regenerator\src\ui_batch\audioTemplates\"

  Dim regen As New oRegenerator
  If Not regen.fncRunTransform( _
    "E:\temp\C23077\C23077_1\ncc.html", CHARSET_WESTERN, "windows-1252", DTB_AUDIONCC, _
    False, False, "E:\dtbs\testsuite\meta.xml", _
    True, False, "hauy", True, True, False, 300, 500, 1000, 50, True, _
    True, True, True, False, False, SMILTARGET_NOCHANGE _
     ) Then
         'transform failed
       Stop
      Else
        'transform succeeded
       Stop
        'If Not regen.fncRenderFiles(RENDER_NEWFOLDER, "E:\dtbs\testsuite\two\out\") Then
        If Not regen.fncRenderFiles(RENDER_REPLACE, "") Then
        'rendering failed
       Stop
      Else
        'rendering succeeded
       Stop
      End If
    End If
    regen.fncTerminateObject
    Set regen = Nothing
'
' ****************************************************************
' *********************** end dll debug **************************
' ****************************************************************

End Sub

Public Function fncInstancesRunning() As Boolean
  fncInstancesRunning = False
  If lInstancesRunning > 0 Then fncInstancesRunning = True
End Function
