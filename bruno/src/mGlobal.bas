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
 
Public oBruno As cBruno
Public oDriverList As cDriverList
Public oUi As Form
Public oRegistryControl As cRegistryControl

Public Const BIF_RETURNONLYFSDIRS   As Long = &H1
Public Const BIF_DONTGOBELOWDOMAIN  As Long = &H2
Public Const BIF_VALIDATE           As Long = &H20
Public Const BIF_EDITBOX            As Long = &H10
Public Const BIF_NEWDIALOGSTYLE     As Long = &H40
Public Const BIF_USENEWUI As Long = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)

Public Const DIRECTION_UP = 0
Public Const DIRECTION_DOWN = 1

Public Const OUTPUT_TYPE_D202 = 0
Public Const OUTPUT_TYPE_Z39 = 1

Public Const DOCTYPE_UNKNOWN = 0
Public Const DOCTYPE_XHTML1_TRANSITIONAL = 1
Public Const DOCTYPE_XHTML1_STRICT = 2
Public Const DOCTYPE_XHTML11_TRANSITIONAL = 3
Public Const DOCTYPE_XHTML11_STRICT = 4
Public Const DOCTYPE_DTBOOK = 5

Public Const TYPE_ABSTRACT_CONTENTDOC = 0
Public Const TYPE_ABSTRACT_SMIL = 1
Public Const TYPE_ABSTRACT_NAVIGATION = 2

Public Const TYPE_ACTUAL_202NCC = 0
Public Const TYPE_ACTUAL_202SMIL = 1
Public Const TYPE_ACTUAL_202MSMIL = 2
Public Const TYPE_ACTUAL_202CONTENT = 3
Public Const TYPE_ACTUAL_Z39NCX = 4
Public Const TYPE_ACTUAL_Z39OPF = 5
Public Const TYPE_ACTUAL_Z39CONTENT = 6
Public Const TYPE_ACTUAL_Z39SMIL = 7
Public Const TYPE_ACTUAL_Z39RESOURCE = 8
Public Const TYPE_ACTUAL_W3CSMIL = 9
Public Const TYPE_ACTUAL_AUXILLIARY = 10  'images, stylesheets, etc
Public Const TYPE_ACTUAL_LPP = 11  '*.lpp file for lpPro
Public Const TYPE_ACTUAL_MDF = 12  '*.mdf file for lpPro

Public Const RELATION_CHILD = 100
Public Const RELATION_DESCENDANT = 101
Public Const RELATION_SIBLING = 102
Public Const RELATION_PARENT = 103
Public Const RELATION_ANCESTOR = 104
Public Const RELATION_SELF = 105
Public Const RELATION_UNKNOWN = 106

Public Const STATUS_IDLE = 0
Public Const STATUS_WORKING = 1
Public Const STATUS_ABORTED = 2
Public Const STATUS_DONE = 3
Public Const STATUS_UNKNOWN = 4

'**********************************************
'**************** change list *****************'
'20040830
'  first public beta
'20041029
'  fix: whitespace truncation on inlines (reported by Jesper Klein)
'  fix: problem in QAPlayer caused by first two headings being in same smilfile (reported by David Gordon and Stefan Kropf)
'  fix: problem in content doc display in LpPro when head/title was empty or whitespace only (reported by David Gordon)
'  fix: Lpp creation error when several input documents pointed to auxilliary files in different folders but with same filename
'  feature: outputpath and driver selection stored between sessions (suggested by Per Sennels)
'20041222
'  fix: runtime error when no namespace declaration present (Miki Azuma)
'  feature: support of url() statements in CSS for inclusion of auxilliary files (DFA ITT)
'  feature: working status indication in caption (Niels Thogersen)
'  fix: handling of spaces in input and output directory paths (Sean Brooks)
'  feature/fix: validation of XHTML now also checks for first body element being an h1 class title (Sean Brooks)
'  fix: keeping infobar message history
'20050310
'  (Danny CNIB) fixed pretty print routine that removed alphanumeric entities in XHTML
'  (Niels IBOS) added import into lpp and ncc of following meta items in xhtml sourcedoc: dc:source, ncc:multimediaType, ncc:narrator, ncc:sourceDate, ncc:sourceEdition, ncc:sourcePublisher
'  (Sean CNIB) fixed SQL error when metadata in source document contained single quotes
'20070119
'  feature: default XHTML driver supporting representation in NCC of full set of skippable elements
'  fix/feature: occurences of div.prodnode and div.sidebar in content document turned into span in ncc
'  fix: mixed content nodes no longer loose text in ncc (ie "<h1> my <em>inline</em> text</h1>" before fix became "<h1> my  text</h1>" in ncc)
'  fix: bodyref attrs removed from ncc
'20070308
'  various fixes to the Z39.86-2005 rendering
'20070910
'  various fixes to the mixed content and sync-omit handling (see cAbstractDocuments.fncOmitOmit)


'**********************************************
'**********************************************


'**********************************************
'**************** to do area ******************
'if a navinclude elem doesnt become a syncpoint
'(example: dtbook:poem with no text, only list children



'**********************************************
'**********************************************


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
   ByVal lpParameters As String, ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

'Private Declare Function Beep Lib "kernel32" _
'(ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Sub Main()
 Dim bTesting As Boolean
 
 On Error GoTo errhandler
 
 'test availability of base objects
 bTesting = True
 
 Dim oTemp As Object
 Dim bTestMSXML4 As Boolean
 bTestMSXML4 = True
 Set oTemp = CreateObject("Msxml2.DOMDocument.4.0")
 If oTemp Is Nothing Then
    MsgBox "MSXML 4 not available: exiting", vbOKOnly
    Exit Sub
 End If
 bTestMSXML4 = False
 
 bTesting = False
 'end test avail
 
 'set the registry control object
 Set oRegistryControl = New cRegistryControl
 oRegistryControl.sBaseKey = "Software\DaisyWare\bruno"
 '... the settings are loaded in frmMain
 
 
 'set up main objects
 Set oBruno = New cBruno
 Set oUi = frmMain
 Set oDriverList = New cDriverList
 subApplyUserLcid
 frmMain.Caption = oBruno.sAppVersion
 frmMain.Show
 frmMain.cmbDriverList.SetFocus
 
 
Exit Sub
errhandler:
 If bTesting Then
  If bTestMSXML4 Then MsgBox "MSXML 4 not available: exiting", vbOKOnly
 End If
 
End Sub


