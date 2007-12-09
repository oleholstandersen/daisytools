VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oJobItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

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



Public sPath As String                    ' Absolute path of job item
Public lDtbType As Long                   ' DTB type
Public lInputCharset As Long              ' Charset
Public lCharsetOther As Long              ' IANA charset
Public bolPreserveMeta As Boolean         ' Preserve meta data in DTB on / off
Public sMetaImport As String              ' Absolute path of file to import meta data from (can make use of path variables)
Public bolSeqRename As Boolean            ' Sequential rename on / off
Public bolUseNumeric As Boolean           ' Use numeric renaming on / off
Public sPrefix As String                  ' Prefix to use while renaming files

Public bolSaveSame As Boolean             ' Replace files in same folder on / off
Public sNewFolder As String               ' Absolute path of new folder (can make use of path variables)

Public bolRegRun As Boolean               ' Indicates wheter a regeneration has been tried or not
Public bolRegResult As Boolean            ' Regeneration succeeded true / false
Public bolRendered As Boolean             ' Indicates wheter the book has been rendered to disk or not
Public bolValResult As Boolean            ' Indicates wheter the book has passed validation without any errors true / false
Public sErrorType As String               ' Can be 'error' or 'warning'
Public sErrorClass As String              ' Can be 'critical' or 'non-critical'
  
Public sRenderedTo As String              ' This variable contains the path of where the job has been rendered to
                                    
Public sID As String                      ' DTB ID, either dc:identifier or ncc:identifier
  
Public objLog As oLogItem                 ' This jobs personal log
  
Public bolMoveBook As Boolean             ' Move book or not?
