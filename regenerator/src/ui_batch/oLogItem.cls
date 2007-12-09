VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oLogItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Daisy 2.02 Regenerator Batch UI
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

Private sprivLog As String
Public bolCached As Boolean
Public sCacheFile As String
Public objFile As Object

Public Property Let sLog(isLog As String)
  sprivLog = isLog
  If Len(sprivLog) >= 32000 Then subSaveCache
End Property

Public Property Get sLog() As String
  If bolCached Then
    Set objFile = Nothing
    Set objFile = oFSO.opentextfile(sCacheFile, 1, True)
    sLog = objFile.readall & sprivLog
    Set objFile = Nothing
    Set objFile = oFSO.opentextfile(sCacheFile, 8, True)
  Else
    sLog = sprivLog
  End If
End Property

Public Function fncAddToLog(isLog As String)
  sLog = sprivLog & isLog & vbCrLf
End Function

Private Sub subSaveCache()
  If sCacheFile = "" Then
    Do
      sCacheFile = sTempPath & Format(Int(Rnd * 10000000), "00000000") & ".chc"
    Loop Until Not oFSO.fileexists(sCacheFile)
    Set objFile = oFSO.opentextfile(sCacheFile, 8, True)
  End If
  
  objFile.write (sprivLog)
  
  sprivLog = ""
  bolCached = True
End Sub

Private Sub Class_Terminate()
  If Not bolCached Then Exit Sub
  objFile.Close
  oFSO.deletefile sCacheFile
End Sub
