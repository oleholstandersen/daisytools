VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private sAppPath As String
Private sDtdPath As String
Private sDriverPath As String
Private sShellPath As String
Private sCssPath As String
Private sOutputPath As String
Private sInputPath As String

Public Function InputPath() As String
  InputPath = sInputPath
End Function

Public Function OutputPath() As String
  OutputPath = sOutputPath
End Function

Public Function AppPath() As String
  AppPath = sAppPath
End Function

Public Function DtdPath() As String
  DtdPath = sDtdPath
End Function

Public Function DriverPath() As String
  DriverPath = sDriverPath
End Function

Public Function ShellPath() As String
  ShellPath = sShellPath
End Function

Public Function CssPath() As String
  CssPath = sCssPath
End Function

Public Function setInputPath(isInputPath As String) As Boolean
  sInputPath = isInputPath
End Function

Public Function setOutputPath(isOutputPath As String) As Boolean
  sOutputPath = isOutputPath
End Function

Private Sub Class_Initialize()
  'debug.print "oPaths.initialize"
  sAppPath = App.Path: If Right(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
  sDtdPath = sAppPath & "externals\dtds\"
  sDriverPath = sAppPath & "externals\drivers\"
  sShellPath = sAppPath & "externals\shells\"
  sCssPath = sAppPath & "externals\css\"
End Sub

Private Sub Class_Terminate()
  'debug.print "oPaths.terminate"
End Sub
