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
