VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBruno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public oInputDocuments As cInputDocuments
Public oOutputDocuments As cOutputDocuments
Public oOutputSaver As cOutputSaver
Public oAbstractDocuments As cAbstractDocuments
Public oPaths As cPaths
Public oDriver As cDriver
Public oCmn As cCmn
Public sAppVersion As String
Public lStatus As Long

Private Sub Class_Initialize()
  'debug.Print "cBruno.initialize"
  Set oPaths = New cPaths
  Set oInputDocuments = New cInputDocuments
  Set oCmn = New cCmn
  sAppVersion = "bruno " & App.Major & "." & App.Minor & "." & App.Revision
  Me.fncSetStatus (STATUS_IDLE)
End Sub

Public Function fncCreateDriver(sDriverFullPath As String) As Boolean
    If Not oDriver Is Nothing Then Set oDriver = Nothing
    Set oDriver = New cDriver
    If oDriver.fncInstantiate(sDriverFullPath) Then fncCreateDriver = True
End Function

Public Function fncCreateAbstractDocuments() As Boolean
  Set oAbstractDocuments = New cAbstractDocuments
  If oAbstractDocuments.fncInstantiate Then
    fncCreateAbstractDocuments = True
    'kill the input doc arrays to save memory
    'oBruno.oInputDocuments.fncResetArrays
    'no dont because it is used when rendering auxifiles in outputsaver
    oBruno.oInputDocuments.fncResetDomObjectsInArrays
    
  Else
    'oAbstractDocuments.Instantiate failed
    fncCreateAbstractDocuments = False
  End If
End Function

Public Function fncAddInputDocument(sFullPath As String) As Boolean
  fncAddInputDocument = False
  If oInputDocuments.fncAddDocument(sFullPath) Then fncAddInputDocument = True
End Function

Public Function fncCreateOutputDocuments() As Boolean
  Set oOutputDocuments = New cOutputDocuments
  If oOutputDocuments.fncInstantiate() Then
    fncCreateOutputDocuments = True
  Else
    'oOutputDocuments.Instantiate failed
    fncCreateOutputDocuments = False
  End If
End Function

Public Function fncRenderOutputDocuments(sOutputPath As String) As Boolean
  oPaths.setOutputPath sOutputPath
  Set oOutputSaver = New cOutputSaver
  If oOutputSaver.fncRenderOutputDocuments Then fncRenderOutputDocuments = True
End Function

Public Function fncTerminateChildren() As Boolean
Dim bNothing As Boolean
  fncTerminateChildren = True
  bNothing = True
  If Not oInputDocuments Is Nothing Then
    Set oInputDocuments.oInputMetadata = Nothing: If Not oInputDocuments.oInputMetadata Is Nothing Then bNothing = False
    Set oInputDocuments = Nothing: If Not oInputDocuments Is Nothing Then bNothing = False
    Set oAbstractDocuments = Nothing: If Not oAbstractDocuments Is Nothing Then bNothing = False
    Set oOutputDocuments = Nothing: If Not oOutputDocuments Is Nothing Then bNothing = False
    Set oDriver = Nothing: If Not oDriver Is Nothing Then bNothing = False
    Set oPaths = Nothing: If Not oPaths Is Nothing Then bNothing = False
    Set oOutputSaver = Nothing: If Not oOutputSaver Is Nothing Then bNothing = False
    If Not bNothing Then fncTerminateChildren = False
  End If
End Function

Public Function fncSetStatus(lInStatus As Long)
'Public Const STATUS_IDLE = 0
'Public Const STATUS_WORKING = 1
'Public Const STATUS_ABORTED = 2
'Public Const STATUS_DONE = 3
'Public Const STATUS_UNKNOWN = 4

  Select Case lInStatus
    Case STATUS_IDLE
      Me.lStatus = STATUS_IDLE
    Case STATUS_WORKING
      Me.lStatus = STATUS_WORKING
    Case STATUS_ABORTED
      Me.lStatus = STATUS_ABORTED
    Case STATUS_DONE
      Me.lStatus = STATUS_DONE
    Case Else
      Me.lStatus = STATUS_UNKNOWN
  End Select


End Function

Public Function fncGetStatus() As Long
  fncGetStatus = Me.lStatus
End Function

Private Sub Class_Terminate()
  'debug.Print "cBruno.terminate"
End Sub
