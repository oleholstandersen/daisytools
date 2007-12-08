VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCmn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Cmn is short for Common

Public oFsoCmn As cFsoCmn
Public oDomCmn As cDomCmn
Public oUriParser As cUriParser

Private Sub Class_Initialize()
 'debug.Print "cCmn.Initialize"
 Set oFsoCmn = New cFsoCmn
 Set oDomCmn = New cDomCmn
 Set oUriParser = New cUriParser
End Sub

Private Sub Class_Terminate()
 'debug.Print "cCmn.Terminate"
 Set oFsoCmn = Nothing
 Set oDomCmn = Nothing
 Set oUriParser = Nothing
End Sub
