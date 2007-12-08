VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInputDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sFullPath As String
Public sFileName As String
Public sDirectory As String
Public sEncoding As String
Public bValid As Boolean
Public bWellformed As Boolean
Public oDom As New MSXML2.DOMDocument40
Public sDocTypeNiceName As String

Private Sub Class_Initialize()
    'debug.Print "cInputDocument.initialize"
    oDom.async = False
    oDom.validateOnParse = False
    oDom.resolveExternals = False
    oDom.preserveWhiteSpace = False
    oDom.setProperty "SelectionLanguage", "XPath"
    'oDom.setProperty "SelectionNamespaces", oBruno.oDriver.sDefaultNameSpace
    oDom.setProperty "NewParser", True
End Sub
