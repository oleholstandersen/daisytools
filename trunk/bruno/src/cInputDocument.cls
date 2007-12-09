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
