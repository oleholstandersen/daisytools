VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oUserControl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

' *** User Control ***
' This object lets the user set and get properties used by the validator
'
' xobjEvent = Set to the dll event object
' get/set TempPath = Get / Set the path where the validator can create temporar
'   files
' get/set DTDPath = Get / Set the path where the DTDs needed by the
'   validator are found
' get/set ADTDPath = Get / Set the path where the ADTDs needed by the
'   validator are found
' get/set VTMPath = Get / Set the path where the VTM file needed by the validator
'   is found
' fncGetProgress = Returns the current validations progress in percent
' get/set CancelFlag = Get / Set the cancel flag for the validator to halt the
'   current validation. This flag must be resetted before trying to validate again.

Public xobjEvent As oEvent

'Private privDtbDcIdentifier As String

Public Function fncSetLightMode(ibolLightMode As Boolean) As Boolean
  bolLightMode = ibolLightMode
  fncSetLightMode = True
End Function

Public Function fncSetDisableAudioTests(ibolDisableAudioTests As Boolean) As Boolean
  bolDisableAudioTests = ibolDisableAudioTests
  fncSetDisableAudioTests = True
End Function

Public Property Get propLightMode() As Boolean
  propLightMode = bolLightMode
End Property

Public Property Get propDisableAudioTests() As Boolean
  propDisableAudioTests = bolDisableAudioTests
End Property

Public Property Get propTempPath() As String
  propTempPath = sTempPath
End Property

Public Function fncSetTempPath(isTempPath As String) As Boolean
  sTempPath = fncGetPathName(isTempPath)
  fncSetTempPath = True
End Function

Public Property Get propDTDPath() As String
  propDTDPath = sDtdPath
End Property

Public Function fncSetDTDPath(ByVal isDTDPath As String) As Boolean
  sDtdPath = fncGetPathName(isDTDPath)
  fncSetDTDPath = True
End Function

Public Property Get propAdtdPath() As String
  propAdtdPath = sAdtdPath
End Property

Public Function fncSetAdtdPath(isAdtdPath As String) As Boolean
    sAdtdPath = fncGetPathName(isAdtdPath)
    fncSetAdtdPath = True
End Function
    
Public Property Get propVtmPath() As String
  propVtmPath = sVtmPath
End Property

Public Function fncSetVtmPath(isVtmPath As String) As Boolean
    sVtmPath = fncGetPathName(isVtmPath)
    fncSetVtmPath = True
End Function

Public Function fncGetProgress() As Byte
  fncGetProgress = lTotalProgress
End Function

Public Function fncGetDllVersion() As String
  fncGetDllVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Property Get propTimeSpan() As Long
  propTimeSpan = lTimeSpan
End Property

Public Function fncSetCancelFlag() As Boolean
  bolCancelValidation = True
End Function

Public Function fncClearCancelFlag() As Boolean
  bolCancelValidation = False
End Function

Public Function fncGetCancelFlag() As Boolean
  fncGetCancelFlag = bolCancelValidation
End Function

Public Function fncSetTimeSpan(ilSpanMs As Long) As Long
  lTimeSpan = ilSpanMs
End Function

Public Function fncSetAdvancedADTD(ibolAdvancedADTD As Boolean)
  bolAdvancedADTD = ibolAdvancedADTD
End Function

Public Function fncGetAdvancedADTD() As Boolean
  fncGetAdvancedADTD = bolAdvancedADTD
End Function

' This function initializes the validator. This function doesn't have to be used,
' but can be used if a developer wants more controll of the Validator Engine
' behaviour.
Public Function fncInitializeValidator() As Boolean
  On Error Resume Next
  
  Dim objTestDom As Object
  Set objTestDom = CreateObject("Msxml2.DOMDocument.4.0")
  If objTestDom Is Nothing Then
    objEvent.subLog "Msxml 4 not installed or not correctly registered!"
    Exit Function
  End If
  Set objTestDom = Nothing
  
  If Not fncLoadRulesFiles Then Exit Function
  If Not fncParseVtm Then Exit Function
  
  fncInitializeValidator = True
End Function

' This function deinitalizes the validator. This function doesn't have to be used,
' but might be usable due to the unreliability of termination of VB classes.
Public Function fncDeinitializeValidator() As Boolean
  Dim lCounter As Long
  
  ReDim Preserve aProgressItem(0)
  lProgressCount = 0
  
  For lCounter = 0 To lVtmItemsCount - 1
    Set aVtmItems(lCounter) = Nothing
  Next lCounter
  
  ReDim aVtmItems(0)
  lVtmItemsCount = 0
  bolVtmIsLoaded = False
  
  Set objRulesSmil = Nothing
  Set objRulesNcc = Nothing
  Set objRulesContent = Nothing
  Set objRulesMasterSmil = Nothing
  Set objRulesDiscinfo = Nothing
  Set objRulesMultivolume = Nothing
  
  bolRulesAreLoaded = False
  
  Set xobjEvent = Nothing
  
  fncDeinitializeValidator = True
End Function

Private Sub Class_Initialize()
    Set xobjEvent = objEvent
End Sub
