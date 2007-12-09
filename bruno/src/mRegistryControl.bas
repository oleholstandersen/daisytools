Attribute VB_Name = "mRegistryControl"
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

' This an API structure for registry key use
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

' This is the HKEY API registry constants
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

' API Error constant
Public Const ERROR_MORE_DATA = 234

' Some Registry API functions

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

' This is the Base key which the fncSaveRegistryData and fncLoadRegistryData defaults
' to if isBaseKey is ommited
Public sBaseKey As String

' This function save a registry key value, the function will default to
' HKEY_CURRENT_USER if ilBaseKey is ommited, the function will default to sBaseKey
' if isBaseKey is ommited.
Public Function fncSaveRegistryData(isKey As String, ivValue As Variant, _
  Optional ilBaseKey As Variant, Optional isBaseKey As Variant) As Boolean

  fncSaveRegistryData = False

  Dim templCounter As Long, lSuccess As Long, lResult As Long, lDisposition As Long
  Dim typsecurity As SECURITY_ATTRIBUTES
  
  With typsecurity
    .nLength = Len(typsecurity)
    .bInheritHandle = True
    .lpSecurityDescriptor = 0
  End With
  
  Dim lKey As Long, sKey As String, lType As Long

' Set default values if optional variables are ommited
  If IsMissing(ilBaseKey) Then lKey = &H80000002 Else lKey = ilBaseKey
  If IsMissing(isBaseKey) Then sKey = sBaseKey Else sKey = isBaseKey

  isKey = sKey & "\" & isKey
  
' Go trough each key in the isKey
  templCounter = InStr(1, isKey, "\", vbBinaryCompare)
  Do Until templCounter = 0

' Try to open the current key, use the last key handle (lResult) as parent key
    lSuccess = RegOpenKeyEx(lKey, Left$(isKey, templCounter - 1), 0, &HF003F, lResult)

' If the key doesn't exist, create it
    If Not lSuccess = 0 Then
      lSuccess = RegCreateKeyEx(lKey, Left$(isKey, templCounter - 1), 0, "None", _
        0, &HF003F, typsecurity, lResult, lDisposition)
      
      If Not lSuccess = 0 Then Exit Function
    Else
      lSuccess = RegCloseKey(lKey)
      lKey = lResult
      isKey = Right$(isKey, Len(isKey) - templCounter)
      templCounter = InStr(1, isKey, "\", vbBinaryCompare)
    End If
  Loop
  
  Dim lData As Long, sData As String
  
' Save the key according to it's type
  If VarType(ivValue) = vbLong Or VarType(ivValue) = vbInteger Then
    lData = ivValue
    lSuccess = RegSetValueEx(lKey, isKey, 0, 4, lData, Len(lData))
  ElseIf VarType(ivValue) = vbString Then
    sData = ivValue
    lSuccess = RegSetValueEx(lKey, isKey, 0, 1, ByVal sData, Len(sData))
  ElseIf VarType(ivValue) = vbBoolean Then
    lData = CLng(ivValue)
    lSuccess = RegSetValueEx(lKey, isKey, 0, 4, lData, Len(lData))
  End If
    
  fncSaveRegistryData = True
End Function

' This function save a registry key value, the function will default to
' HKEY_CURRENT_USER if ilBaseKey is ommited, the function will default to sBaseKey
' if isBaseKey is ommited.
Public Function fncLoadRegistryData(isKey As String, ByRef ivValue As Variant, _
  Optional ilBaseKey As Variant, Optional isBaseKey As Variant, _
  Optional ivDefault As Variant) As Boolean

  fncLoadRegistryData = False

  Dim templCounter As Long, lSuccess As Long, lResult As Long, lDisposition As Long
  Dim typsecurity As SECURITY_ATTRIBUTES
  
  With typsecurity
    .nLength = Len(typsecurity)
    .bInheritHandle = True
    .lpSecurityDescriptor = 0
  End With
  
  Dim lKey As Long, sKey As String, lType As Long
  
' Set default values if optional variables are ommited
  If IsMissing(ilBaseKey) Then lKey = &H80000002 Else lKey = ilBaseKey
  If IsMissing(isBaseKey) Then sKey = sBaseKey Else sKey = isBaseKey
  
  isKey = sKey & "\" & isKey
  
' Go trough each key in the isKey
  templCounter = InStr(1, isKey, "\", vbBinaryCompare)
  Do Until templCounter = 0

' Try to open the current key, use the last key handle (lResult) as parent key
    lSuccess = RegOpenKeyEx(lKey, Left$(isKey, templCounter - 1), 0, &HF003F, lResult)
      
' If the key doesn't exist, set the value to the supplied default value if existing
    If Not lSuccess = 0 Then
      If Not IsMissing(ivDefault) Then ivValue = ivDefault
      Exit Function
    End If
      
    lSuccess = RegCloseKey(lKey)
    lKey = lResult
    isKey = Right$(isKey, Len(isKey) - templCounter)
    templCounter = InStr(1, isKey, "\", vbBinaryCompare)
  Loop
  
  Dim lData As Long, sData As String
  
  lData = 0
' Allocate data so sData can contain all of the data from the key
  sData = String(255, Chr(0))
  
' Load the key according to it's type
  If VarType(ivValue) = vbLong Or VarType(ivValue) = vbInteger Then
TryAgain:
' Try to get the key value
    lSuccess = RegQueryValueEx(lKey, isKey, 0, 4, lData, templCounter)
' if the successcode is ERROR_MORE_DATA, get the value again
    If lSuccess = ERROR_MORE_DATA Then GoTo TryAgain
' If the key doesn't exist, set the value to the supplied default value if existing
    If Not lSuccess = 0 Then
      If Not IsMissing(ivDefault) Then ivValue = ivDefault
      Exit Function
    End If
    ivValue = lData
  ElseIf VarType(ivValue) = vbString Then
TryAgain2:
    lSuccess = RegQueryValueEx(lKey, isKey, 0, 1, ByVal sData, templCounter)
    If lSuccess = ERROR_MORE_DATA Then GoTo TryAgain2
    If Not lSuccess = 0 Then
      If Not IsMissing(ivDefault) Then ivValue = ivDefault
      Exit Function
    End If

' We allocated 255 bytes of data, but the string can be shorter, so search for a 0
' otherwise VB will go haywire when we try to use the variable
    For templCounter = 1 To Len(sData)
      If Asc(Mid$(sData, templCounter, 1)) = 0 Then Exit For
    Next templCounter
    
    If templCounter = 1 Then ivValue = "" Else ivValue = Left$(sData, templCounter - 1)
  ElseIf VarType(ivValue) = vbBoolean Then
TryAgain3:
    lSuccess = RegQueryValueEx(lKey, isKey, 0, 1, lData, templCounter)
    If lSuccess = ERROR_MORE_DATA Then GoTo TryAgain3
    If Not lSuccess = 0 Then
      If Not IsMissing(ivDefault) Then ivValue = ivDefault
      Exit Function
    End If
    ivValue = CBool(lData)
  End If
    
  fncLoadRegistryData = True
End Function
