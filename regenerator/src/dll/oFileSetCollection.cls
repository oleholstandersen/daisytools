VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oFileSetCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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


Option Explicit

Private eprivType As Long    ' enuFileType
Private sprivDomData As String
Private sprivFileName As String
Private sprivSmilTitle As String
Private lprivOwnerType As Long
Private sprivAbsPath As String

Public Property Get eType() As Long
  eType = eprivType
End Property

Friend Property Let eType(ieType As Long)
  eprivType = ieType
End Property

Public Property Get sDomData() As String
  sDomData = sprivDomData
End Property

Friend Property Let sDomData(isDomData As String)
  sprivDomData = isDomData
End Property

Public Property Get sFileName() As String
  sFileName = sprivFileName
End Property

Friend Property Let sFileName(isFileName As String)
  sprivFileName = isFileName
End Property

Public Property Get sSmilTitle() As String
  sSmilTitle = sprivSmilTitle
End Property

Friend Property Let sSmilTitle(isSmilTitle As String)
  sprivSmilTitle = isSmilTitle
End Property

Public Property Get lOwnerType() As Long
  lOwnerType = lprivOwnerType
End Property

Friend Property Let lOwnerType(ilOwnerType As Long)
  lprivOwnerType = ilOwnerType
End Property

Public Property Get sAbsPath() As String
  sAbsPath = sprivAbsPath
End Property

Friend Property Let sAbsPath(isAbsPath As String)
  sprivAbsPath = isAbsPath
End Property
